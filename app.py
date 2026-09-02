import io
import json
import os
import tempfile
import time
import fitz  # PyMuPDF
from google import genai
from google.genai import types
import openpyxl
from openpyxl.drawing.image import Image as OpenPyXlImage
import pandas as pd
from PIL import Image as PILImage
from pydantic import BaseModel, Field
import streamlit as st

# Set Streamlit Page Configuration
st.set_page_config(
    page_title="Gemini Catalog Extractor",
    page_icon="📦",
    layout="wide",
)

st.title("📦 Gemini Catalog Extractor (Excel with Images)")
st.write(
    "Upload any product catalog PDF to extract specifications and generate an Excel schedule with embedded page thumbnails."
)

# -----------------------------------------------------------------------------
# API Client Setup
# -----------------------------------------------------------------------------
api_key = os.getenv("GEMINI_API_KEY")

if not api_key:
    try:
        if "GEMINI_API_KEY" in st.secrets:
            api_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        api_key = None

if not api_key:
    api_key = st.text_input("Enter Google Gemini API Key:", type="password")

if not api_key:
    st.error(
        "🔑 Google Gemini API Key not found. Please set `GEMINI_API_KEY` in Streamlit Secrets."
    )
    st.stop()

client = genai.Client(api_key=api_key)


# -----------------------------------------------------------------------------
# Pydantic Schemas
# -----------------------------------------------------------------------------
class ProductSpec(BaseModel):
    page_number: int = Field(
        description="The 1-based page number of the catalog where this product appears"
    )
    category: str = Field(
        description="e.g., Task Chair, Conference Table, Executive Desk"
    )
    model_number: str = Field(description="Model or SKU code if visible")
    length_mm: str = Field(description="Length dimension or 'N/A'")
    width_mm: str = Field(description="Width/Depth dimension or 'N/A'")
    height_mm: str = Field(description="Height dimension or 'N/A'")
    primary_materials: str = Field(description="Materials mentioned or visible")
    color_finish: str = Field(
        description="Color or surface finish description"
    )
    key_features: str = Field(
        description="Brief summary of notable design features"
    )


class CatalogExtraction(BaseModel):
    products: list[ProductSpec] = Field(
        description="List of all extracted product specifications"
    )


# -----------------------------------------------------------------------------
# Helper Functions
# -----------------------------------------------------------------------------
def save_uploaded_file_to_disk(uploaded_file):
    """Saves uploaded file chunks directly to a temporary file on disk."""
    temp_dir = tempfile.gettempdir()
    file_path = os.path.join(temp_dir, uploaded_file.name)

    with open(file_path, "wb") as f:
        # Write in 4MB chunks to prevent memory spikes
        while chunk := uploaded_file.read(4 * 1024 * 1024):
            f.write(chunk)

    return file_path


def extract_single_page_pdf(doc, page_num_1_based):
    """Extracts a single page from a PyMuPDF doc and returns raw PDF bytes."""
    new_doc = fitz.open()
    new_doc.insert_pdf(
        doc, from_page=page_num_1_based - 1, to_page=page_num_1_based - 1
    )
    output_stream = io.BytesIO()
    new_doc.save(output_stream)
    single_bytes = output_stream.getvalue()
    new_doc.close()
    return single_bytes


def render_page_thumbnail(doc, page_num_1_based, max_size=(120, 120)):
    """Render a specific PDF page into a PIL Image thumbnail."""
    page_idx = page_num_1_based - 1
    if page_idx < 0 or page_idx >= len(doc):
        return None

    page = doc[page_idx]
    pix = page.get_pixmap(dpi=100)
    img = PILImage.open(io.BytesIO(pix.tobytes("png")))
    img.thumbnail(max_size)
    return img


def create_excel_with_images(df, doc):
    """Generates an Excel workbook with embedded page thumbnails."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Product Schedule"

    headers = ["Thumbnail"] + list(df.columns)
    ws.append(headers)

    ws.column_dimensions["A"].width = 18

    for idx, row in df.iterrows():
        excel_row = idx + 2
        ws.row_dimensions[excel_row].height = 90

        for col_idx, value in enumerate(row, start=2):
            ws.cell(row=excel_row, column=col_idx, value=str(value))

        actual_page_num = int(row.get("page_number", 1))
        pil_img = render_page_thumbnail(doc, actual_page_num)

        if pil_img:
            img_io = io.BytesIO()
            pil_img.save(img_io, format="PNG")
            img_io.seek(0)

            img_obj = OpenPyXlImage(img_io)
            cell_address = f"A{excel_row}"
            ws.add_image(img_obj, cell_address)

    output_stream = io.BytesIO()
    wb.save(output_stream)
    return output_stream.getvalue()


def process_single_page_with_retry(single_pdf_bytes, page_num, max_retries=3):
    """Processes a single page through Gemini API with fallback retry."""
    models_to_try = ["gemini-3.6-flash", "gemini-2.5-flash", "gemini-1.5-flash"]
    prompt = f"""
    You are analyzing Page {page_num} of a product catalog.
    Extract every product item present on this page along with all specification details.
    Always set `page_number` to {page_num}.
    If this page contains no technical specifications or product details (e.g. cover page, blank page, pure marketing photo), return an empty list for `products`.
    """

    for model_name in models_to_try:
        for attempt in range(max_retries):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=[
                        types.Part.from_bytes(
                            data=single_pdf_bytes, mime_type="application/pdf"
                        ),
                        prompt,
                    ],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=CatalogExtraction,
                        temperature=0.1,
                    ),
                )
                parsed = json.loads(response.text)
                return parsed.get("products", [])
            except Exception as e:
                err_str = str(e)
                if "503" in err_str or "UNAVAILABLE" in err_str:
                    time.sleep((attempt + 1) * 3)
                else:
                    break
    return []


# -----------------------------------------------------------------------------
# Main Application UI
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Choose a catalog PDF file (supports large files up to 500MB)", type=["pdf"]
)

if uploaded_file:
    with st.spinner("Loading PDF file..."):
        # Save chunked stream to temp path
        temp_pdf_path = save_uploaded_file_to_disk(uploaded_file)
        doc = fitz.open(temp_pdf_path)
        total_pages = len(doc)

    st.sidebar.header("📄 Page Processing Options")
    process_mode = st.sidebar.radio(
        "Select Range:", ["Process Sample Range", "Process Full Document"]
    )

    if process_mode == "Process Sample Range":
        start_page = st.sidebar.number_input(
            "Start Page", min_value=1, max_value=total_pages, value=1
        )
        default_end = min(start_page + 3, total_pages)
        end_page = st.sidebar.number_input(
            "End Page",
            min_value=start_page,
            max_value=total_pages,
            value=default_end,
        )
    else:
        start_page = 1
        end_page = total_pages

    if st.button("Extract Specifications & Generate Excel", type="primary"):
        all_extracted_products = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        pages_to_process = list(range(start_page, end_page + 1))

        for i, current_page in enumerate(pages_to_process):
            status_text.text(
                f"Processing catalog page {current_page} of {end_page}..."
            )
            single_bytes = extract_single_page_pdf(doc, current_page)
            page_products = process_single_page_with_retry(
                single_bytes, current_page
            )
            all_extracted_products.extend(page_products)
            progress_bar.progress((i + 1) / len(pages_to_process))

        status_text.empty()
        progress_bar.empty()

        if all_extracted_products:
            df = pd.DataFrame(all_extracted_products)
            st.success(
                f"Extraction complete! Found {len(df)} product item(s) across pages {start_page} to {end_page}."
            )

            st.subheader("Extracted Specifications Preview")
            st.dataframe(df, use_container_width=True)

            excel_bytes = create_excel_with_images(df, doc)

            st.download_button(
                label="📥 Download Excel Schedule with Images (.xlsx)",
                data=excel_bytes,
                file_name=f"{os.path.splitext(uploaded_file.name)[0]}_schedule.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning(
                "No structured product specifications were found in the selected range."
            )

        # Cleanup temp file on completion
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
