import base64
import gc
import io
import os
import tempfile
import time
import fitz  # PyMuPDF
from openai import OpenAI
import openpyxl
from openpyxl.drawing.image import Image as OpenPyXlImage
import pandas as pd
from PIL import Image as PILImage
from pydantic import BaseModel, Field
import streamlit as st

# -----------------------------------------------------------------------------
# Streamlit Page Configuration
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="OpenAI Catalog Extractor",
    page_icon="📦",
    layout="wide",
)

st.title("📦 Catalog Extractor (Excel with Images)")
st.write(
    "Upload any product catalog PDF to extract specifications and generate an Excel schedule with embedded page thumbnails."
)

# -----------------------------------------------------------------------------
# API Key Resolution & Client Initialization
# -----------------------------------------------------------------------------
api_key = None

if "OPENAI_API_KEY" in st.secrets and st.secrets["OPENAI_API_KEY"]:
    api_key = st.secrets["OPENAI_API_KEY"].strip()
elif os.getenv("OPENAI_API_KEY"):
    api_key = os.getenv("OPENAI_API_KEY").strip()

if not api_key:
    api_key = st.sidebar.text_input("Enter OpenAI API Key:", type="password")

if not api_key:
    st.error(
        "🔑 OpenAI API Key not found. Please set `OPENAI_API_KEY` in Streamlit Secrets."
    )
    st.stop()

# Initialize OpenAI Client
client = OpenAI(api_key=api_key)


# -----------------------------------------------------------------------------
# Pydantic Schemas for Structured Output
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
# Helper Functions (Memory & Streaming Optimized)
# -----------------------------------------------------------------------------
def save_uploaded_file_to_disk(uploaded_file):
    """Saves uploaded file chunks directly to a temporary file on disk."""
    temp_dir = tempfile.gettempdir()
    file_path = os.path.join(temp_dir, uploaded_file.name)

    with open(file_path, "wb") as f:
        while chunk := uploaded_file.read(4 * 1024 * 1024):
            f.write(chunk)

    return file_path


def extract_page_as_base64_jpeg(doc, page_num_1_based):
    """Renders a single PDF page into a JPEG image encoded in base64 for OpenAI Vision."""
    page_idx = page_num_1_based - 1
    page = doc[page_idx]
    pix = page.get_pixmap(dpi=150)

    img = PILImage.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=85)

    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def render_page_thumbnail(doc, page_num_1_based, max_size=(100, 100)):
    """Render thumbnail using low DPI and JPEG compression for Excel insertion."""
    page_idx = page_num_1_based - 1
    if page_idx < 0 or page_idx >= len(doc):
        return None

    page = doc[page_idx]
    pix = page.get_pixmap(dpi=72)
    img = PILImage.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    img.thumbnail(max_size)
    return img


def create_excel_with_images(df, doc):
    """Generates Excel schedule with embedded thumbnails."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Product Schedule"

    headers = ["Thumbnail"] + list(df.columns)
    ws.append(headers)

    ws.column_dimensions["A"].width = 16

    for idx, row in df.iterrows():
        excel_row = idx + 2
        ws.row_dimensions[excel_row].height = 80

        for col_idx, value in enumerate(row, start=2):
            ws.cell(row=excel_row, column=col_idx, value=str(value))

        actual_page_num = int(row.get("page_number", 1))
        pil_img = render_page_thumbnail(doc, actual_page_num)

        if pil_img:
            img_io = io.BytesIO()
            pil_img.save(img_io, format="JPEG", quality=70)
            img_io.seek(0)

            img_obj = OpenPyXlImage(img_io)
            ws.add_image(img_obj, f"A{excel_row}")

    output_stream = io.BytesIO()
    wb.save(output_stream)
    return output_stream.getvalue()


def process_single_page_with_openai(doc, page_num, max_retries=3):
    """Processes a single page through OpenAI's gpt-4o-mini Vision with Structured Outputs."""
    base64_image = extract_page_as_base64_jpeg(doc, page_num)

    prompt = f"""
    You are analyzing Page {page_num} of a commercial furniture/interior product catalog.
    
    Task:
    1. Extract every furniture, fixture, or equipment item listed on this page.
    2. TRANSLATE ALL EXTRACTED TEXT (product names, category descriptions, materials, finishes, and key features) INTO ENGLISH.
    3. Output all values in English for: category, model_number, dimensions (length_mm, width_mm, height_mm), primary_materials, color_finish, and key_features.
    4. Set `page_number` to {page_num} for every item.
    5. If details like dimensions or SKU are missing, set them to "N/A".
    6. Only return an empty list if the page has zero products.
    """

    last_error = None

    for attempt in range(max_retries):
        try:
            completion = client.beta.chat.completions.parse(
                model="gpt-4o-mini",
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/jpeg;base64,{base64_image}"
                                },
                            },
                        ],
                    }
                ],
                response_format=CatalogExtraction,
            )
            parsed_data = completion.choices[0].message.parsed
            return parsed_data.products
        except Exception as e:
            last_error = str(e)
            if "429" in last_error or "rate_limit" in last_error:
                time.sleep((attempt + 1) * 3)
            else:
                break

    st.sidebar.warning(f"⚠️ Page {page_num}: {last_error}")
    return []


# -----------------------------------------------------------------------------
# Main Application UI
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Choose a catalog PDF file (supports large files up to 500MB)", type=["pdf"]
)

if uploaded_file:
    with st.spinner("Loading PDF file..."):
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
        default_end = min(start_page + 5, total_pages)
        end_page = st.sidebar.number_input(
            "End Page",
            min_value=start_page,
            max_value=total_pages,
            value=default_end,
        )
    else:
        start_page = 1
        end_page = total_pages

    st.info(f"Ready to process pages **{start_page} to {end_page}** (Total: {total_pages} pages in document).")

    if st.button("Extract Specifications & Generate Excel", type="primary"):
        all_extracted_products = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        pages_to_process = list(range(start_page, end_page + 1))

        for i, current_page in enumerate(pages_to_process):
            status_text.text(f"Processing page {current_page} of {end_page}...")
            page_products = process_single_page_with_openai(doc, current_page)

            if page_products:
                # Convert Pydantic objects to dicts
                products_dict = [p.model_dump() for p in page_products]
                st.sidebar.write(f"✅ Page {current_page}: Found {len(products_dict)} item(s)")
                all_extracted_products.extend(products_dict)

            progress_bar.progress((i + 1) / len(pages_to_process))
            gc.collect()

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
            st.warning("No structured product specifications were found in the selected range.")

        doc.close()
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        gc.collect()
