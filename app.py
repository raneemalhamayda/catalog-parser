import gc
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
from openpyxl.styles import Alignment, Font, PatternFill
import pandas as pd
from PIL import Image as PILImage
from pydantic import BaseModel, Field
import streamlit as st

# -----------------------------------------------------------------------------
# Streamlit Page Configuration
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Finishes Schedule Generator",
    page_icon="📋",
    layout="wide",
)

st.title("📋 Finishes Schedule Extractor & Generator")
st.write(
    "Upload vendor catalogs or specifications to automatically extract items and generate a pre-formatted multi-tab Finishes Schedule in Excel."
)

# -----------------------------------------------------------------------------
# API Key Resolution & Client Initialization
# -----------------------------------------------------------------------------
api_key = None

if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
    api_key = st.secrets["GEMINI_API_KEY"].strip()
elif os.getenv("GEMINI_API_KEY"):
    api_key = os.getenv("GEMINI_API_KEY").strip()

if not api_key:
    api_key = st.sidebar.text_input("Enter Google Gemini API Key:", type="password")

if not api_key:
    st.error(
        "🔑 Google Gemini API Key not found. Please set `GEMINI_API_KEY` in Streamlit Secrets."
    )
    st.stop()

os.environ["GEMINI_API_KEY"] = api_key
client = genai.Client(api_key=api_key)


# -----------------------------------------------------------------------------
# Pydantic Schemas for Finishes Schedule
# -----------------------------------------------------------------------------
class FinishItemSpec(BaseModel):
    page_number: int = Field(
        description="The 1-based catalog page number where this finish appears"
    )
    finish_type: str = Field(
        description="Must be classified strictly into one of: Wood, Stone, Carpet, Tiling, Paint, Wallcovering, Metal, Glass & Mirror, Miscellaneous"
    )
    code: str = Field(
        description="Finish code or designation, e.g., SK.P.01, P.01, WD.02, or 'N/A'"
    )
    name_description: str = Field(
        description="Full description including product name, code, material details, application instructions, and color name/RAL."
    )
    manufacturer: str = Field(
        description="Manufacturer name and full contact/address details if shown, otherwise 'N/A'"
    )
    supplier: str = Field(
        description="Local distributor or supplier details if listed, otherwise 'N/A'"
    )
    location: str = Field(
        description="Target room or application area if mentioned, otherwise 'N/A'"
    )
    remarks: str = Field(
        description="Any extra notes or technical compliance details"
    )


class FinishesScheduleExtraction(BaseModel):
    items: list[FinishItemSpec] = Field(
        description="List of all extracted finish specification items"
    )


# -----------------------------------------------------------------------------
# Helper Functions
# -----------------------------------------------------------------------------
def save_uploaded_file_to_disk(uploaded_file):
    temp_dir = tempfile.gettempdir()
    file_path = os.path.join(temp_dir, uploaded_file.name)

    with open(file_path, "wb") as f:
        while chunk := uploaded_file.read(4 * 1024 * 1024):
            f.write(chunk)

    return file_path


def extract_single_page_pdf(doc, page_num_1_based):
    new_doc = fitz.open()
    new_doc.insert_pdf(
        doc, from_page=page_num_1_based - 1, to_page=page_num_1_based - 1
    )
    output_stream = io.BytesIO()
    new_doc.save(output_stream)
    single_bytes = output_stream.getvalue()
    new_doc.close()
    return single_bytes


def render_page_thumbnail(doc, page_num_1_based, max_size=(110, 110)):
    page_idx = page_num_1_based - 1
    if page_idx < 0 or page_idx >= len(doc):
        return None

    page = doc[page_idx]
    pix = page.get_pixmap(dpi=96)
    img = PILImage.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    img.thumbnail(max_size)
    return img


def create_formatted_schedule_excel(df, doc):
    """Generates a multi-tab Excel workbook formatted by material type."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    material_categories = [
        "Wood",
        "Stone",
        "Carpet",
        "Tiling",
        "Paint",
        "Wallcovering",
        "Metal",
        "Glass & Mirror",
        "Miscellaneous",
    ]

    header_fill = PatternFill(
        start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"
    )
    bold_font = Font(name="Calibri", size=10, bold=True)
    regular_font = Font(name="Calibri", size=9)
    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    align_left = Alignment(horizontal="left", vertical="top", wrap_text=True)

    headers = [
        "TYPE",
        "CODE",
        "PHOTO",
        "NAME / DESCRIPTION",
        "MANUFACTURER",
        "SUPPLIER",
        "LOCATION",
        "REMARKS",
    ]

    col_widths = {
        "A": 12,  # TYPE
        "B": 12,  # CODE
        "C": 18,  # PHOTO
        "D": 35,  # NAME / DESCRIPTION
        "E": 28,  # MANUFACTURER
        "F": 25,  # SUPPLIER
        "G": 28,  # LOCATION
        "H": 20,  # REMARKS
    }

    for cat in material_categories:
        ws = wb.create_sheet(title=cat)

        # Apply Headers
        ws.append(headers)
        for col_num in range(1, 9):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = bold_font
            cell.alignment = align_center

        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

        # Filter extracted data by material tab
        cat_items = pd.DataFrame()
        if not df.empty and "finish_type" in df.columns:
            cat_items = df[
                df["finish_type"].str.strip().str.lower() == cat.lower()
            ]

        # Populate rows
        for idx, (_, row) in enumerate(cat_items.iterrows()):
            excel_row = idx + 2
            ws.row_dimensions[excel_row].height = 100

            ws.cell(
                row=excel_row, column=1, value=str(row.get("finish_type", ""))
            ).alignment = align_center
            ws.cell(
                row=excel_row, column=2, value=str(row.get("code", ""))
            ).alignment = align_center

            # Photo (Thumbnail)
            actual_page_num = int(row.get("page_number", 1))
            pil_img = render_page_thumbnail(doc, actual_page_num)
            if pil_img:
                img_io = io.BytesIO()
                pil_img.save(img_io, format="JPEG", quality=80)
                img_io.seek(0)
                img_obj = OpenPyXlImage(img_io)
                ws.add_image(img_obj, f"C{excel_row}")

            ws.cell(
                row=excel_row,
                column=4,
                value=str(row.get("name_description", "")),
            ).alignment = align_left
            ws.cell(
                row=excel_row, column=5, value=str(row.get("manufacturer", ""))
            ).alignment = align_left
            ws.cell(
                row=excel_row, column=6, value=str(row.get("supplier", ""))
            ).alignment = align_left
            ws.cell(
                row=excel_row, column=7, value=str(row.get("location", ""))
            ).alignment = align_left
            ws.cell(
                row=excel_row, column=8, value=str(row.get("remarks", ""))
            ).alignment = align_left

            for col_idx in range(1, 9):
                ws.cell(row=excel_row, column=col_idx).font = regular_font

    output_stream = io.BytesIO()
    wb.save(output_stream)
    return output_stream.getvalue()


def process_single_page_with_retry(single_pdf_bytes, page_num, max_retries=5):
    models_to_try = ["gemini-3.5-flash-lite", "gemini-2.5-flash"]

    prompt = f"""
    You are analyzing Page {page_num} of a commercial interior finishes/materials catalog or specification sheet.
    
    Task:
    1. Extract every finish item listed on this page.
    2. Classify `finish_type` strictly into one of: Wood, Stone, Carpet, Tiling, Paint, Wallcovering, Metal, Glass & Mirror, Miscellaneous.
    3. Extract or assign the finish `code` if available (e.g., SK.P.01).
    4. DIMENSION FOCUS: Actively scan the page for any explicit OR implied dimensions (e.g., sheet size, tile format, roll width, thickness/depth, length, height, yield per m²). 
       - Look for numbers followed by mm, cm, m, inches, or format notations like "60x60", "1200x600x20".
       - If dimensions are embedded inside product text, separate them clearly and include them at the start of `name_description`.
       - If dimensions are completely absent from the page, explicitly write "[DIMENSIONS MISSING FROM CATALOG]" inside `remarks`.
    5. Compile full technical text into `name_description` (translate any foreign text into English).
    6. Extract `manufacturer` name, contact, and address details.
    7. Extract `supplier`, `location`, and `remarks` if mentioned, otherwise write "N/A".
    8. Set `page_number` to {page_num}.
    """

    last_error = None

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
                        response_schema=FinishesScheduleExtraction,
                    ),
                )
                parsed = json.loads(response.text)
                return parsed.get("items", [])
            except Exception as e:
                last_error = str(e)
                if (
                    "429" in last_error
                    or "RESOURCE_EXHAUSTED" in last_error
                    or "503" in last_error
                ):
                    wait_time = (attempt + 1) * 10
                    st.sidebar.info(
                        f"⏳ Quota pause on Page {page_num}. Waiting {wait_time}s before retrying..."
                    )
                    time.sleep(wait_time)
                elif "404" in last_error or "NOT_FOUND" in last_error:
                    break
                else:
                    break

    st.sidebar.warning(f"⚠️ Page {page_num}: {last_error}")
    return []


# -----------------------------------------------------------------------------
# Main Application UI
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Choose a catalog or specification PDF file", type=["pdf"]
)

if uploaded_file:
    with st.spinner("Loading PDF file..."):
        temp_pdf_path = save_uploaded_file_to_disk(uploaded_file)
        doc = fitz.open(temp_pdf_path)
        total_pages = len(doc)

    st.sidebar.header("📄 Page Options")
    process_mode = st.sidebar.radio(
        "Select Range:", ["Process Sample Range", "Process Full Document"]
    )

    if process_mode == "Process Sample Range":
        start_page = st.sidebar.number_input(
            "Start Page", min_value=1, max_value=total_pages, value=1
        )
        default_end = min(start_page + 4, total_pages)
        end_page = st.sidebar.number_input(
            "End Page",
            min_value=start_page,
            max_value=total_pages,
            value=default_end,
        )
    else:
        start_page = 1
        end_page = total_pages

    st.info(
        f"Ready to process pages **{start_page} to {end_page}** (Total: {total_pages} pages in document)."
    )

    if st.button("Generate Finishes Schedule Excel", type="primary"):
        all_extracted_items = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        pages_to_process = list(range(start_page, end_page + 1))

        for i, current_page in enumerate(pages_to_process):
            status_text.text(f"Processing page {current_page} of {end_page}...")
            single_bytes = extract_single_page_pdf(doc, current_page)
            page_items = process_single_page_with_retry(
                single_bytes, current_page
            )

            if page_items:
                st.sidebar.write(
                    f"✅ Page {current_page}: Found {len(page_items)} item(s)"
                )

            all_extracted_items.extend(page_items)
            progress_bar.progress((i + 1) / len(pages_to_process))

            del single_bytes
            gc.collect()
            time.sleep(2.5)

        status_text.empty()
        progress_bar.empty()

        if all_extracted_items:
            df = pd.DataFrame(all_extracted_items)
            st.success(
                f"Extraction complete! Extracted {len(df)} item(s) across pages {start_page} to {end_page}."
            )

            st.subheader("Schedule Preview")
            st.dataframe(df, use_container_width=True)

            excel_bytes = create_formatted_schedule_excel(df, doc)

            st.download_button(
                label="📥 Download Multi-Tab Finishes Schedule (.xlsx)",
                data=excel_bytes,
                file_name=f"{os.path.splitext(uploaded_file.name)[0]}_Finishes_Schedule.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("No specification items were found in the selected range.")

        doc.close()
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        gc.collect()
