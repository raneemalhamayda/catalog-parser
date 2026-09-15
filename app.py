import io
import os
import tempfile
import fitz  # PyMuPDF
import openpyxl
from openpyxl.drawing.image import Image as OpenPyxlImage
from google import genai
from google.genai import types
import pandas as pd
from pydantic import BaseModel, Field
import streamlit as st

st.set_page_config(page_title="Finishes Schedule Extractor", layout="wide")

st.title("📋 Finishes Schedule Extractor")
st.markdown("Extract material finish schedules and visual pattern swatches into a categorized Excel sheet.")

# -----------------------------------------------------------------------------
# 1. API Key Setup
# -----------------------------------------------------------------------------
api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    api_key = st.secrets.get("GEMINI_API_KEY", "")

client = genai.Client(api_key=api_key) if api_key else None

# -----------------------------------------------------------------------------
# 2. Pydantic Schemas
# -----------------------------------------------------------------------------
class FinishItem(BaseModel):
    item_code: str = Field(description="Pattern ID or material code (e.g. A1 2176, C8 9483)", default="N/A")
    category: str = Field(description="Category (e.g., Sculpture Pattern, Fabric, Wood)", default="General")
    material_name: str = Field(description="Material description or pattern group name")
    specification: str = Field(description="Technical details, composition, or dimensions", default="")
    page_number: int = Field(description="1-based page number", default=1)

class ScheduleSchema(BaseModel):
    finishes: list[FinishItem] = Field(description="List of extracted finish items")

# -----------------------------------------------------------------------------
# 3. Native Visual PDF Extraction
# -----------------------------------------------------------------------------
def process_full_catalog(tmp_pdf_path):
    if not client:
        st.error("GEMINI_API_KEY is missing from Streamlit Secrets.")
        return pd.DataFrame()

    file_ref = client.files.upload(file=tmp_pdf_path)

    prompt = """
    Analyze every page of the PDF catalog. Extract all finish materials and visual design patterns.
    - Set 'item_code' to the exact pattern code (e.g., 'A1 2176', 'C8 9483').
    - Set 'category' to 'Sculpture Pattern', 'Fabric', 'Wood', 'Metal', etc.
    - Set 'material_name' to the pattern group title (e.g., 'Sculpture Design Pattern').
    - Set 'page_number' to the 1-based page number where the item appears.
    """

    response = client.models.generate_content(
        model="gemini-3.5-flash",
        contents=[file_ref, prompt],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=ScheduleSchema,
        ),
    )

    data = ScheduleSchema.model_validate_json(response.text)
    items = [item.model_dump() for item in data.finishes]
    return pd.DataFrame(items)

# -----------------------------------------------------------------------------
# 4. Memory-Safe Excel Generator with Embedded Swatches
# -----------------------------------------------------------------------------
def create_excel_with_images(df, pdf_path):
    output = io.BytesIO()
    doc = fitz.open(pdf_path)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    categories = df["category"].unique() if "category" in df.columns else ["Master Schedule"]

    for cat in categories:
        sheet_title = str(cat).strip()[:31]
        ws = wb.create_sheet(title=sheet_title)

        ws.append(["Item Code", "Category", "Material Name", "Technical Specification", "Visual Swatch"])
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 20
        ws.column_dimensions["C"].width = 30
        ws.column_dimensions["D"].width = 45
        ws.column_dimensions["E"].width = 22

        cat_df = df[df["category"] == cat] if "category" in df.columns else df

        for row_idx, (_, row) in enumerate(cat_df.iterrows(), start=2):
            ws.cell(row=row_idx, column=1, value=str(row["item_code"]))
            ws.cell(row=row_idx, column=2, value=str(row["category"]))
            ws.cell(row=row_idx, column=3, value=str(row["material_name"]))
            ws.cell(row=row_idx, column=4, value=str(row["specification"]))

            # Safe image rendering wrapper
            try:
                page_num = max(1, int(row.get("page_number", 1))) - 1
                if page_num < len(doc):
                    page = doc[page_num]
                    
                    # Low-DPI rasterization (90 DPI prevents Streamlit Cloud OOM crashes)
                    pix = page.get_pixmap(dpi=90)
                    img_bytes = pix.tobytes("png")

                    if img_bytes:
                        img_file = io.BytesIO(img_bytes)
                        img = OpenPyxlImage(img_file)
                        
                        # Resize thumbnail for cell
                        img.width = 110
                        img.height = 90

                        ws.row_dimensions[row_idx].height = 75
                        ws.add_image(img, f"E{row_idx}")
            except Exception:
                # If image rendering fails for a row, skip image gracefully without crashing the app
                pass

    doc.close()
    wb.save(output)
    return output.getvalue()

# -----------------------------------------------------------------------------
# 5. Streamlit UI
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader("Upload PDF Catalog", type=["pdf"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_pdf_path = tmp_file.name

    try:
        with st.spinner("Processing document & preparing Excel download..."):
            df = process_full_catalog(tmp_pdf_path)

        if not df.empty:
            st.success(f"Extracted {len(df)} items successfully!")
            st.dataframe(df[["item_code", "category", "material_name", "specification", "page_number"]], use_container_width=True)

            excel_bytes = create_excel_with_images(df, tmp_pdf_path)
            st.download_button(
                "📥 Download Excel Schedule (.xlsx)",
                data=excel_bytes,
                file_name=f"{os.path.splitext(uploaded_file.name)[0]}_Schedule.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error("No schedule items found.")
    finally:
        if os.path.exists(tmp_pdf_path):
            os.remove(tmp_pdf_path)
