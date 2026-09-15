import io
import os
import tempfile
import pandas as pd
import streamlit as st
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

st.set_page_config(page_title="Finishes Schedule Extractor", layout="wide")

st.title("📋 Finishes Schedule Extractor")
st.markdown("Extract material finish schedules and design pattern codes from PDF catalogs into a categorized Excel file.")

# ----------------------------------------------------------import io
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

st.set_page_config(
    page_title="Finishes Schedule Extractor with Visual Swatches", layout="wide"
)

st.title("📋 Finishes Schedule Extractor")
st.markdown(
    "Extract material finish schedules and visual pattern swatches directly into a categorized Excel sheet with embedded image previews."
)

# -----------------------------------------------------------------------------
# 1. API Key Setup
# -----------------------------------------------------------------------------
api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    api_key = st.secrets.get("GEMINI_API_KEY", "")

client = genai.Client(api_key=api_key) if api_key else None


# -----------------------------------------------------------------------------
# 2. Pydantic Extraction Schemas
# -----------------------------------------------------------------------------
class FinishItem(BaseModel):
    item_code: str = Field(
        description="Material code, pattern ID, product tag, or swatch number (e.g., ST-01, A1 2176, C8 9483)",
        default="N/A",
    )
    category: str = Field(
        description="Material category (e.g., Sculpture Pattern, Design Pattern, Fabric, Wood, Metal, Stone)",
        default="General",
    )
    material_name: str = Field(
        description="Full material description or pattern group title (e.g., Sculpture Design Pattern)"
    )
    specification: str = Field(
        description="Complete technical specifications, dimensions, composition, or application details",
        default="",
    )
    page_number: int = Field(
        description="1-based page number where this item or visual pattern swatch was found",
        default=1,
    )


class ScheduleSchema(BaseModel):
    finishes: list[FinishItem] = Field(
        description="Exhaustive list of all extracted material finish items and pattern codes across all pages"
    )


# -----------------------------------------------------------------------------
# 3. Native Visual PDF Extraction Pipeline
# -----------------------------------------------------------------------------
def process_full_catalog(uploaded_file, tmp_pdf_path):
    if not client:
        st.error("GEMINI_API_KEY is missing from Streamlit Secrets.")
        return pd.DataFrame()

    file_ref = client.files.upload(file=tmp_pdf_path)

    prompt = """
    Analyze every single page of the attached PDF catalog, including visual matrix pages, pattern grids, and spec tables.
    
    Extraction Rules:
    1. Extract standard finish materials (Fabrics, Acoustic Panels, Wood, Metal, Glass, Stone, Paint).
    2. ALSO extract visual design patterns, swatch codes, pattern IDs, and design keys (e.g., 'Sculpture design pattern', pattern codes like 'A1 2176', 'A2 3692', 'C8 9483', 'D23 14291', etc.).
    3. For visual pattern grids:
       - Set 'item_code' to the exact pattern code/ID (e.g., 'A1 2176' or 'C8 9483').
       - Set 'category' to 'Sculpture Pattern' or 'Design Pattern'.
       - Set 'material_name' to the pattern title (e.g., 'Sculpture Design Pattern').
       - Set 'specification' to any associated dimensions or notes.
       - Set 'page_number' to the exact 1-based page number where the swatch image appears.
    4. Capture EVERY item across ALL pages!
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
# 4. Crop Visual Swatch Images and Build Excel with Image Columns
# -----------------------------------------------------------------------------
def create_excel_with_images(df, pdf_path):
    output = io.BytesIO()
    doc = fitz.open(pdf_path)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    categories = df["category"].unique() if "category" in df.columns else ["Master Schedule"]

    for cat in categories:
        sheet_title = str(cat).strip()[:31]
        ws = wb.create_sheet(title=sheet_title)

        # Write Headers
        headers = [
            "Item Code",
            "Category",
            "Material Name",
            "Technical Specification",
            "Visual Swatch Image",
        ]
        ws.append(headers)

        # Adjust header formatting & column widths
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 20
        ws.column_dimensions["C"].width = 30
        ws.column_dimensions["D"].width = 45
        ws.column_dimensions["E"].width = 25  # Image Column Width

        cat_df = df[df["category"] == cat] if "category" in df.columns else df

        for row_idx, (_, row) in enumerate(cat_df.iterrows(), start=2):
            ws.cell(row=row_idx, column=1, value=str(row["item_code"]))
            ws.cell(row=row_idx, column=2, value=str(row["category"]))
            ws.cell(row=row_idx, column=3, value=str(row["material_name"]))
            ws.cell(row=row_idx, column=4, value=str(row["specification"]))

            # Extract & crop swatch thumbnail from PDF page
            page_num = max(1, int(row.get("page_number", 1))) - 1
            if page_num < len(doc):
                page = doc[page_num]
                image_list = page.get_images(full=True)

                if image_list:
                    # Render page to raster pixmap to capture pattern swatch crops
                    pix = page.get_pixmap(dpi=150)
                    img_bytes = pix.tobytes("png")

                    img_file = io.BytesIO(img_bytes)
                    img = OpenPyxlImage(img_file)

                    # Scale image down to fit nicely in Excel cell
                    img.width = 120
                    img.height = 100

                    # Row height expansion to accommodate thumbnail
                    ws.row_dimensions[row_idx].height = 80

                    # Anchor image inside 'Visual Swatch Image' column (Column E)
                    cell_ref = f"E{row_idx}"
                    ws.add_image(img, cell_ref)

    doc.close()
    wb.save(output)
    return output.getvalue()


# -----------------------------------------------------------------------------
# 5. Streamlit UI
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader("Upload PDF Catalog", type=["pdf"])

if uploaded_file:
    # Save temporary PDF for processing and image extraction
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_pdf_path = tmp_file.name

    try:
        with st.spinner("Extracting schedule data & cropping pattern swatch images..."):
            df = process_full_catalog(uploaded_file, tmp_pdf_path)

        if not df.empty:
            st.success(
                f"Successfully extracted {len(df)} schedule items with visual pattern swatches!"
            )
            st.dataframe(
                df[
                    [
                        "item_code",
                        "category",
                        "material_name",
                        "specification",
                        "page_number",
                    ]
                ],
                use_container_width=True,
            )

            excel_bytes = create_excel_with_images(df, tmp_pdf_path)
            st.download_button(
                "📥 Download Excel Schedule with Embedded Swatches (.xlsx)",
                data=excel_bytes,
                file_name=f"{os.path.splitext(uploaded_file.name)[0]}_Visual_Schedule.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error("No schedule items or pattern swatches found in this PDF.")
    finally:
        if os.path.exists(tmp_pdf_path):
            os.remove(tmp_pdf_path)------------------
# 1. API Key Setup
# -----------------------------------------------------------------------------
api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    api_key = st.secrets.get("GEMINI_API_KEY", "")

client = genai.Client(api_key=api_key) if api_key else None

# -----------------------------------------------------------------------------
# 2. Pydantic Extraction Schemas
# -----------------------------------------------------------------------------
class FinishItem(BaseModel):
    item_code: str = Field(
        description="Material code, pattern ID, product tag, or swatch number (e.g. ST-01, A1 2176, C8 9483, or N/A)",
        default="N/A"
    )
    category: str = Field(
        description="Material category (e.g., Fabric, Sculpture Pattern, Design Pattern, Acoustic Panels, Wood, Metal, Stone, Glass, Paint)",
        default="General"
    )
    material_name: str = Field(
        description="Full material description or pattern series title (e.g., Sculpture Design Pattern, Polyester Fiber Grooved Panel)"
    )
    specification: str = Field(
        description="Complete technical specifications, dimensions, composition, pattern numbers, or application details",
        default=""
    )

class ScheduleSchema(BaseModel):
    finishes: list[FinishItem] = Field(
        description="Exhaustive list of all extracted material finish items and pattern codes across the entire document"
    )

# -----------------------------------------------------------------------------
# 3. Native Visual PDF Extraction Pipeline
# -----------------------------------------------------------------------------
def process_full_catalog(uploaded_file):
    if not client:
        st.error("GEMINI_API_KEY is missing from Streamlit Secrets.")
        return pd.DataFrame()

    # Save temporary file for Gemini File API upload
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_path = tmp_file.name

    try:
        # Upload PDF directly to process visual pattern grids alongside spec tables
        file_ref = client.files.upload(file=tmp_path)

        prompt = """
        Analyze every single page of the attached PDF catalog, including visual matrix pages, pattern grids, and spec tables.
        
        Extraction Rules:
        1. Extract standard finish materials (Fabrics, Acoustic Panels, Wood, Metal, Glass, Stone, Paint).
        2. ALSO extract visual design patterns, swatch codes, pattern IDs, and design keys (e.g., 'Sculpture design pattern', pattern codes like 'A1 2176', 'A2 3692', 'C8 9483', 'D23 14291', etc.).
        3. For visual pattern grids:
           - Set 'item_code' to the exact pattern code/ID (e.g., 'A1 2176' or 'C8 9483').
           - Set 'category' to 'Design Pattern' or 'Sculpture Pattern'.
           - Set 'material_name' to the overall pattern group title (e.g., 'Sculpture Design Pattern').
           - Set 'specification' to any associated dimensions, numbers, or layout notes.
        4. Capture EVERY item across ALL pages—do not omit visual swatches or pattern grids!
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
        df = pd.DataFrame(items)

        if not df.empty:
            df.columns = ["Item Code", "Category", "Material Name", "Technical Specification"]

        return df

    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

def create_categorized_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if "Category" in df.columns:
            for cat, group in df.groupby("Category"):
                sheet_title = str(cat).strip()[:31]  # Excel 31-character limit
                group.to_excel(writer, sheet_name=sheet_title, index=False)
        else:
            df.to_excel(writer, sheet_name="Master Schedule", index=False)
    return output.getvalue()

# -----------------------------------------------------------------------------
# 4. Streamlit UI
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader("Upload PDF Catalog", type=["pdf"])

if uploaded_file:
    with st.spinner("Processing all document pages (spec tables & visual pattern grids)..."):
        df = process_full_catalog(uploaded_file)

    if not df.empty:
        st.success(f"Successfully extracted {len(df)} schedule items & design pattern codes!")
        st.dataframe(df, use_container_width=True)

        excel_bytes = create_categorized_excel(df)
        st.download_button(
            "📥 Download Excel Schedule (.xlsx)",
            data=excel_bytes,
            file_name=f"{os.path.splitext(uploaded_file.name)[0]}_Schedule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error("No material schedule items or pattern codes could be extracted from this PDF.")
