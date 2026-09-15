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
