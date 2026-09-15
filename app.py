import io
import os
from google import genai
from google.genai import types
import pandas as pd
import pdfplumber
from pydantic import BaseModel, Field
import streamlit as st

st.set_page_config(page_title="Finishes Schedule Extractor", layout="wide")

st.title("📋 Finishes Schedule Extractor & Translator")
st.markdown(
    "Parses bilingual catalog summary pages, translates Chinese specifications to English, and outputs a clean Excel schedule."
)

# -----------------------------------------------------------------------------
# 1. API Key Setup
# -----------------------------------------------------------------------------
api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    api_key = st.secrets.get("GEMINI_API_KEY", "")

client = genai.Client(api_key=api_key) if api_key else None


# -----------------------------------------------------------------------------
# 2. Pydantic Schema for Clean Schedule Output
# -----------------------------------------------------------------------------
class FinishItem(BaseModel):
    item_code: str = Field(
        description="Material code or tag (e.g., ST-01, WD-02, or N/A)",
        default="N/A",
    )
    category: str = Field(
        description="Category in English (e.g. Stone, Metal, Wood, Fabric, Glass, Paint)",
        default="General",
    )
    material_name: str = Field(
        description="Material/Item description translated fully into English"
    )
    specification: str = Field(
        description="Technical specifications, dimensions, or finishes translated into English",
        default="",
    )


class ScheduleSchema(BaseModel):
    finishes: list[FinishItem] = Field(
        description="Clean list of extracted and translated material finish items"
    )


# -----------------------------------------------------------------------------
# 3. Extraction & AI Translation Pipeline
# -----------------------------------------------------------------------------
def process_catalog_pdf(pdf_file, last_n_pages=5):
    raw_text = ""
    with pdfplumber.open(pdf_file) as pdf:
        total = len(pdf.pages)
        target_indices = list(range(max(0, total - last_n_pages), total))

        for idx in target_indices:
            page_text = pdf.pages[idx].extract_text()
            if page_text:
                raw_text += f"\n--- PAGE {idx + 1} ---\n" + page_text

    if not raw_text.strip():
        return pd.DataFrame()

    if not client:
        st.error(
            "GEMINI_API_KEY is not configured in Streamlit Secrets. Please set it to enable auto-translation."
        )
        return pd.DataFrame()

    # Pass raw text to Gemini for cleaning & translation
    prompt = """
    You are an interior design technical document specialist.
    Analyze the raw catalog text provided below:
    1. Filter out general marketing sentences, promotional slogans, and narrative fluff.
    2. Extract only real material finish items, specifications, and hardware components.
    3. Translate all Chinese text completely into professional English.
    4. Categorize materials strictly (e.g., Metal, Wood, Stone, Fabric, Glass, Paint).
    """

    response = client.models.generate_content(
        model="gemini-3.5-flash",
        contents=[prompt, raw_text],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=ScheduleSchema,
        ),
    )

    data = ScheduleSchema.model_validate_json(response.text)
    items = [item.model_dump() for item in data.finishes]
    df = pd.DataFrame(items)

    # Rename columns cleanly for Excel
    df.columns = [
        "Item Code",
        "Category",
        "Material Name (English)",
        "Technical Specification",
    ]
    return df


def create_categorized_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Group into tabs by Category
        if "Category" in df.columns:
            for cat, group in df.groupby("Category"):
                sheet_title = str(cat).strip()[:31]
                group.to_excel(writer, sheet_name=sheet_title, index=False)
        else:
            df.to_excel(writer, sheet_name="Master Schedule", index=False)
    return output.getvalue()


# -----------------------------------------------------------------------------
# 4. Streamlit UI
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload PDF Catalog", type=["pdf"], help="Supports Chinese & English catalogs"
)

if uploaded_file:
    num_pages = st.sidebar.number_input(
        "Summary pages to read from end", min_value=1, value=5
    )

    with st.spinner(
        "Extracting, translating Chinese text, and structuring table..."
    ):
        df = process_catalog_pdf(uploaded_file, last_n_pages=num_pages)

    if not df.empty:
        st.success(f"Successfully extracted & translated {len(df)} finish items!")
        st.dataframe(df, use_container_width=True)

        excel_bytes = create_categorized_excel(df)
        st.download_button(
            "📥 Download Translated Excel Schedule (.xlsx)",
            data=excel_bytes,
            file_name=f"{os.path.splitext(uploaded_file.name)[0]}_English_Schedule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error(
            "Could not extract material schedules from the selected pages."
        )
