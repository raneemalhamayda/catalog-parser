import io
import os
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title="Catalog Finishes Parser", layout="wide")

st.title("📋 Finishes Schedule Extractor")
st.markdown("Upload a PDF catalog to extract material schedules into a categorized Excel file.")

# Sidebar Controls
st.sidebar.header("📄 Page Range Settings")
parse_mode = st.sidebar.radio(
    "Pages to Extract",
    ["Auto-Detect Summary Pages", "All Pages", "Last N Pages"],
    index=0
)

def extract_finishes(pdf_file, parse_mode, last_n=5):
    all_rows = []
    with pdfplumber.open(pdf_file) as pdf:
        total = len(pdf.pages)
        
        # Determine target pages
        if parse_mode == "Last N Pages":
            target_indices = list(range(max(0, total - last_n), total))
        elif parse_mode == "Auto-Detect Summary Pages":
            keywords = ["schedule", "finishes", "summary", "legend", "material"]
            target_indices = [
                i for i, page in enumerate(pdf.pages)
                if any(kw in (page.extract_text() or "").lower() for kw in keywords)
            ]
            if not target_indices:
                target_indices = list(range(max(0, total - 5), total))
        else:
            target_indices = list(range(total))

        st.info(f"Extracting from {len(target_indices)} pages...")
        prog = st.progress(0)

        for step, page_idx in enumerate(target_indices):
            page = pdf.pages[page_idx]
            
            # Try table extraction
            tables = page.extract_tables({"vertical_strategy": "lines", "horizontal_strategy": "lines", "snap_tolerance": 5})
            if not tables or not any(tables):
                tables = page.extract_tables({"vertical_strategy": "text", "horizontal_strategy": "text", "snap_tolerance": 4})

            if tables and any(tables):
                for table in tables:
                    if table:
                        for row in table:
                            cleaned = [" ".join(cell.split()) if cell else "" for cell in row]
                            if any(cleaned):
                                all_rows.append(cleaned)
            else:
                # Text fallback
                text = page.extract_text()
                if text:
                    for line in text.split("\n"):
                        parts = re.split(r"\s{2,}|\t", line.strip())
                        if len(parts) >= 2:
                            all_rows.append(parts)

            prog.progress((step + 1) / len(target_indices))

    if not all_rows:
        return pd.DataFrame()

    max_cols = max(len(r) for r in all_rows)
    normalized = [r + [""] * (max_cols - len(r)) for r in all_rows]
    df = pd.DataFrame(normalized)
    
    # Set headers
    first_row = df.iloc[0].tolist()
    if any(first_row):
        df = df.iloc[1:].reset_index(drop=True)
        df.columns = [str(c).title() if str(c).strip() else f"Col_{i+1}" for i, c in enumerate(first_row)]

    return df

def create_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Finishes Schedule", index=False)
    return output.getvalue()

uploaded_file = st.file_uploader("Upload PDF Catalog", type=["pdf"])

if uploaded_file:
    last_n_val = 5
    if parse_mode == "Last N Pages":
        last_n_val = st.sidebar.number_input("N pages", min_value=1, value=5)

    with st.spinner("Processing PDF..."):
        df = extract_finishes(uploaded_file, parse_mode, last_n_val)

    if not df.empty:
        st.success(f"Extracted {len(df)} rows!")
        st.dataframe(df, use_container_width=True)
        
        excel_bytes = create_excel(df)
        st.download_button(
            "📥 Download Excel (.xlsx)",
            data=excel_bytes,
            file_name="Finishes_Schedule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("No tables found in selected page range.")
