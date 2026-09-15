import io
import os
import pandas as pd
import pdfplumber
import streamlit as st

# -----------------------------------------------------------------------------
# 1. Page Configuration & Styling
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Finishes Schedule Extractor",
    page_icon="📋",
    layout="wide",
)

st.title("📋 Finishes Schedule Extractor")
st.markdown(
    "Upload an interior design or architectural catalog PDF to extract materials and finish schedules into a categorized multi-tab Excel file."
)

# -----------------------------------------------------------------------------
# 2. Sidebar Page Range Controls (Target Summary Pages)
# -----------------------------------------------------------------------------
st.sidebar.header("📄 Page Range Settings")
st.sidebar.info(
    "Catalog summaries are usually located in the last few pages. Restricting the page range prevents processing overhead and keeps the LLM focused."
)

parse_mode = st.sidebar.radio(
    "Pages to Extract",
    ["Last N Pages (Summary Only)", "Custom Page Range", "All Pages"],
    index=0,
)

# -----------------------------------------------------------------------------
# 3. PDF Parsing Engine (Preserves Full Context)
# -----------------------------------------------------------------------------
def extract_finishes_from_pdf(pdf_file, start_page: int, end_page: int):
    """Parses selected PDF pages and preserves complete multi-word cell text without truncating context."""
    all_rows = []

    with pdfplumber.open(pdf_file) as pdf:
        total_pdf_pages = len(pdf.pages)
        
        # Clamp page bounds
        s_page = max(0, start_page - 1)
        e_page = min(total_pdf_pages, end_page)
        target_pages = pdf.pages[s_page:e_page]
        num_target_pages = len(target_pages)

        if num_target_pages == 0:
            return pd.DataFrame()

        progress_bar = st.progress(0)

        for idx, page in enumerate(target_pages):
            # Extract tables using text-alignment tolerances to keep full context together
            tables = page.extract_tables(
                table_settings={
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "snap_tolerance": 3,
                    "join_tolerance": 3,
                }
            )

            if not tables:
                # Fallback: Extract full text lines cleanly
                text = page.extract_text(layout=False)
                if text:
                    lines = text.split("\n")
                    for line in lines:
                        # Split by tab or double space to retain full phrases intact
                        parts = [p.strip() for p in line.split("  ") if p.strip()]
                        if len(parts) >= 2:
                            all_rows.append(parts)
            else:
                for table in tables:
                    if not table or len(table) < 1:
                        continue

                    for row in table:
                        # Clean whitespace while KEEPING complete multi-word strings intact
                        cleaned_row = [
                            " ".join(cell.split()) if cell else ""
                            for cell in row
                        ]

                        # Keep row if at least one cell has meaningful text
                        if any(cleaned_row):
                            all_rows.append(cleaned_row)

            progress_bar.progress((idx + 1) / num_target_pages)

    if not all_rows:
        return pd.DataFrame()

    # Equalize row lengths across all data rows
    max_cols = max(len(r) for r in all_rows)
    normalized_rows = [r + [""] * (max_cols - len(r)) for r in all_rows]

    # Assign headers from first detected row
    headers = [f"Column_{i+1}" for i in range(max_cols)]
    df = pd.DataFrame(normalized_rows, columns=headers)

    # Clean out repeated header rows
    first_row = df.iloc[0].tolist()
    df = df[~(df == first_row).all(axis=1)].reset_index(drop=True)
    df.columns = [str(c).title() if str(c).strip() else f"Field_{i+1}" for i, c in enumerate(first_row)]

    return df

# -----------------------------------------------------------------------------
# 4. Context-Aware Categorization & Multi-Tab Export
# -----------------------------------------------------------------------------
def export_to_multi_tab_excel(df):
    """Categorizes finishes while keeping full text description context in each sheet."""
    output = io.BytesIO()

    # Keywords for category matching
    category_keywords = {
        "Stone & Marble": ["stone", "marble", "granite", "travertine"],
        "Wood & Timber": ["wood", "timber", "veneer", "oak", "walnut"],
        "Tiles & Ceramics": ["tile", "ceramic", "porcelain"],
        "Metals": ["metal", "brass", "bronze", "steel", "aluminum"],
        "Fabrics & Leather": ["fabric", "leather", "upholstery", "carpet"],
        "Glass & Glazing": ["glass", "glazing", "mirror"],
    }

    # Find candidate column for categorization
    cat_col = None
    for col in df.columns:
        col_lower = str(col).lower()
        if any(k in col_lower for k in ["material", "category", "finish", "description", "item", "type"]):
            cat_col = col
            break

    if not cat_col:
        cat_col = df.columns[0]

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        assigned_rows = {cat: [] for cat in category_keywords.keys()}
        unassigned_rows = []

        for _, row in df.iterrows():
            row_text_full = " ".join([str(val).lower() for val in row.values])
            matched = False

            for cat_name, keywords in category_keywords.items():
                if any(kw in row_text_full for kw in keywords):
                    assigned_rows[cat_name].append(row)
                    matched = True
                    break

            if not matched:
                unassigned_rows.append(row)

        # Write categorized tabs with complete row context
        wrote_any_sheet = False
        for cat_name, rows_list in assigned_rows.items():
            if rows_list:
                cat_df = pd.DataFrame(rows_list)
                cat_df.to_excel(writer, sheet_name=cat_name[:31], index=False)
                wrote_any_sheet = True

        if unassigned_rows:
            gen_df = pd.DataFrame(unassigned_rows)
            gen_df.to_excel(writer, sheet_name="General Finishes", index=False)
            wrote_any_sheet = True

        if not wrote_any_sheet:
            df.to_excel(writer, sheet_name="Master Schedule", index=False)

    return output.getvalue()

# -----------------------------------------------------------------------------
# 5. App Execution
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload Catalog PDF", type=["pdf"], help="Select a Finishes Schedule PDF catalog"
)

if uploaded_file is not None:
    # Determine page bounds
    with pdfplumber.open(uploaded_file) as temp_pdf:
        total_pages = len(temp_pdf.pages)

    st.sidebar.write(f"Total PDF Pages: **{total_pages}**")

    if parse_mode == "Last N Pages (Summary Only)":
        num_summary = st.sidebar.number_input(
            "Number of summary pages from the end", min_value=1, max_value=total_pages, value=min(5, total_pages)
        )
        start_p = total_pages - num_summary + 1
        end_p = total_pages
    elif parse_mode == "Custom Page Range":
        start_p = st.sidebar.number_input("Start Page", min_value=1, max_value=total_pages, value=max(1, total_pages - 10))
        end_p = st.sidebar.number_input("End Page", min_value=start_p, max_value=total_pages, value=total_pages)
    else:
        start_p = 1
        end_p = total_pages

    st.info(f"Processing `{uploaded_file.name}` (Pages {start_p} to {end_p})...")

    with st.spinner("Extracting summary table schedule..."):
        extracted_df = extract_finishes_from_pdf(uploaded_file, start_p, end_p)

    if not extracted_df.empty:
        st.success(
            f"Extracted **{len(extracted_df)} schedule rows** across pages {start_p}–{end_p} with full text context preserved!"
        )

        st.subheader("Schedule Data Preview")
        st.dataframe(extracted_df, use_container_width=True)

        excel_data = export_to_multi_tab_excel(extracted_df)

        st.download_button(
            label="📥 Download Categorized Excel Schedule (.xlsx)",
            data=excel_data,
            file_name=f"{os.path.splitext(uploaded_file.name)[0]}_Summary_Schedule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error(
            "No table data found in the selected page range. Try adjusting the page range in the sidebar."
        )
