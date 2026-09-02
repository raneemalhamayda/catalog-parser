import streamlit as st
from google import genai
from PIL import Image
import json
import pandas as pd

# ==========================================
# 1. API Configuration
# ==========================================
# Paste your free Google AI Studio API key here
GEMINI_API_KEY = "AQ.Ab8RN6LrP7qnKYjGOBafR0u3lLfLZRpRayrqwHxqWQqN-ITtIA"

client = genai.Client(api_key=GEMINI_API_KEY)

# Streamlit Page Config
st.set_page_config(
    page_title="AI FF&E Catalog Parser",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Visual FF&E Catalog Parser (Free Tier)")
st.write("Upload a catalog page image to extract product specs, dimensions, and materials for free.")

uploaded_file = st.file_uploader(
    "Drag and drop a catalog PDF page or image...", 
    type=["jpg", "jpeg", "png"]
)

if uploaded_file is not None:
    col1, col2 = st.columns([1, 1])
    
    # Open image using Pillow
    image = Image.open(uploaded_file)
    
    with col1:
        st.subheader("Source Image")
        st.image(image, use_container_width=True)
    
    with col2:
        st.subheader("Extracted Specification")
        
        if st.button("Extract Specs with AI", type="primary", use_container_width=True):
            with st.spinner("Analyzing catalog page with Gemini..."):
                try:
                    prompt = """
                    Analyze this interior design catalog image and extract the product specifications into structured JSON format.
                    
                    Extract the following keys strictly:
                    - category: (e.g., Conference Table, Task Chair, Credenza)
                    - model_number: (e.g., HFIS-MDH23)
                    - length_mm: (numeric value in mm if available)
                    - width_mm: (numeric value in mm if available)
                    - height_mm: (numeric value in mm if available)
                    - primary_materials: (list of materials seen or mentioned)
                    - color_finish: (color of frame, top, or fabric)
                    - key_features: (list of notable features like cable management, ergonomics)
                    
                    Return ONLY raw, valid JSON. Do not include markdown triple backticks.
                    """
                    
                    # Call Gemini 2.5 Flash Vision Model (Free Tier)
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=[image, prompt]
                    )
                    
                    # Clean and parse response text
                    raw_text = response.text.strip()
                    if raw_text.startswith("```"):
                        raw_text = raw_text.split("```")[1]
                        if raw_text.startswith("json"):
                            raw_text = raw_text[4:]
                    raw_text = raw_text.strip()
                    
                    parsed_data = json.loads(raw_text)
                    
                    st.success("Extraction Complete!")
                    st.json(parsed_data)
                    
                    # Prepare Excel/CSV Download
                    df = pd.DataFrame([parsed_data])
                    csv_data = df.to_csv(index=False).encode('utf-8')
                    
                    st.download_button(
                        label="📥 Download as CSV (Excel Format)",
                        data=csv_data,
                        file_name=f"spec_{parsed_data.get('model_number', 'item')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                    
                except Exception as e:
                    st.error(f"Error processing image: {e}")