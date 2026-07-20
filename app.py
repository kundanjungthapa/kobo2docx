import os
import io
import re
import tempfile
import zipfile
import pandas as pd
import streamlit as st

from src.reader import KoboReader
from src.translator import detect_languages, get_choices_map
from src.docx_builder import create_document

# ---------------------------------------------------------
# PAGE CONFIGURATION
# ---------------------------------------------------------
st.set_page_config(
    page_title="Kobo2Docx | Survey Questionnaire Formatter",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Load External CSS Stylesheet
def load_css(file_path="assets/style.css"):
    if os.path.exists(file_path):
        with open(file_path, "r") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

load_css()

# ---------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/ms-word.png", width=64)
    st.title("Kobo2Docx Engine")
    st.caption("Version 1.2.0 • Production")
    
    st.markdown("🌐 **Project Landing Site**: [kundanjungthapa.github.io/kobo2docx](https://kundanjungthapa.github.io/kobo2docx/)")
    st.divider()
    
    st.markdown("### ⚙️ Formatting Rules")
    st.markdown("""
    - **Header Colors**: Deep Blue (`#2E5B82`)
    - **Group Banners**: Full-width Gray (`#F0F0F0`)
    - **Question Typography**: 11pt Cambria
    - **Skip-Logic Items**: Indented 0.25", 10pt Fainter Gray
    - **Text Fields**: Clean dotted lines (`........`)
    """)
    st.divider()
    
    st.markdown("### 🛠️ Help & Specs")
    with st.expander("Required Sheet Headers"):
        st.markdown("""
        **Survey Sheet:** `type`, `name`, `label::<lang>`, `hint::<lang>`, `relevant`  
        **Choices Sheet:** `list_name`, `name`, `label::<lang>`
        """)

# ---------------------------------------------------------
# HERO SECTION
# ---------------------------------------------------------
st.markdown("""
<div class="hero-box">
    <div class="hero-title">📋 Kobo2Docx Questionnaire Generator</div>
    <div class="hero-subtitle">Convert raw KoboToolbox / ODK Excel forms into cleanly styled, publication-ready Word documents in seconds.</div>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# WORKFLOW CARDS
# ---------------------------------------------------------
st.markdown("### 🚀 How It Works")
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown("""
    <div class="step-card">
        <div class="step-num">Step 1</div>
        <div class="step-title">Upload Form</div>
        <p style="font-size: 0.85rem; color: #64748b; margin-top: 0.4rem;">Select your .xlsx file containing survey, choices, and settings.</p>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="step-card">
        <div class="step-num">Step 2</div>
        <div class="step-title">Auto-Detect</div>
        <p style="font-size: 0.85rem; color: #64748b; margin-top: 0.4rem;">Identifies languages, choice maps, and group hierarchies.</p>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div class="step-card">
        <div class="step-num">Step 3</div>
        <div class="step-title">Apply Styles</div>
        <p style="font-size: 0.85rem; color: #64748b; margin-top: 0.4rem;">Formats skip logic, hints, fonts, and dotted response lines.</p>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown("""
    <div class="step-card">
        <div class="step-num">Step 4</div>
        <div class="step-title">Download</div>
        <p style="font-size: 0.85rem; color: #64748b; margin-top: 0.4rem;">Get single .docx files or a complete multi-language ZIP archive.</p>
    </div>
    """, unsafe_allow_html=True)

st.divider()

# ---------------------------------------------------------
# CONVERSION CONSOLE
# ---------------------------------------------------------
st.markdown("### 📂 Conversion Console")

uploaded_file = st.file_uploader(
    "Drag and drop your KoboToolbox file here (.xlsx)", 
    type=["xlsx"],
    help="Supports single or multi-language XLSForms"
)

if uploaded_file is not None:
    with st.status("Reading XLSForm structure...", expanded=True) as status:
        st.write("📂 Saving uploaded workbook to temporary workspace...")
        
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = os.path.join(temp_dir, uploaded_file.name)
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            try:
                st.write("🔍 Parsing survey, choices, and settings sheets...")
                reader = KoboReader(input_path)
                title = reader.get_title()
                
                st.write("🌐 Detecting multi-language columns...")
                detected_languages = detect_languages(reader.survey_df.columns)
                if not detected_languages:
                    detected_languages['Default'] = 'label'

                status.update(label="XLSForm Parsed Successfully!", state="complete", expanded=False)
                
                # Metrics Banner
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Form Title", str(title)[:20] + "..." if len(str(title)) > 20 else title)
                m2.metric("Total Rows", len(reader.survey_df))
                m3.metric("Choices Registered", len(reader.choices_df))
                m4.metric("Languages Found", len(detected_languages))

                # Actions & Data Preview
                tab1, tab2 = st.columns([2, 1])
                
                with tab1:
                    st.markdown("#### 📄 Detected Languages & Generation")
                    st.write(f"Found the following target languages: **{', '.join(detected_languages.keys())}**")
                    
                    if st.button("⚡ Generate Word Documents Now", type="primary", use_container_width=True):
                        with st.spinner("Rendering tables and building .docx files..."):
                            generated_files = []
                            base_name = os.path.splitext(uploaded_file.name)[0]

                            for lang_name, label_col in detected_languages.items():
                                clean_lang = re.sub(r'[^a-zA-Z0-9]', '_', lang_name).strip('_')
                                out_filename = f"{base_name}_{clean_lang}.docx"
                                out_path = os.path.join(temp_dir, out_filename)

                                choices_dict = get_choices_map(reader.choices_df, label_col)
                                create_document(lang_name, label_col, reader.survey_df, choices_dict, title, out_path)
                                generated_files.append((out_filename, out_path))

                            st.balloons()
                            st.success("All documents formatted successfully!")

                            st.markdown("#### 📥 Ready Downloads")
                            if len(generated_files) == 1:
                                filename, filepath = generated_files[0]
                                with open(filepath, "rb") as f:
                                    st.download_button(
                                        label=f"Download {filename}",
                                        data=f.read(),
                                        file_name=filename,
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        use_container_width=True
                                    )
                            else:
                                zip_buffer = io.BytesIO()
                                with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                                    for filename, filepath in generated_files:
                                        zip_file.write(filepath, arcname=filename)
                                zip_buffer.seek(0)

                                st.download_button(
                                    label="Download All Languages (.ZIP)",
                                    data=zip_buffer,
                                    file_name=f"{base_name}_all_languages.zip",
                                    mime="application/zip",
                                    use_container_width=True
                                )

                with tab2:
                    st.markdown("#### 🔍 Form Data Preview")
                    st.dataframe(
                        reader.survey_df[['type', 'name']].head(8), 
                        use_container_width=True, 
                        height=220
                    )

            except Exception as e:
                status.update(label="Conversion Error", state="error")
                st.error(f"Failed to process file: {str(e)}")

else:
    st.info("💡 Upload a `.xlsx` file above to start formatting.")