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

# ---------------------------------------------------------
# CUSTOM CSS STYLING (Modern Dashboard UI)
# ---------------------------------------------------------
st.markdown("""
<style>
    /* Global font & background tuning */
    .main {
        background-color: #f8fafc;
    }
    
    /* Hero Header Styling */
    .hero-box {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        padding: 2.5rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }
    .hero-title {
        font-size: 2.4rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }
    .hero-subtitle {
        font-size: 1.1rem;
        opacity: 0.9;
    }

    /* Process Card Styling */
    .step-card {
        background: white;
        padding: 1.25rem;
        border-radius: 10px;
        border-left: 4px solid #3b82f6;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 1rem;
    }
    .step-num {
        font-weight: 700;
        color: #3b82f6;
        font-size: 0.9rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .step-title {
        font-size: 1.1rem;
        font-weight: 600;
        color: #1e293b;
        margin-top: 0.2rem;
    }

    /* Feature Badge Cards */
    .feature-badge {
        background: #f1f5f9;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 0.75rem 1rem;
        font-size: 0.9rem;
        color: #334155;
        font-weight: 500;
    }

    /* Hide default Streamlit footer padding */
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# SIDEBAR: DOCUMENTATION & CONFIGURATION
# ---------------------------------------------------------
# ---------------------------------------------------------
# SIDEBAR: DOCUMENTATION & CONFIGURATION
# ---------------------------------------------------------
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/ms-word.png", width=64)
    st.title("Kobo2Docx Engine")
    st.caption("Version 1.2.0 • Production")
    
    # Cross-referencing link to GitHub Pages
    st.markdown("🌐 **Project Site**: [kundanjungthapa.github.io/kobo2docx](https://kundanjungthapa.github.io/kobo2docx/)")
    
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
        **Survey Sheet:**
        - `type`, `name`
        - `label::<lang>` (e.g., `label::English (en)`)
        - `hint::<lang>`
        - `relevant` (Skip logic)
        
        **Choices Sheet:**
        - `list_name`, `name`, `label::<lang>`
        """)
        
    st.info("Developed for KoboToolbox & ODK XLSForm structures.")

# ---------------------------------------------------------
# MAIN LANDING & HERO SECTION
# ---------------------------------------------------------
st.markdown("""
<div class="hero-box">
    <div class="hero-title">📋 Kobo2Docx Questionnaire Generator</div>
    <div class="hero-subtitle">Convert raw KoboToolbox / ODK Excel forms into cleanly styled, publication-ready Word documents in seconds.</div>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# INTERACTIVE WORKFLOW STEPS
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
# CONVERSION CONSOLE (MAIN INTERACTION ZONE)
# ---------------------------------------------------------
st.markdown("### 📂 Conversion Console")

uploaded_file = st.file_uploader(
    "Drag and drop your KoboToolbox file here (.xlsx)", 
    type=["xlsx"],
    help="Supports single or multi-language XLSForms"
)

if uploaded_file is not None:
    # Processing Status Panel
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

                # Tabbed Preview Area
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

                            # Download Section
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
    # Empty State Information Box
    st.info("💡 Upload a `.xlsx` file above to start formatting.")