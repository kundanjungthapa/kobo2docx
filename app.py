import os
import io
import re
import tempfile
import zipfile
import streamlit as st

from src.reader import KoboReader
from src.translator import detect_languages, get_choices_map
from src.docx_builder import create_document

# Streamlit Page Setup
st.set_page_config(
    page_title="Kobo2Docx Converter",
    page_icon="📝",
    layout="centered"
)

st.title("📝 KoboToolbox to Word (.docx) Converter")
st.write(
    "Upload your KoboToolbox / ODK Excel questionnaire (`.xlsx`) below. "
    "The tool will automatically parse all languages and provide downloadable Word documents."
)

# File Upload Widget
uploaded_file = st.file_uploader("Upload XLSForm File (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.success(f"File uploaded successfully: **{uploaded_file.name}**")
    
    if st.button("Convert to Word (.docx)"):
        with st.spinner("Processing questionnaire and generating documents..."):
            # Create a temporary directory on the server
            with tempfile.TemporaryDirectory() as temp_dir:
                input_path = os.path.join(temp_dir, uploaded_file.name)
                
                # Write uploaded bytes to disk in temporary directory
                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                try:
                    # Parse Kobo Form
                    reader = KoboReader(input_path)
                    title = reader.get_title()
                    detected_languages = detect_languages(reader.survey_df.columns)

                    if not detected_languages:
                        detected_languages['Default'] = 'label'

                    generated_files = []
                    base_name = os.path.splitext(uploaded_file.name)[0]

                    # Process each language
                    for lang_name, label_col in detected_languages.items():
                        clean_lang = re.sub(r'[^a-zA-Z0-9]', '_', lang_name).strip('_')
                        out_filename = f"{base_name}_{clean_lang}.docx"
                        out_path = os.path.join(temp_dir, out_filename)

                        choices_dict = get_choices_map(reader.choices_df, label_col)
                        create_document(lang_name, label_col, reader.survey_df, choices_dict, title, out_path)
                        generated_files.append((out_filename, out_path))

                    # Single language file download button
                    if len(generated_files) == 1:
                        filename, filepath = generated_files[0]
                        with open(filepath, "rb") as f:
                            st.download_button(
                                label=f"📥 Download {filename}",
                                data=f.read(),
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                            
                    # Multi-language ZIP download button
                    elif len(generated_files) > 1:
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                            for filename, filepath in generated_files:
                                zip_file.write(filepath, arcname=filename)
                        zip_buffer.seek(0)

                        zip_name = f"{base_name}_word_docs.zip"
                        st.download_button(
                            label="📥 Download All Languages (ZIP Archive)",
                            data=zip_buffer,
                            file_name=zip_name,
                            mime="application/zip"
                        )

                except Exception as e:
                    st.error(f"An error occurred while processing the file: {e}")