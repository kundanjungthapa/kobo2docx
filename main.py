import os
import re
from src.reader import KoboReader
from src.translator import detect_languages, get_choices_map
from src.docx_builder import create_document

INPUT_DIR = "inputs"
OUTPUT_DIR = "outputs"

def main():
    if not os.path.exists(INPUT_DIR):
        os.makedirs(INPUT_DIR)
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    xlsx_files = [f for f in os.listdir(INPUT_DIR) if f.endswith('.xlsx') and not f.startswith('~$')]

    if not xlsx_files:
        print(f"No .xlsx files found in '{INPUT_DIR}/'. Place your Kobo XLSX file there and run again.")
        return

    for file_name in xlsx_files:
        input_path = os.path.join(INPUT_DIR, file_name)
        print(f"\nProcessing: {file_name}")

        reader = KoboReader(input_path)
        title = reader.get_title()
        detected_languages = detect_languages(reader.survey_df.columns)

        if not detected_languages:
            detected_languages['Default'] = 'label'

        base_name = os.path.splitext(file_name)[0]

        for lang_name, label_col in detected_languages.items():
            clean_lang = re.sub(r'[^a-zA-Z0-9]', '_', lang_name).strip('_')
            out_filename = f"{base_name}_{clean_lang}.docx"
            out_path = os.path.join(OUTPUT_DIR, out_filename)

            choices_dict = get_choices_map(reader.choices_df, label_col)
            create_document(lang_name, label_col, reader.survey_df, choices_dict, title, out_path)

if __name__ == "__main__":
    main()