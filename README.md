# 📋 Kobo2Docx: KoboToolbox / ODK to Word Questionnaire Generator

**Kobo2Docx** is a Python tool that converts KoboToolbox, ODK, or XLSForm Excel questionnaires (`.xlsx`) into publication-ready, structured Word documents (`.docx`).

It automatically scans multi-language forms, parses group logic, applies visual styling for skip logic, and formats responses for enumerator and review use.

---

## ✨ Features

- **🌐 Automatic Multi-Language Detection**: Scans for `label::<language>` columns (e.g., English, Nepali, French) and generates a separate, fully mapped `.docx` file for every language.
- **⚡ Batch Processing**: Converts multiple `.xlsx` questionnaires placed in the `inputs/` folder simultaneously.
- **📑 Group & Section Banners**: Converts `begin_group` / `begin_repeat` blocks into full-width shaded section headers (`SECTION: NAME`).
- **🎯 Visual Skip-Logic Styling**: Questions dependent on skip conditions (`relevant`) are automatically indented and formatted in **10 pt fainter text**, making dependencies instantly recognizable.
- **🖊️ Clean Input Prompts**: Replaces generic text fields with clean dotted lines (`............`) and choice options with bulleted checkboxes (`□ Option`).
- **🛡️ Robust Fallbacks**: Gracefully handles missing optional headers (e.g., `repeat_count`, `hints`, `constraint_message`) without throwing errors.

---

## 🛠️ Folder Structure

```text
kobo2docx/
├── config/             # Styling configuration, font hierarchy, and color palettes
├── src/                # Parser modules (reader, language detector, docx builder)
├── inputs/             # Place your source Kobo .xlsx questionnaire files here
├── outputs/            # Converted .docx questionnaires are generated here
├── main.py             # Primary batch execution script
└── requirements.txt    # Project dependencies (pandas, openpyxl, python-docx)
