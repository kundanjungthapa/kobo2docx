# KoboToolbox XLSForm Technical Specification

Kobo2Docx parses standard ODK / KoboToolbox survey structures.

## Supported Sheets & Columns

### 1. `survey` Sheet
- `type`: `select_one <list_name>`, `select_multiple <list_name>`, `text`, `integer`, `decimal`, `date`, `time`, `note`, `begin_group`, `end_group`, `begin_repeat`, `end_repeat`.
- `name`: Technical identifier for variables.
- `label::<Language>`: Human-readable text (e.g., `label::English (en)`, `label::Nepali (ne)`).
- `hint::<Language>`: Instructional hint text displayed beneath the question.
- `relevant`: Skip-logic expressions. Triggers 10 pt fainter font and indent formatting.

### 2. `choices` Sheet
- `list_name`: Key corresponding to `select_one` or `select_multiple` types.
- `name`: Choice item code.
- `label::<Language>`: Choice item label text per language.

### 3. `settings` Sheet
- `form_title`: Title used for the document header.