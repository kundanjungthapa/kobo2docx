import re
import pandas as pd

def detect_languages(columns):
    """Finds all columns matching pattern 'label::<Language>'."""
    lang_map = {}
    pattern = re.compile(r"^label::(.+)$", re.IGNORECASE)
    for col in columns:
        match = pattern.match(col.strip())
        if match:
            lang_name = match.group(1).strip()
            lang_map[lang_name] = col
    return lang_map

def get_choices_map(choices_df, label_col):
    """Maps list_name to choices for a specific target language."""
    choices_dict = {}
    if not choices_df.empty and label_col in choices_df.columns:
        for list_name, group in choices_df.groupby('list_name'):
            choices_dict[str(list_name)] = []
            for _, row in group.iterrows():
                val = row.get(label_col)
                if pd.isna(val) or not str(val).strip():
                    val = row.get('name', '')
                choices_dict[str(list_name)].append(str(val).strip())
    return choices_dict

def format_relevance(relevant_str):
    """Formats raw XLSForm skip logic for human reading."""
    if pd.isna(relevant_str) or not str(relevant_str).strip():
        return ""
    return f"[ Skip Logic: Show only if {str(relevant_str).strip()} ]"