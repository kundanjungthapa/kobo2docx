import os
import pandas as pd

class KoboReader:
    def __init__(self, file_path):
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Input file not found: {file_path}")
            
        self.xls = pd.ExcelFile(file_path)
        self.survey_df = pd.read_excel(self.xls, 'survey')
        self.choices_df = pd.read_excel(self.xls, 'choices') if 'choices' in self.xls.sheet_names else pd.DataFrame()
        self.settings_df = pd.read_excel(self.xls, 'settings') if 'settings' in self.xls.sheet_names else pd.DataFrame()

    def get_title(self):
        if 'form_title' in self.settings_df.columns and not self.settings_df.empty:
            return self.settings_df['form_title'].iloc[0]
        return "Survey Questionnaire"