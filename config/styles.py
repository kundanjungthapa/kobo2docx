from docx.shared import Inches, Pt, RGBColor

# Margins
PAGE_MARGIN = Inches(0.75)

# Column Widths (S.N., Question Details, Options)
COL_WIDTHS = [Inches(0.6), Inches(4.3), Inches(2.1)]

# Color Palette (Hex & RGB)
COLOR_PRIMARY_HEX = "2E5B82"     # Deep Blue (Header)
COLOR_GROUP_BG_HEX = "F0F0F0"    # Light Gray (Section Rows)

# Text Color Hierarchy
COLOR_MAIN_TEXT_RGB = RGBColor(0, 0, 0)         # Standard Black
COLOR_SKIP_QUESTION_RGB = RGBColor(90, 90, 90)  # Fainter Gray for Skip-Logic Questions
COLOR_HINT_RGB = RGBColor(120, 120, 120)        # Muted Hint Text
COLOR_HEADER_TEXT_RGB = RGBColor(255, 255, 255)
COLOR_GROUP_TEXT_RGB = RGBColor(30, 30, 30)

# Typography Sizes
FONT_SIZE_MAIN = Pt(11)
FONT_SIZE_SKIP = Pt(10)
FONT_SIZE_HINT = Pt(9)
FONT_NAME = "Cambria"