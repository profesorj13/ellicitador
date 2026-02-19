"""Design tokens and constants extracted from the Educabot template."""

from pathlib import Path

# Paths
ASSETS_DIR = Path(__file__).parent / "assets"
TEMPLATE_PATH = ASSETS_DIR / "template.docx"
KNOWLEDGE_DIR = Path(__file__).parent / "knowledge"
DATA_DIR = Path(__file__).parent / "data"

# Colors (hex without #)
COLOR_PRIMARY = "2a205e"       # Deep blue-purple (headings, footer)
COLOR_SECONDARY = "A0A0A0"    # Gray (divider lines)
COLOR_CLIENT = "4A90D9"        # Light blue (client name)
COLOR_HEADING3 = "434343"      # Dark gray
COLOR_TABLE_HEADER = "2a205e"  # Same as primary
COLOR_TABLE_HEADER_TEXT = "FFFFFF"
COLOR_TABLE_ALT_ROW = "f2f0f8" # Light purple
COLOR_WHITE = "FFFFFF"

# Fonts
FONT_BODY = "DM Sans"
FONT_HEADING = "DM Sans SemiBold"

# Sizes (in half-points: multiply pt by 2)
SIZE_H1 = 40    # 20pt
SIZE_H2 = 28    # 14pt
SIZE_H3 = 28    # 14pt
SIZE_BODY = 22  # 11pt
SIZE_FOOTER = 20  # 10pt
SIZE_SMALL = 18  # 9pt

# Page dimensions (DXA: 1440 = 1 inch)
PAGE_WIDTH = 11909   # A4 width
PAGE_HEIGHT = 16834  # A4 height
MARGIN = 1440        # 1 inch margins all around
CONTENT_WIDTH = PAGE_WIDTH - (2 * MARGIN)  # 9029 DXA

# Reference number format
REF_PREFIX = "EDU-PRO"

# Contact info
CONTACT_NAME = "Jorge Frisancho"
CONTACT_ROLE = "Director Comercial"
CONTACT_EMAIL = "info@educabot.com"
CONTACT_WEB = "www.educabot.com"
