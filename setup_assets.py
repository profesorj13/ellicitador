"""One-time setup: extract logo, fonts, and clean template from the original .docx."""

import shutil
import zipfile
from pathlib import Path

TEMPLATE_SOURCE = Path("Referencias/V1 _ Hoja Membretada Educabot 2025.docx")
ASSETS_DIR = Path("assets")
LOGO_PATH = ASSETS_DIR / "logo.png"
FONTS_DIR = ASSETS_DIR / "fonts"
TEMPLATE_PATH = ASSETS_DIR / "template.docx"


def setup():
    ASSETS_DIR.mkdir(exist_ok=True)
    FONTS_DIR.mkdir(exist_ok=True)

    if not TEMPLATE_SOURCE.exists():
        raise FileNotFoundError(f"Template not found: {TEMPLATE_SOURCE}")

    with zipfile.ZipFile(TEMPLATE_SOURCE, "r") as z:
        # Extract logo
        z.extract("word/media/image1.png", ASSETS_DIR)
        extracted_logo = ASSETS_DIR / "word" / "media" / "image1.png"
        shutil.move(str(extracted_logo), str(LOGO_PATH))

        # Extract fonts
        font_files = [n for n in z.namelist() if n.startswith("word/fonts/") and n.endswith(".ttf")]
        for font_file in font_files:
            z.extract(font_file, ASSETS_DIR)
            extracted = ASSETS_DIR / font_file
            shutil.move(str(extracted), str(FONTS_DIR / extracted.name))

        # Clean up extracted directory structure
        word_dir = ASSETS_DIR / "word"
        if word_dir.exists():
            shutil.rmtree(word_dir)

    # Copy template as-is
    shutil.copy2(TEMPLATE_SOURCE, TEMPLATE_PATH)

    print(f"Logo:     {LOGO_PATH} ({LOGO_PATH.stat().st_size} bytes)")
    print(f"Fonts:    {list(FONTS_DIR.glob('*.ttf'))}")
    print(f"Template: {TEMPLATE_PATH} ({TEMPLATE_PATH.stat().st_size} bytes)")
    print("Setup complete.")


if __name__ == "__main__":
    setup()
