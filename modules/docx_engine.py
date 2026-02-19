"""DOCX generation engine: clones the Educabot template and injects proposal content."""

import copy
import io
from itertools import groupby

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsdecls
from docx.shared import Pt, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from config import (
    TEMPLATE_PATH, COLOR_PRIMARY, COLOR_CLIENT, COLOR_HEADING3,
    COLOR_TABLE_HEADER, COLOR_TABLE_HEADER_TEXT, COLOR_TABLE_ALT_ROW,
    COLOR_WHITE, FONT_BODY, FONT_HEADING, SIZE_H1, SIZE_H2, SIZE_H3,
    SIZE_BODY, CONTENT_WIDTH, CONTACT_NAME, CONTACT_ROLE, CONTACT_EMAIL,
    CONTACT_WEB,
)
from modules.types import ProposalData, PricingItem


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------

def _rgb(hex_color: str) -> RGBColor:
    return RGBColor(int(hex_color[:2], 16), int(hex_color[2:4], 16), int(hex_color[4:], 16))


def _set_run_font(run, font_name: str, size_half_pts: int, bold: bool = False, color: str | None = None):
    """Configure a run's font: name (all slots), size, bold, color."""
    run.font.name = font_name
    rpr = run._element.get_or_add_rPr()
    # Set eastAsia font
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)
    rfonts.set(qn("w:eastAsia"), font_name)
    rfonts.set(qn("w:cs"), font_name)

    run.font.size = Pt(size_half_pts / 2)
    run.font.bold = bold
    if color:
        run.font.color.rgb = _rgb(color)


def _set_cell_shading(cell, color_hex: str):
    """Set cell background color using OxmlElement (ShadingType.CLEAR pattern)."""
    shading_elm = OxmlElement("w:shd")
    shading_elm.set(qn("w:val"), "clear")
    shading_elm.set(qn("w:color"), "auto")
    shading_elm.set(qn("w:fill"), color_hex)
    cell._tc.get_or_add_tcPr().append(shading_elm)


def _set_cell_width(cell, width_dxa: int):
    """Set explicit cell width in DXA."""
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_width = OxmlElement("w:tcW")
    tc_width.set(qn("w:w"), str(width_dxa))
    tc_width.set(qn("w:type"), "dxa")
    tc_pr.append(tc_width)


def _set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
    """Set cell internal margins in DXA."""
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = OxmlElement("w:tcMar")
    for side, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        margins.append(el)
    tc_pr.append(margins)


def _remove_cell_borders(cell):
    """Remove all borders from a cell."""
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement("w:tcBorders")
    for side in ["top", "left", "bottom", "right"]:
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "none")
        border.set(qn("w:sz"), "0")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "auto")
        tc_borders.append(border)
    tc_pr.append(tc_borders)


def _set_table_borders(table, color="D0D0D0", size="4"):
    """Set thin borders on the entire table."""
    tbl_pr = table._tbl.tblPr
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        table._tbl.insert(0, tbl_pr)
    borders = OxmlElement("w:tblBorders")
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), size)
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), color)
        borders.append(border)
    tbl_pr.append(borders)


# ---------------------------------------------------------------------------
# Document manipulation
# ---------------------------------------------------------------------------

def _clear_body(doc: Document):
    """Remove all paragraphs and tables from the body, preserving <w:sectPr>."""
    body = doc.element.body
    # Collect elements to remove (everything except sectPr)
    to_remove = []
    for child in body:
        if child.tag != qn("w:sectPr"):
            to_remove.append(child)
    for el in to_remove:
        body.remove(el)


def _add_paragraph(doc: Document, text: str, font_name: str = FONT_BODY,
                   size: int = SIZE_BODY, bold: bool = False,
                   color: str | None = None,
                   alignment: WD_ALIGN_PARAGRAPH | None = None,
                   space_before: int | None = None,
                   space_after: int | None = None):
    """Add a formatted paragraph to the document."""
    # Insert before sectPr
    body = doc.element.body
    sect_pr = body.find(qn("w:sectPr"))

    p = doc.add_paragraph()
    # Move paragraph before sectPr
    if sect_pr is not None:
        body.remove(p._element)
        sect_pr.addprevious(p._element)

    run = p.add_run(text)
    _set_run_font(run, font_name, size, bold, color)

    if alignment is not None:
        p.alignment = alignment
    if space_before is not None:
        p.paragraph_format.space_before = Pt(space_before)
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)

    return p


def _setup_heading_numbering(doc: Document) -> int:
    """Create a multilevel numbering definition for section headings.

    Level 0 (ilvl 0) → H2 sections: renders "1. ", "2. ", "3. "
    Level 1 (ilvl 1) → H3 sub-sections: renders "1.1 ", "1.2 ", "2.1 "

    Returns the numId to reference in heading paragraphs via w:numPr.
    """
    numbering_part = doc.part.numbering_part
    numbering_elm = numbering_part._element

    existing = numbering_elm.findall(qn("w:abstractNum"))
    abstract_id = len(existing) + 200
    num_id = abstract_id

    abstract_num = OxmlElement("w:abstractNum")
    abstract_num.set(qn("w:abstractNumId"), str(abstract_id))

    # --- Level 0: H2 sections → "1." "2." ---
    lvl0 = OxmlElement("w:lvl")
    lvl0.set(qn("w:ilvl"), "0")

    start0 = OxmlElement("w:start")
    start0.set(qn("w:val"), "1")
    lvl0.append(start0)

    numFmt0 = OxmlElement("w:numFmt")
    numFmt0.set(qn("w:val"), "decimal")
    lvl0.append(numFmt0)

    lvlText0 = OxmlElement("w:lvlText")
    lvlText0.set(qn("w:val"), "%1.")
    lvl0.append(lvlText0)

    lvlJc0 = OxmlElement("w:lvlJc")
    lvlJc0.set(qn("w:val"), "left")
    lvl0.append(lvlJc0)

    suff0 = OxmlElement("w:suff")
    suff0.set(qn("w:val"), "space")
    lvl0.append(suff0)

    pPr0 = OxmlElement("w:pPr")
    ind0 = OxmlElement("w:ind")
    ind0.set(qn("w:left"), "0")
    ind0.set(qn("w:firstLine"), "0")
    pPr0.append(ind0)
    lvl0.append(pPr0)

    # Run properties for the number character itself (font, color, bold, size)
    rPr0 = OxmlElement("w:rPr")
    rFonts0 = OxmlElement("w:rFonts")
    rFonts0.set(qn("w:ascii"), FONT_HEADING)
    rFonts0.set(qn("w:hAnsi"), FONT_HEADING)
    rFonts0.set(qn("w:eastAsia"), FONT_HEADING)
    rFonts0.set(qn("w:cs"), FONT_HEADING)
    rPr0.append(rFonts0)
    bold0 = OxmlElement("w:b")
    rPr0.append(bold0)
    color0 = OxmlElement("w:color")
    color0.set(qn("w:val"), COLOR_PRIMARY)
    rPr0.append(color0)
    sz0 = OxmlElement("w:sz")
    sz0.set(qn("w:val"), str(SIZE_H2))
    rPr0.append(sz0)
    szCs0 = OxmlElement("w:szCs")
    szCs0.set(qn("w:val"), str(SIZE_H2))
    rPr0.append(szCs0)
    lvl0.append(rPr0)

    abstract_num.append(lvl0)

    # --- Level 1: H3 sub-sections → "1.1" "1.2" ---
    lvl1 = OxmlElement("w:lvl")
    lvl1.set(qn("w:ilvl"), "1")

    start1 = OxmlElement("w:start")
    start1.set(qn("w:val"), "1")
    lvl1.append(start1)

    numFmt1 = OxmlElement("w:numFmt")
    numFmt1.set(qn("w:val"), "decimal")
    lvl1.append(numFmt1)

    lvlText1 = OxmlElement("w:lvlText")
    lvlText1.set(qn("w:val"), "%1.%2")
    lvl1.append(lvlText1)

    lvlJc1 = OxmlElement("w:lvlJc")
    lvlJc1.set(qn("w:val"), "left")
    lvl1.append(lvlJc1)

    suff1 = OxmlElement("w:suff")
    suff1.set(qn("w:val"), "space")
    lvl1.append(suff1)

    pPr1 = OxmlElement("w:pPr")
    ind1 = OxmlElement("w:ind")
    ind1.set(qn("w:left"), "0")
    ind1.set(qn("w:firstLine"), "0")
    pPr1.append(ind1)
    lvl1.append(pPr1)

    # Run properties for H3 number character
    rPr1 = OxmlElement("w:rPr")
    rFonts1 = OxmlElement("w:rFonts")
    rFonts1.set(qn("w:ascii"), FONT_HEADING)
    rFonts1.set(qn("w:hAnsi"), FONT_HEADING)
    rFonts1.set(qn("w:eastAsia"), FONT_HEADING)
    rFonts1.set(qn("w:cs"), FONT_HEADING)
    rPr1.append(rFonts1)
    color1 = OxmlElement("w:color")
    color1.set(qn("w:val"), COLOR_HEADING3)
    rPr1.append(color1)
    sz1 = OxmlElement("w:sz")
    sz1.set(qn("w:val"), str(SIZE_H3))
    rPr1.append(sz1)
    szCs1 = OxmlElement("w:szCs")
    szCs1.set(qn("w:val"), str(SIZE_H3))
    rPr1.append(szCs1)
    lvl1.append(rPr1)

    abstract_num.append(lvl1)

    numbering_elm.append(abstract_num)

    # Concrete numbering instance
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abstract_ref = OxmlElement("w:abstractNumId")
    abstract_ref.set(qn("w:val"), str(abstract_id))
    num.append(abstract_ref)
    numbering_elm.append(num)

    return num_id


def _add_heading(doc: Document, text: str, level: int = 1,
                 num_id: int | None = None):
    """Add a heading paragraph using native Word heading styles.

    Uses built-in 'Heading 1/2/3' styles so Google Docs recognises them.
    When num_id is provided, attaches native Word numbering (w:numPr)
    so the numbers are generated automatically by Word/Google Docs.
    """
    style_config = {
        1: (FONT_BODY, SIZE_H1, True, COLOR_PRIMARY, 18, 8),
        2: (FONT_HEADING, SIZE_H2, True, COLOR_PRIMARY, 14, 6),
        3: (FONT_HEADING, SIZE_H3, False, COLOR_HEADING3, 10, 4),
    }
    font, size, bold, color, before, after = style_config.get(level, style_config[3])

    body = doc.element.body
    sect_pr = body.find(qn("w:sectPr"))

    p = doc.add_heading(text, level=level)
    if sect_pr is not None:
        body.remove(p._element)
        sect_pr.addprevious(p._element)

    for run in p.runs:
        _set_run_font(run, font, size, bold, color)

    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)

    # Attach native numbering so Word/Google Docs generate "1.", "1.1" etc.
    if num_id is not None and level in (2, 3):
        ilvl_val = 0 if level == 2 else 1
        pPr = p._element.get_or_add_pPr()
        numPr = OxmlElement("w:numPr")
        ilvl_el = OxmlElement("w:ilvl")
        ilvl_el.set(qn("w:val"), str(ilvl_val))
        numPr.append(ilvl_el)
        numId_el = OxmlElement("w:numId")
        numId_el.set(qn("w:val"), str(num_id))
        numPr.append(numId_el)
        pPr.append(numPr)

    return p


def _add_body_text(doc: Document, text: str, space_after: int = 6):
    """Add body text paragraph."""
    return _add_paragraph(doc, text, FONT_BODY, SIZE_BODY, space_after=space_after,
                          alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)


def _add_empty_line(doc: Document):
    """Add an empty paragraph as spacer."""
    return _add_paragraph(doc, "", size=SIZE_BODY, space_after=0)


def _add_divider(doc: Document):
    """Add a horizontal line divider."""
    p = _add_paragraph(doc, "", space_before=6, space_after=6)
    pPr = p._element.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "A0A0A0")
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


# ---------------------------------------------------------------------------
# Metadata block
# ---------------------------------------------------------------------------

def _add_metadata_block(doc: Document, data: ProposalData):
    """Add the initial metadata block: date, ref, title, client."""
    _add_paragraph(doc, data.date, FONT_BODY, SIZE_BODY,
                   alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_after=6)
    _add_paragraph(doc, f"Ref: {data.ref_number}", FONT_BODY, SIZE_BODY,
                   bold=True, space_after=6)
    _add_empty_line(doc)
    _add_heading(doc, f"Proyecto: {data.project_title}", level=1)
    _add_paragraph(doc, data.client_name, FONT_HEADING, SIZE_H3,
                   color=COLOR_CLIENT, space_after=8)


# ---------------------------------------------------------------------------
# Pricing table
# ---------------------------------------------------------------------------

def _add_table_before_sectpr(doc: Document, rows: int, cols: int):
    """Add a table to the document body, before the sectPr element."""
    table = doc.add_table(rows=rows, cols=cols)
    body = doc.element.body
    sect_pr = body.find(qn("w:sectPr"))
    if sect_pr is not None:
        body.remove(table._tbl)
        sect_pr.addprevious(table._tbl)
    return table


def _create_pricing_table(doc: Document, items: list[PricingItem]):
    """Create a pricing table with header, grouped items, subtotals, and total."""
    if not items:
        return

    # Group items by category
    sorted_items = sorted(items, key=lambda x: x.category)
    groups = [(cat, list(group_items)) for cat, group_items in groupby(sorted_items, key=lambda x: x.category)]

    # Calculate total rows: header + items + category subtotals + total
    n_rows = 1 + len(items) + len(groups) + 1
    cols = 5  # Tipo | Producto | Cantidad | Valor Unitario | Sub-Total

    table = _add_table_before_sectpr(doc, n_rows, cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Column widths (DXA) - must sum to CONTENT_WIDTH (9029)
    col_widths = [1500, 3129, 1100, 1400, 1900]

    # Set table width
    tbl_pr = table._tbl.tblPr
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        table._tbl.insert(0, tbl_pr)
    tbl_width = OxmlElement("w:tblW")
    tbl_width.set(qn("w:w"), str(CONTENT_WIDTH))
    tbl_width.set(qn("w:type"), "dxa")
    tbl_pr.append(tbl_width)

    _set_table_borders(table, color="E0E0E0", size="4")

    # Header row
    headers = ["Tipo", "Producto", "Cantidad", "Valor Unit. (USD)", "Sub-Total (USD)"]
    header_row = table.rows[0]
    for i, (header_text, width) in enumerate(zip(headers, col_widths)):
        cell = header_row.cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(header_text)
        _set_run_font(run, FONT_HEADING, SIZE_BODY, bold=True, color=COLOR_WHITE)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i >= 2 else WD_ALIGN_PARAGRAPH.LEFT
        _set_cell_shading(cell, COLOR_TABLE_HEADER)
        _set_cell_width(cell, width)
        _set_cell_margins(cell)

    # Data rows
    row_idx = 1
    for cat_idx, (category, group_items) in enumerate(groups):
        for item_idx, item in enumerate(group_items):
            row = table.rows[row_idx]
            values = [
                item.category,
                item.product,
                str(item.quantity),
                f"${item.unit_price:,.2f}",
                f"${item.subtotal:,.2f}",
            ]
            aligns = [
                WD_ALIGN_PARAGRAPH.LEFT,
                WD_ALIGN_PARAGRAPH.LEFT,
                WD_ALIGN_PARAGRAPH.CENTER,
                WD_ALIGN_PARAGRAPH.RIGHT,
                WD_ALIGN_PARAGRAPH.RIGHT,
            ]

            # Alternate row shading
            use_shading = (row_idx % 2 == 0)

            for i, (val, width, align) in enumerate(zip(values, col_widths, aligns)):
                cell = row.cells[i]
                cell.text = ""
                p = cell.paragraphs[0]
                run = p.add_run(val)
                _set_run_font(run, FONT_BODY, SIZE_BODY)
                p.alignment = align
                _set_cell_width(cell, width)
                _set_cell_margins(cell)
                if use_shading:
                    _set_cell_shading(cell, COLOR_TABLE_ALT_ROW)

            row_idx += 1

        # Category subtotal row
        cat_subtotal = sum(it.subtotal for it in group_items)
        sub_row = table.rows[row_idx]
        for i, width in enumerate(col_widths):
            cell = sub_row.cells[i]
            cell.text = ""
            _set_cell_width(cell, width)
            _set_cell_margins(cell)
            _set_cell_shading(cell, "EAEAEA")

        # Merge first 4 cells for subtotal label
        merge_cell = sub_row.cells[0]
        p = merge_cell.paragraphs[0]
        run = p.add_run(f"Subtotal {category}")
        _set_run_font(run, FONT_HEADING, SIZE_BODY, bold=True)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        total_cell = sub_row.cells[4]
        p = total_cell.paragraphs[0]
        run = p.add_run(f"${cat_subtotal:,.2f}")
        _set_run_font(run, FONT_HEADING, SIZE_BODY, bold=True)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        row_idx += 1

    # Grand total row
    total_row = table.rows[row_idx]
    for i, width in enumerate(col_widths):
        cell = total_row.cells[i]
        cell.text = ""
        _set_cell_width(cell, width)
        _set_cell_margins(cell)
        _set_cell_shading(cell, COLOR_TABLE_HEADER)

    total_label_cell = total_row.cells[0]
    p = total_label_cell.paragraphs[0]
    run = p.add_run("TOTAL")
    _set_run_font(run, FONT_HEADING, SIZE_BODY, bold=True, color=COLOR_WHITE)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    total_value_cell = total_row.cells[4]
    p = total_value_cell.paragraphs[0]
    run = p.add_run(f"${sum(it.subtotal for it in items):,.2f}")
    _set_run_font(run, FONT_HEADING, SIZE_BODY, bold=True, color=COLOR_WHITE)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    _add_empty_line(doc)


# ---------------------------------------------------------------------------
# Commercial terms
# ---------------------------------------------------------------------------

def _ensure_bullet_numbering(doc: Document) -> int:
    """Ensure a bullet numbering definition exists in the document. Returns the numId."""
    numbering_part = doc.part.numbering_part
    numbering_elm = numbering_part._element

    # Check if we already added our bullet definition
    existing = numbering_elm.findall(qn("w:abstractNum"))
    abstract_id = len(existing) + 100  # Use high ID to avoid conflicts
    num_id = abstract_id

    # Create abstract numbering definition for bullets
    abstract_num = OxmlElement("w:abstractNum")
    abstract_num.set(qn("w:abstractNumId"), str(abstract_id))

    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    lvl.append(start)
    num_fmt = OxmlElement("w:numFmt")
    num_fmt.set(qn("w:val"), "bullet")
    lvl.append(num_fmt)
    lvl_text = OxmlElement("w:lvlText")
    lvl_text.set(qn("w:val"), "\u2022")
    lvl.append(lvl_text)
    lvl_jc = OxmlElement("w:lvlJc")
    lvl_jc.set(qn("w:val"), "left")
    lvl.append(lvl_jc)

    pPr = OxmlElement("w:pPr")
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "720")
    ind.set(qn("w:hanging"), "360")
    pPr.append(ind)
    lvl.append(pPr)

    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), "Symbol")
    rFonts.set(qn("w:hAnsi"), "Symbol")
    rFonts.set(qn("w:hint"), "default")
    rPr.append(rFonts)
    lvl.append(rPr)

    abstract_num.append(lvl)
    numbering_elm.append(abstract_num)

    # Create concrete numbering instance
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abstract_ref = OxmlElement("w:abstractNumId")
    abstract_ref.set(qn("w:val"), str(abstract_id))
    num.append(abstract_ref)
    numbering_elm.append(num)

    return num_id


def _add_commercial_terms(doc: Document, terms: list[str]):
    """Add commercial terms as a bulleted list."""
    if not terms:
        return

    num_id = _ensure_bullet_numbering(doc)

    body = doc.element.body
    sect_pr = body.find(qn("w:sectPr"))

    for term in terms:
        p = doc.add_paragraph()
        if sect_pr is not None:
            body.remove(p._element)
            sect_pr.addprevious(p._element)

        # Set bullet numbering via XML
        pPr = p._element.get_or_add_pPr()
        numPr = OxmlElement("w:numPr")
        ilvl = OxmlElement("w:ilvl")
        ilvl.set(qn("w:val"), "0")
        numPr.append(ilvl)
        numId_el = OxmlElement("w:numId")
        numId_el.set(qn("w:val"), str(num_id))
        numPr.append(numId_el)
        pPr.append(numPr)

        run = p.add_run(term)
        _set_run_font(run, FONT_BODY, SIZE_BODY)

    _add_empty_line(doc)


# ---------------------------------------------------------------------------
# Multi-paragraph content sections
# ---------------------------------------------------------------------------

def _add_content_section(doc: Document, title: str, content: str, level: int = 2,
                         num_id: int | None = None):
    """Add a section with title and multi-paragraph content.

    When num_id is provided, headings get native Word numbering via w:numPr.
    Sub-sections (lines ending with ':') become H3 with automatic sub-numbering.
    Lines starting with '- ' are rendered as indented bullets.
    """
    if not content:
        return

    _add_heading(doc, title, level=level, num_id=num_id)

    for paragraph_text in content.split("\n\n"):
        text = paragraph_text.strip()
        if not text:
            continue

        lines = text.split("\n")
        for line in lines:
            line = line.strip()
            if not line:
                continue

            if line.startswith("- "):
                # Bullet item — indented body text with bullet char
                p = _add_paragraph(doc, f"\u2022  {line[2:]}", FONT_BODY, SIZE_BODY,
                                   space_after=4, space_before=0)
                pPr = p._element.get_or_add_pPr()
                ind = OxmlElement("w:ind")
                ind.set(qn("w:left"), "720")
                ind.set(qn("w:hanging"), "180")
                pPr.append(ind)
            elif line.endswith(":") and num_id:
                # Sub-heading — Word auto-numbers as N.1, N.2 etc.
                _add_heading(doc, line[:-1], level=3, num_id=num_id)
            else:
                _add_body_text(doc, line)

    _add_empty_line(doc)


# ---------------------------------------------------------------------------
# Contact block
# ---------------------------------------------------------------------------

def _add_contact_block(doc: Document):
    """Add contact information at the end of the document."""
    _add_empty_line(doc)
    _add_paragraph(doc, CONTACT_NAME, FONT_BODY, SIZE_BODY, bold=True, space_after=2)
    _add_paragraph(doc, CONTACT_ROLE, FONT_BODY, SIZE_BODY, space_after=2)
    _add_paragraph(doc, CONTACT_EMAIL, FONT_BODY, SIZE_BODY, space_after=2)
    _add_paragraph(doc, CONTACT_WEB, FONT_BODY, SIZE_BODY, space_after=2)


# ---------------------------------------------------------------------------
# Main generator
# ---------------------------------------------------------------------------

def generate_proposal(data: ProposalData) -> bytes:
    """Generate a proposal .docx by cloning the template and injecting content.

    Returns the document as bytes ready for download.
    """
    doc = Document(str(TEMPLATE_PATH))

    # Step 1: Clear body (preserve sectPr with header/footer refs)
    _clear_body(doc)

    # Step 2: Add metadata block
    _add_metadata_block(doc, data)

    # Step 2b: Setup native heading numbering (Word generates 1., 2., 1.1, etc.)
    heading_num_id = _setup_heading_numbering(doc)

    if data.is_formal:
        # Formal proposal: full document
        _add_content_section(doc, "Introducción", data.content.introduction,
                             num_id=heading_num_id)

        _add_content_section(doc, "Presentación de la Empresa",
                             data.content.company_presentation,
                             num_id=heading_num_id)

        if data.content.technical_proposal:
            _add_content_section(doc, "Propuesta Técnica",
                                 data.content.technical_proposal,
                                 num_id=heading_num_id)

        if data.content.deliverables:
            _add_content_section(doc, "Entregables", data.content.deliverables,
                                 num_id=heading_num_id)
    else:
        # Informal: just a brief intro
        if data.content.introduction:
            _add_content_section(doc, "Introducción", data.content.introduction,
                                 num_id=heading_num_id)

    # Step 3: Pricing table
    if data.pricing_items:
        _add_heading(doc, "Propuesta Comercial", level=2, num_id=heading_num_id)
        _create_pricing_table(doc, data.pricing_items)

    # Step 4: Commercial terms
    if data.commercial_terms:
        _add_heading(doc, "Condiciones Comerciales", level=2, num_id=heading_num_id)
        _add_commercial_terms(doc, data.commercial_terms)

    # Step 5: Conclusion
    if data.content.conclusion:
        _add_content_section(doc, "Conclusión", data.content.conclusion,
                             num_id=heading_num_id)

    # Step 6: Contact block
    _add_contact_block(doc)

    # Serialize to bytes
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
