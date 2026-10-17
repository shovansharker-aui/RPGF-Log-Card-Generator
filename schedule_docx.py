"""
schedule_docx.py

Word version of the weekly Maintenance Schedule: the company header table on
every page, "Page X of Y" in the footer, and the schedule table where every
frequency has a tick box beside it.
"""

import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt, Twips

FONT = "Times New Roman"

HEADERS = ["Sl No", "Equipment Name", "Equipment ID", "Frequency",
           "Schedule Date"]

# Same widths as the PDF (165 mm in total).
COLUMN_MM = [14, 51, 34, 39, 27]

SHADE = "E6E6E6"


def _logo_path():

    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))

    return base / "assets" / "header_logo.png"


def _set_font(run, size, bold=False):

    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold

    properties = run._element.get_or_add_rPr()
    fonts = properties.find(qn("w:rFonts"))

    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        properties.insert(0, fonts)

    for attribute in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(attribute), FONT)


def _tight(paragraph, align=None, after=0):

    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = 1.0

    if align is not None:
        paragraph.alignment = align


def _shade(cell, fill):

    properties = cell._tc.get_or_add_tcPr()

    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), fill)

    properties.append(shading)


def _cell_margins(table, top=40, bottom=40, left=80, right=80):

    properties = table._tbl.tblPr

    margins = OxmlElement("w:tblCellMar")

    for side, value in (("top", top), ("left", left), ("bottom", bottom),
                        ("right", right)):
        element = OxmlElement(f"w:{side}")
        element.set(qn("w:w"), str(value))
        element.set(qn("w:type"), "dxa")
        margins.append(element)

    properties.append(margins)


def _field(paragraph, instruction, size=12):
    """A PAGE / NUMPAGES field."""

    for kind in ("begin", None, "separate", "text", "end"):

        run = paragraph.add_run()

        _set_font(run, size, bold=True)

        if kind in ("begin", "separate", "end"):
            mark = OxmlElement("w:fldChar")
            mark.set(qn("w:fldCharType"), kind)
            run._r.append(mark)

        elif kind is None:
            code = OxmlElement("w:instrText")
            code.set(qn("xml:space"), "preserve")
            code.text = f" {instruction} "
            run._r.append(code)

        else:
            run.text = "1"


DOC_TITLE = "Maintenance Schedule Checklist"


def _build_header(section, doc_title=DOC_TITLE):
    """Company header table, repeated on every page."""

    header = section.header
    header.is_linked_to_previous = False

    holder = header.paragraphs[0]
    _tight(holder)

    table = header.add_table(2, 2, Twips(9351))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # The table must come before the closing paragraph of the header.
    holder._p.addprevious(table._tbl)

    widths = (Twips(1780), Twips(7571))

    for row in table.rows:
        for cell, width in zip(row.cells, widths):
            cell.width = width

    # ---- row 1: logo + company lines
    top = table.rows[0]
    top.height = Twips(907)
    top.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

    left_cell, right_cell = top.cells

    logo_paragraph = left_cell.paragraphs[0]
    _tight(logo_paragraph)
    logo_paragraph.paragraph_format.left_indent = Mm(4)
    logo_paragraph.paragraph_format.space_before = Pt(3)

    if _logo_path().exists():
        logo_paragraph.add_run().add_picture(str(_logo_path()),
                                             width=Cm(1.27))

    lines = (("RENATA PLC", 16, True),
             ("Rajendrapur General Facility", 12, True),
             ("Noyapara, Bhawal Mirzapur, Rajendrapur, Gazipur-1700, "
              "Bangladesh", 10, False))

    for index, (text, size, bold) in enumerate(lines):

        paragraph = (right_cell.paragraphs[0] if index == 0
                     else right_cell.add_paragraph())

        _tight(paragraph, WD_ALIGN_PARAGRAPH.RIGHT)

        _set_font(paragraph.add_run(text), size, bold)

    # ---- row 2: document title
    bottom = table.rows[1]
    bottom.height = Twips(485)
    bottom.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    title_left, title_right = bottom.cells

    for cell in (title_left, title_right):
        properties = cell._tc.get_or_add_tcPr()
        vertical = OxmlElement("w:vAlign")
        vertical.set(qn("w:val"), "center")
        properties.append(vertical)

    paragraph = title_left.paragraphs[0]
    _tight(paragraph)
    _set_font(paragraph.add_run("Doc. Title:"), 14, True)

    paragraph = title_right.paragraphs[0]
    _tight(paragraph, WD_ALIGN_PARAGRAPH.CENTER)
    _set_font(paragraph.add_run(doc_title), 15, True)

    # keep the closing paragraph from adding visible height
    _set_font(holder.add_run(""), 1)


def _build_footer(section):

    footer = section.footer
    footer.is_linked_to_previous = False

    paragraph = footer.paragraphs[0]
    _tight(paragraph, WD_ALIGN_PARAGRAPH.CENTER)

    _set_font(paragraph.add_run("Page "), 12, True)
    _field(paragraph, "PAGE")
    _set_font(paragraph.add_run(" of "), 12, True)
    _field(paragraph, "NUMPAGES")


def _checkbox():
    """A real Word tick-box (content control) the user can click."""

    return parse_xml(
        f'<w:sdt {nsdecls("w")} '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
        '<w:sdtPr>'
        '<w14:checkbox>'
        '<w14:checked w14:val="0"/>'
        '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>'
        '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>'
        '</w14:checkbox>'
        '</w:sdtPr>'
        '<w:sdtContent>'
        '<w:r><w:rPr>'
        '<w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" '
        'w:hAnsi="MS Gothic" w:hint="eastAsia"/>'
        '<w:sz w:val="22"/>'
        '</w:rPr><w:t>☐</w:t></w:r>'
        '</w:sdtContent>'
        '</w:sdt>'
    )


def _write_frequency(cell, frequency):
    """'W, M' -> 'W [ ],  M [ ]' with a tick box beside every frequency."""

    paragraph = cell.paragraphs[0]
    _tight(paragraph, WD_ALIGN_PARAGRAPH.CENTER)

    names = [part.strip() for part in frequency.split(",") if part.strip()]

    for index, name in enumerate(names):

        if index:
            _set_font(paragraph.add_run(",  "), 10)

        # non-breaking space: a name never separates from its tick box
        _set_font(paragraph.add_run(f"{name} "), 10)

        paragraph._p.append(_checkbox())


def _text_cell(cell, text, align, bold=False, size=10):

    paragraph = cell.paragraphs[0]
    _tight(paragraph, align)
    _set_font(paragraph.add_run(text), size, bold)


def write_docx(rows, docx_file, title, doc_title=DOC_TITLE):
    """Create the Word schedule. `rows` are the schedule rows (name, id,
    frequency, date); `title` is the 'Schedule for 3rd Week of ...' line."""

    document = Document()

    section = document.sections[0]

    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = section.right_margin = Mm(22.5)
    section.top_margin = Mm(41)
    section.bottom_margin = Mm(22)
    section.header_distance = Mm(12.7)
    section.footer_distance = Mm(12.7)

    normal = document.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(10)

    _build_header(section, doc_title)
    _build_footer(section)

    heading = document.paragraphs[0] if document.paragraphs else \
        document.add_paragraph()

    _tight(heading, WD_ALIGN_PARAGRAPH.CENTER, after=6)
    _set_font(heading.add_run(title), 12, True)

    table = document.add_table(rows=1, cols=len(HEADERS))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _cell_margins(table)

    widths = [Mm(value) for value in COLUMN_MM]

    for cell, text, width in zip(table.rows[0].cells, HEADERS, widths):
        cell.width = width
        _text_cell(cell, text, WD_ALIGN_PARAGRAPH.CENTER, bold=True)

    # repeat the heading row on every page
    heading_row = table.rows[0]._tr.get_or_add_trPr()
    repeat = OxmlElement("w:tblHeader")
    repeat.set(qn("w:val"), "true")
    heading_row.append(repeat)

    for number, item in enumerate(rows, start=1):

        row = table.add_row()

        no_split = OxmlElement("w:cantSplit")
        row._tr.get_or_add_trPr().append(no_split)

        for cell, width in zip(row.cells, widths):
            cell.width = width

        cells = row.cells

        _text_cell(cells[0], str(number), WD_ALIGN_PARAGRAPH.CENTER)
        _text_cell(cells[1], item["name"], WD_ALIGN_PARAGRAPH.LEFT)
        _text_cell(cells[2], item["id"], WD_ALIGN_PARAGRAPH.LEFT)
        _write_frequency(cells[3], item["frequency"])
        _text_cell(cells[4], item["date"].strftime("%d.%m.%Y"),
                   WD_ALIGN_PARAGRAPH.CENTER)

        if number % 2 == 0:
            for cell in cells:
                _shade(cell, SHADE)

    document.core_properties.title = doc_title
    document.core_properties.author = "RPGF Engineering"

    Path(docx_file).parent.mkdir(parents=True, exist_ok=True)

    document.save(str(docx_file))
