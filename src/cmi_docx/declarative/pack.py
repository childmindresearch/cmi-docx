"""Packing logic to convert declarative components to python-docx objects."""

import pathlib
from typing import TYPE_CHECKING, Any

import docx
from docx.enum import text as docx_text
from docx.enum.section import WD_ORIENTATION
from docx.shared import Pt, RGBColor

if TYPE_CHECKING:
    from cmi_docx.declarative.image import ImageRun
    from cmi_docx.declarative.paragraph import Paragraph, TextRun
    from cmi_docx.declarative.section import Footer, Header
    from cmi_docx.declarative.table import Table, TableCell, TableRow


def pack(doc: Any) -> Any:
    """Convert a declarative Document into a python-docx Document.

    Args:
        doc: The declarative Document to convert.

    Returns:
        A python-docx Document ready to be saved.
    """
    docx_doc = docx.Document()

    if doc.creator:
        docx_doc.core_properties.author = doc.creator
    if doc.title:
        docx_doc.core_properties.title = doc.title
    if doc.subject:
        docx_doc.core_properties.subject = doc.subject
    if doc.description:
        docx_doc.core_properties.comments = doc.description
    if doc.keywords:
        docx_doc.core_properties.keywords = doc.keywords
    if doc.category:
        docx_doc.core_properties.category = doc.category

    for section in doc.sections:
        _pack_section(docx_doc, section)

    return docx_doc


def _pack_section(docx_doc: Any, section: Any) -> None:
    """Pack a Section into a python-docx document.

    Args:
        docx_doc: The python-docx Document.
        section: The declarative Section.
    """
    if section.children:
        for child in section.children:
            _pack_block_element(docx_doc, child)

    docx_section = docx_doc.sections[-1]

    if section.properties:
        props = section.properties
        if props.page_size:
            if "width" in props.page_size:
                docx_section.page_width = props.page_size["width"]
            if "height" in props.page_size:
                docx_section.page_height = props.page_size["height"]

        if props.page_margins:
            if "top" in props.page_margins:
                docx_section.top_margin = props.page_margins["top"]
            if "bottom" in props.page_margins:
                docx_section.bottom_margin = props.page_margins["bottom"]
            if "left" in props.page_margins:
                docx_section.left_margin = props.page_margins["left"]
            if "right" in props.page_margins:
                docx_section.right_margin = props.page_margins["right"]
            if "header" in props.page_margins:
                docx_section.header_distance = props.page_margins["header"]
            if "footer" in props.page_margins:
                docx_section.footer_distance = props.page_margins["footer"]
            if "gutter" in props.page_margins:
                docx_section.gutter = props.page_margins["gutter"]

        if props.page_orientation:
            if props.page_orientation.lower() == "landscape":
                docx_section.orientation = WD_ORIENTATION.LANDSCAPE
            elif props.page_orientation.lower() == "portrait":
                docx_section.orientation = WD_ORIENTATION.PORTRAIT

    if section.headers:
        for header_type, header in section.headers.items():
            _pack_header(docx_section, header_type, header)

    if section.footers:
        for footer_type, footer in section.footers.items():
            _pack_footer(docx_section, footer_type, footer)


def _pack_header(docx_section: Any, header_type: str, header: "Header") -> None:
    """Pack a Header into a python-docx section.

    Args:
        docx_section: The python-docx Section.
        header_type: The header type ('default', 'first', 'even').
        header: The declarative Header.
    """
    if header_type == "default":
        docx_header = docx_section.header
    elif header_type == "first":
        docx_header = docx_section.first_page_header
    elif header_type == "even":
        docx_header = docx_section.even_page_header
    else:
        return

    if header.children:
        for child in header.children:
            _pack_block_element(docx_header, child)


def _pack_footer(docx_section: Any, footer_type: str, footer: "Footer") -> None:
    """Pack a Footer into a python-docx section.

    Args:
        docx_section: The python-docx Section.
        footer_type: The footer type ('default', 'first', 'even').
        footer: The declarative Footer.
    """
    if footer_type == "default":
        docx_footer = docx_section.footer
    elif footer_type == "first":
        docx_footer = docx_section.first_page_footer
    elif footer_type == "even":
        docx_footer = docx_section.even_page_footer
    else:
        return

    if footer.children:
        for child in footer.children:
            _pack_block_element(docx_footer, child)


def _pack_block_element(container: Any, element: Any) -> None:
    """Pack a block-level element (Paragraph or Table).

    Args:
        container: The container to add to (Document, Header, Footer, or Cell).
        element: The Paragraph or Table to pack.
    """
    from cmi_docx.declarative.paragraph import Paragraph
    from cmi_docx.declarative.table import Table

    if isinstance(element, Paragraph):
        _pack_paragraph(container, element)
    elif isinstance(element, Table):
        _pack_table(container, element)


def _pack_paragraph(container: Any, para: "Paragraph") -> None:
    """Pack a Paragraph into a container.

    Args:
        container: The container to add to.
        para: The declarative Paragraph.
    """
    docx_para = container.add_paragraph()

    if para.style:
        docx_para.style = para.style

    if para.heading:
        heading_style_names = {
            1: "Heading 1",
            2: "Heading 2",
            3: "Heading 3",
            4: "Heading 4",
            5: "Heading 5",
            6: "Heading 6",
            7: "Heading 7",
            8: "Heading 8",
            9: "Heading 9",
        }
        if para.heading in heading_style_names:
            docx_para.style = heading_style_names[para.heading]

    if para.alignment:
        docx_para.alignment = para.alignment

    if para.spacing_before:
        docx_para.paragraph_format.space_before = Pt(para.spacing_before)
    if para.spacing_after:
        docx_para.paragraph_format.space_after = Pt(para.spacing_after)
    if para.line_spacing:
        docx_para.paragraph_format.line_spacing = para.line_spacing

    if para.left_indent:
        docx_para.paragraph_format.left_indent = Pt(para.left_indent)
    if para.right_indent:
        docx_para.paragraph_format.right_indent = Pt(para.right_indent)
    if para.first_line_indent:
        docx_para.paragraph_format.first_line_indent = Pt(para.first_line_indent)

    if para.keep_together is not None:
        docx_para.paragraph_format.keep_together = para.keep_together
    if para.keep_with_next is not None:
        docx_para.paragraph_format.keep_with_next = para.keep_with_next
    if para.page_break_before is not None:
        docx_para.paragraph_format.page_break_before = para.page_break_before
    if para.widow_control is not None:
        docx_para.paragraph_format.widow_control = para.widow_control

    if para.text:
        docx_para.add_run(para.text)
    elif para.children:
        for child in para.children:
            _pack_inline_element(docx_para, child)


def _pack_inline_element(docx_para: Any, element: Any) -> None:
    """Pack an inline element (TextRun, ImageRun, Tab, Break).

    Args:
        docx_para: The python-docx Paragraph.
        element: The inline element to pack.
    """
    from cmi_docx.declarative.image import ImageRun
    from cmi_docx.declarative.paragraph import Break, Tab, TextRun

    if isinstance(element, TextRun):
        _pack_text_run(docx_para, element)
    elif isinstance(element, ImageRun):
        _pack_image_run(docx_para, element)
    elif isinstance(element, Tab):
        docx_para.add_run().add_tab()
    elif isinstance(element, Break):
        if element.type == "page":
            docx_para.add_run().add_break(docx_text.WD_BREAK.PAGE)
        elif element.type == "column":
            docx_para.add_run().add_break(docx_text.WD_BREAK.COLUMN)
        else:
            docx_para.add_run().add_break()


def _pack_text_run(docx_para: Any, run: "TextRun") -> None:
    """Pack a TextRun into a paragraph.

    Args:
        docx_para: The python-docx Paragraph.
        run: The declarative TextRun.
    """
    docx_run = docx_para.add_run(run.text)

    if run.bold is not None:
        docx_run.bold = run.bold
    if run.italic is not None:
        docx_run.italic = run.italic
    if run.underline is not None:
        docx_run.underline = run.underline

    if run.font:
        docx_run.font.name = run.font
    if run.size:
        docx_run.font.size = Pt(run.size)
    if run.color:
        docx_run.font.color.rgb = RGBColor(*run.color)

    if run.superscript:
        docx_run.font.superscript = True
    if run.subscript:
        docx_run.font.subscript = True
    if run.strike:
        docx_run.font.strike = True
    if run.all_caps:
        docx_run.font.all_caps = True
    if run.small_caps:
        docx_run.font.small_caps = True
    if run.highlight:
        docx_run.font.highlight_color = run.highlight


def _pack_image_run(docx_para: Any, image: "ImageRun") -> None:
    """Pack an ImageRun into a paragraph.

    Args:
        docx_para: The python-docx Paragraph.
        image: The declarative ImageRun.
    """
    docx_run = docx_para.add_run()

    width = None
    height = None
    if image.transformation:
        if "width" in image.transformation:
            width = Pt(image.transformation["width"])
        if "height" in image.transformation:
            height = Pt(image.transformation["height"])

    if isinstance(image.data, bytes):
        import io

        docx_run.add_picture(io.BytesIO(image.data), width=width, height=height)
    elif isinstance(image.data, (str, pathlib.Path)):
        docx_run.add_picture(str(image.data), width=width, height=height)


def _pack_table(container: Any, table: "Table") -> None:
    """Pack a Table into a container.

    Args:
        container: The container to add to.
        table: The declarative Table.
    """
    num_rows = len(table.rows)
    num_cols = max(len(row.children) for row in table.rows) if table.rows else 0

    docx_table = container.add_table(rows=num_rows, cols=num_cols)

    if table.style:
        docx_table.style = table.style

    for row_idx, row in enumerate(table.rows):
        _pack_table_row(docx_table.rows[row_idx], row)


def _pack_table_row(docx_row: Any, row: "TableRow") -> None:
    """Pack a TableRow.

    Args:
        docx_row: The python-docx table row.
        row: The declarative TableRow.
    """
    if row.height and "value" in row.height:
        docx_row.height = row.height["value"]

    if row.cant_split is not None:
        docx_row.cant_split = row.cant_split

    for cell_idx, cell in enumerate(row.children):
        _pack_table_cell(docx_row.cells[cell_idx], cell)


def _pack_table_cell(docx_cell: Any, cell: "TableCell") -> None:
    """Pack a TableCell.

    Args:
        docx_cell: The python-docx table cell.
        cell: The declarative TableCell.
    """
    if cell.width and "size" in cell.width:
        docx_cell.width = cell.width["size"]

    if cell.children:
        docx_cell.text = ""
        for child in cell.children:
            _pack_block_element(docx_cell, child)
