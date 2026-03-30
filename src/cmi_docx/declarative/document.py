"""Top-level Document class for declarative API."""

import dataclasses
import io
import pathlib

import docx
from docx import document as docx_document
from docx import section as docx_section
from docx import shared
from docx import table as docx_table
from docx.enum import section as docx_enum_section
from docx.enum import text as docx_text
from docx.text import paragraph as docx_paragraph

from cmi_docx import document as imperative_document
from cmi_docx.declarative import base, image, paragraph, section, table


@dataclasses.dataclass
class DocumentTemplate:
    """Defines a template to use as a base for a document.

    Attributes:
        path: Path to the document file.
        replacements: Dictionary of needle/replacements to make in the template.
    """

    path: pathlib.Path | str
    replacements: dict[str, str] | None = None


@dataclasses.dataclass
class Document(base.Component):
    """A Word document with sections.

    This is the top-level container for a declarative document. All operations
    are async - use await to resolve all async children concurrently.

    Attributes:
        sections: List of Section components or coroutines that resolve to sections.
        creator: Document creator metadata.
        title: Document title metadata.
        subject: Document subject metadata.
        description: Document description metadata.
        keywords: Document keywords metadata.
        category: Document category metadata.
        comments: Document comments metadata.
        styles: Document-level style definitions.
        numbering: Document-level numbering definitions.

    Example:
        >>> async def create_doc():
        ...     doc = Document(sections=[
        ...         Section(children=[
        ...             Paragraph(text="Hello World"),
        ...             fetch_paragraph(),  # async function
        ...         ]),
        ...     ])
        ...     await doc.save("output.docx")
    """

    sections: list[section.Section]
    creator: str | None = None
    title: str | None = None
    subject: str | None = None
    description: str | None = None
    keywords: str | None = None
    category: str | None = None
    comments: str | None = None
    styles: dict[str, str | int | bool] | None = None
    numbering: dict[str, str | int | list[dict[str, str | int]]] | None = None

    async def to_docx(  # noqa: C901
        self, template: DocumentTemplate | None = None
    ) -> docx_document.Document:
        """Convert to a python-docx Document.

        Automatically resolves all async children before converting.

        Returns:
            A python-docx Document object.
        """
        await self.resolve()
        if template is not None:
            docx_doc = docx.Document()
            if template.replacements is not None:
                extended_doc = imperative_document.ExtendDocument(docx_doc)
                for needle, replacement in template.replacements.items():
                    extended_doc.replace(needle, replacement)
        else:
            docx_doc = (
                docx.Document()
                if template is None
                else docx.Document(str(template.path))
            )

        if self.creator:
            docx_doc.core_properties.author = self.creator
        if self.title:
            docx_doc.core_properties.title = self.title
        if self.subject:
            docx_doc.core_properties.subject = self.subject
        if self.description:
            docx_doc.core_properties.comments = self.description
        if self.keywords:
            docx_doc.core_properties.keywords = self.keywords
        if self.category:
            docx_doc.core_properties.category = self.category

        for sec in self.sections:
            _pack_section(docx_doc, sec)

        return docx_doc


def _pack_section(docx_doc: docx_document.Document, sec: section.Section) -> None:  # noqa: C901, PLR0912
    """Pack a Section into a python-docx document.

    Args:
        docx_doc: The python-docx Document.
        sec: The declarative Section.
    """
    if sec.children:
        for child in sec.children:
            _pack_block_element(docx_doc, child)  # ty:ignore[invalid-argument-type] already awaited.

    docx_section = docx_doc.add_section()

    if sec.properties:
        props = sec.properties
        if props.page_size:
            if "width" in props.page_size:
                docx_section.page_width = props.page_size["width"]
            if "height" in props.page_size:
                docx_section.page_height = props.page_size["height"]

        if props.page_orientation:
            if props.page_orientation.lower() == "landscape":
                docx_section.orientation = docx_enum_section.WD_ORIENTATION.LANDSCAPE
            elif props.page_orientation.lower() == "portrait":
                docx_section.orientation = docx_enum_section.WD_ORIENTATION.PORTRAIT

    if sec.headers:
        for header_type, header in sec.headers.items():
            _pack_header(docx_section, header_type, header)

    if sec.footers:
        for footer_type, footer in sec.footers.items():
            _pack_footer(docx_section, footer_type, footer)


def _get_header_or_footer(
    section: docx_section.Section, hf_type: str, *, is_header: bool
) -> docx_section.Section | None:
    """Get the appropriate header or footer from a section.

    Args:
        section: The python-docx Section.
        hf_type: The type ('default', 'first', 'even').
        is_header: True for header, False for footer.

    Returns:
        The header or footer object, or None if type is invalid.
    """
    type_mapping = {
        "default": "header" if is_header else "footer",
        "first": "first_page_header" if is_header else "first_page_footer",
        "even": "even_page_header" if is_header else "even_page_footer",
    }

    attr_name = type_mapping.get(hf_type)
    return getattr(section, attr_name, None) if attr_name else None


def _pack_header(
    section: docx_section.Section, header_type: str, header: section.Header
) -> None:
    """Pack a Header into a python-docx section.

    Args:
        section: The python-docx Section.
        header_type: The header type ('default', 'first', 'even').
        header: The declarative Header.
    """
    docx_header = _get_header_or_footer(section, header_type, is_header=True)
    if docx_header and header.children:
        for child in header.children:
            _pack_block_element(docx_header, child)  # ty:ignore[invalid-argument-type] already awaited.


def _pack_footer(
    section: docx_section.Section, footer_type: str, footer: section.Footer
) -> None:
    """Pack a Footer into a python-docx section.

    Args:
        section: The python-docx Section.
        footer_type: The footer type ('default', 'first', 'even').
        footer: The declarative Footer.
    """
    docx_footer = _get_header_or_footer(section, footer_type, is_header=False)
    if docx_footer and footer.children:
        for child in footer.children:
            _pack_block_element(docx_footer, child)  # ty:ignore[invalid-argument-type] already awaited.


def _pack_block_element(
    container: docx_document.Document,
    element: paragraph.Paragraph | table.Table,
) -> None:
    """Pack a block-level element (Paragraph or Table).

    Args:
        container: The container to add to.
        element: The Paragraph or Table to pack.
    """
    if isinstance(element, paragraph.Paragraph):
        return _pack_paragraph(container, element)
    return _pack_table(container, element)


def _pack_paragraph(
    container: docx_document.Document, para: paragraph.Paragraph
) -> None:
    """Pack a Paragraph into a container.

    Args:
        container: The container to add to.
        para: The declarative Paragraph.
    """
    docx_para = container.add_paragraph()
    _pack_paragraph_into_existing(docx_para, para)


def _pack_paragraph_into_existing(  # noqa: C901, PLR0912
    docx_para: docx_paragraph.Paragraph, para: paragraph.Paragraph
) -> None:
    """Pack a Paragraph into an existing python-docx paragraph.

    Args:
        docx_para: The python-docx Paragraph to populate.
        para: The declarative Paragraph.
    """
    if para.style:
        docx_para.style = para.style

    if para.heading:
        docx_para.style = f"Heading {para.heading}"

    if para.alignment:
        docx_para.alignment = para.alignment

    fmt = docx_para.paragraph_format
    if para.spacing_before:
        fmt.space_before = shared.Pt(para.spacing_before)
    if para.spacing_after:
        fmt.space_after = shared.Pt(para.spacing_after)
    if para.line_spacing:
        fmt.line_spacing = para.line_spacing
    if para.left_indent:
        fmt.left_indent = shared.Pt(para.left_indent)
    if para.right_indent:
        fmt.right_indent = shared.Pt(para.right_indent)
    if para.first_line_indent:
        fmt.first_line_indent = shared.Pt(para.first_line_indent)
    if para.keep_together is not None:
        fmt.keep_together = para.keep_together
    if para.keep_with_next is not None:
        fmt.keep_with_next = para.keep_with_next
    if para.page_break_before is not None:
        fmt.page_break_before = para.page_break_before
    if para.widow_control is not None:
        fmt.widow_control = para.widow_control

    if para.text:
        docx_para.add_run(para.text)  # ty:ignore[invalid-argument-type] Text is already awaited.
    elif para.children:
        for child in para.children:
            _pack_inline_element(docx_para, child)  # ty:ignore[invalid-argument-type] already awaited.


def _pack_inline_element(
    para: docx_paragraph.Paragraph,
    element: paragraph.TextRun | image.ImageRun | paragraph.Tab | paragraph.Break,
) -> None:
    """Pack an inline element (TextRun, ImageRun, Tab, Break).

    Args:
        para: The python-docx Paragraph.
        element: The inline element to pack.
    """
    if isinstance(element, paragraph.TextRun):
        _pack_text_run(para, element)
    elif isinstance(element, image.ImageRun):
        _pack_image_run(para, element)
    elif isinstance(element, paragraph.Tab):
        para.add_run().add_tab()
    elif isinstance(element, paragraph.Break):
        break_type_mapping = {
            "page": docx_text.WD_BREAK.PAGE,
            "column": docx_text.WD_BREAK.COLUMN,
        }
        break_type = break_type_mapping.get(element.type)
        if break_type:
            para.add_run().add_break(break_type)
        else:
            para.add_run().add_break()


def _pack_text_run(para: docx_paragraph.Paragraph, run: paragraph.TextRun) -> None:  # noqa: C901
    """Pack a TextRun into a paragraph.

    Args:
        para: The python-docx Paragraph.
        run: The declarative TextRun.
    """
    docx_run = para.add_run(run.text)  # ty:ignore[invalid-argument-type] Already awaited.

    if run.bold is not None:
        docx_run.bold = run.bold
    if run.italic is not None:
        docx_run.italic = run.italic
    if run.underline is not None:
        docx_run.underline = run.underline

    font = docx_run.font
    if run.font:
        font.name = run.font
    if run.size:
        font.size = shared.Pt(run.size)
    if run.color:
        font.color.rgb = shared.RGBColor(*run.color)
    if run.superscript:
        font.superscript = True
    if run.subscript:
        font.subscript = True
    if run.strike:
        font.strike = True
    if run.all_caps:
        font.all_caps = True
    if run.small_caps:
        font.small_caps = True


def _pack_image_run(para: docx_paragraph.Paragraph, img: image.ImageRun) -> None:
    """Pack an ImageRun into a paragraph.

    Args:
        para: The python-docx Paragraph.
        img: The declarative ImageRun.
    """
    docx_run = para.add_run()

    width = None
    height = None
    if img.transformation:
        if "width" in img.transformation:
            width = shared.Pt(img.transformation["width"])
        if "height" in img.transformation:
            height = shared.Pt(img.transformation["height"])

    if isinstance(img.data, bytes):
        docx_run.add_picture(
            io.BytesIO(img.data),
            width=width,
            height=height,
        )
    elif isinstance(img.data, (str, pathlib.Path)):
        docx_run.add_picture(
            str(img.data),
            width=width,
            height=height,
        )


def _pack_table(container: docx_document.Document, tbl: table.Table) -> None:
    """Pack a Table into a container.

    Args:
        container: The container to add to.
        tbl: The declarative Table.
    """
    num_rows = len(tbl.rows)
    num_cols = max(len(row.children) for row in tbl.rows) if tbl.rows else 0  # ty:ignore[unresolved-attribute] already awaited.

    docx_table = container.add_table(rows=num_rows, cols=num_cols)

    if tbl.style:
        docx_table.style = tbl.style

    for row_idx, row in enumerate(tbl.rows):
        _pack_table_row(docx_table.rows[row_idx], row)  # ty:ignore[invalid-argument-type] already awaited.


def _pack_table_row(docx_row: docx_table._Row, row: table.TableRow) -> None:
    """Pack a TableRow.

    Args:
        docx_row: The python-docx table row.
        row: The declarative TableRow.
    """
    for cell_idx, cell in enumerate(row.children):
        _pack_table_cell(docx_row.cells[cell_idx], cell)  # ty:ignore[invalid-argument-type] Cell already awaited.


def _pack_table_cell(docx_cell: docx_table._Cell, cell: table.TableCell) -> None:
    """Pack a TableCell.

    Args:
        docx_cell: The python-docx table cell.
        cell: The declarative TableCell.
    """
    if cell.children:
        for idx, child in enumerate(cell.children):
            if idx == 0 and isinstance(child, paragraph.Paragraph):
                # An empty paragraph is created at instantiation of a cell.
                _pack_paragraph_into_existing(docx_cell.paragraphs[0], child)
            else:
                _pack_block_element(docx_cell, child)  # ty:ignore[invalid-argument-type] already awaited.
