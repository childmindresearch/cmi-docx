"""Extends a python-docx Table cell with additional functionality."""

from docx import oxml, table
from docx.enum import text
from docx.oxml import ns

from cmi_docx import paragraph


class ExtendCell:
    """Extends a python-docx Word cell with additional functionality."""

    def __init__(self, cell: table._Cell) -> None:
        """Initializes an ExtendCell object.

        Args:
            cell: The cell to extend.
        """
        self.cell = cell

    def format(
        self,
        *,
        line_spacing: float | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        bold: bool | None = None,
        italics: bool | None = None,
        font_size: int | None = None,
        font_rgb: tuple[int, int, int] | None = None,
        background_rgb: tuple[int, int, int] | None = None,
        alignment: text.WD_PARAGRAPH_ALIGNMENT | None = None,
    ) -> None:
        """Formats a cell in a Word table.

        Args:
            cell: The cell to format.
            line_spacing: The line spacing of the cell.
            space_before: The spacing before the cell.
            space_after: The spacing after the cell.
            bold: Whether to bold the cell.
            italics: Whether to italicize the cell.
            font_size: The font size of the cell.
            font_rgb: The font color of the cell.
            background_rgb: The background color of the cell.
            alignment: The alignment of the cell.
        """
        for table_paragraph in self.cell.paragraphs:
            paragraph.ExtendParagraph(table_paragraph).format(
                line_spacing=line_spacing,
                bold=bold,
                italics=italics,
                font_size=font_size,
                font_rgb=font_rgb,
                alignment=alignment,
                space_after=space_after,
                space_before=space_before,
            )

        if background_rgb is not None:
            shading = oxml.parse_xml(
                (
                    r'<w:shd {} w:fill="' + f"{rgb_to_hex(*background_rgb)}" + r'"/>'
                ).format(
                    ns.nsdecls("w"),
                ),
            )
            self.cell._tc.get_or_add_tcPr().append(shading)  # noqa: SLF001


def rgb_to_hex(r: int, g: int, b: int) -> str:
    """Converts RGB values to a hexadecimal color code.

    Args:
        r: The red component of the RGB color.
        g: The green component of the RGB color.
        b: The blue component of the RGB color.

    Returns:
        The hexadecimal color code representing the RGB color.
    """
    return f"#{r:02x}{g:02x}{b:02x}".upper()
