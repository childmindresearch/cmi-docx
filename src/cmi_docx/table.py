"""Extends a python-docx Table cell with additional functionality."""

from docx import oxml, table
from docx.oxml import ns

from cmi_docx import paragraph, styles


class ExtendTable:
    """Extends a python-docx Table with additional functionality."""

    def __init__(self, tbl: table.Table) -> None:
        """Initialize the ExtendTable object.

        Args:
            tbl: The table to extend.
        """
        self.table = tbl

    def format(self, style: styles.TableStyle) -> None:
        """Formats a table in a Word document.

        Args:
            style: The style to use.

        Raises:
            ValueError: If table look was not found.
        """
        if style.sections:
            section_names = (
                "firstColumn",
                "firstRow",
                "lastColumn",
                "lastRow",
                "noHBand",
                "noVBand",
            )
            for name in section_names:
                value = getattr(style.sections, name)
                if value is None:
                    continue

                tbl_pr = self.table._tblPr  # noqa: SLF001
                tbl_look = tbl_pr.first_child_found_in("w:tblLook")
                if tbl_look is None:
                    msg = "Table look was not found."
                    raise ValueError(msg)
                tbl_look.set(ns.qn(f"w:{name}"), str(int(value)))


class ExtendCell:
    """Extends a python-docx Word cell with additional functionality."""

    def __init__(self, cell: table._Cell) -> None:
        """Initializes an ExtendCell object.

        Args:
            cell: The cell to extend.
        """
        self.cell = cell

    def format(self, style: styles.CellStyle) -> None:
        """Formats a cell in a Word table.

        Args:
            style: The style to apply to the cell.
        """
        if style.paragraph is not None:
            for table_paragraph in self.cell.paragraphs:
                paragraph.ExtendParagraph(table_paragraph).format(style.paragraph)

        if style.background_rgb is not None:
            shading = oxml.parse_xml(
                (
                    r'<w:shd {} w:fill="'
                    f"{rgb_to_hex(*style.background_rgb)}"
                    r'"/>'
                ).format(
                    ns.nsdecls("w"),
                ),
            )
            self.cell._tc.get_or_add_tcPr().append(shading)  # noqa: SLF001

        if style.borders:
            for border in style.borders:
                self._apply_border(border)

    def _apply_border(self, border: styles.CellBorder) -> None:
        """Applies the borders styling to the cell.

        Args:
            border: The style to apply to the cell.
        """
        tc_pr = self.cell._tc.get_or_add_tcPr()  # noqa: SLF001

        tc_borders = tc_pr.first_child_found_in("w:tcBorders")
        if tc_borders is None:
            tc_borders = oxml.OxmlElement("w:tcBorders")
            tc_pr.append(tc_borders)

        for edge in border.sides:
            tag = f"w:{edge}"
            element = tc_borders.find(ns.qn(tag))
            if element is None:
                element = oxml.OxmlElement(tag)
                tc_borders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color"]:
                if value := getattr(border, key):
                    element.set(ns.qn(f"w:{key}"), str(value))


def rgb_to_hex(red: int, green: int, blue: int) -> str:
    """Converts RGB values to a hexadecimal color code.

    Args:
        red: The red component of the RGB color.
        green: The green component of the RGB color.
        blue: The blue component of the RGB color.

    Returns:
        The hexadecimal color code representing the RGB color.
    """
    return f"#{red:02x}{green:02x}{blue:02x}".upper()
