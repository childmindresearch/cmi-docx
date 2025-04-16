"""Style interfaces for document properties."""

import dataclasses
from typing import Literal

from docx.enum import text


@dataclasses.dataclass
class RunStyle:
    """Dataclass for run style arguments."""

    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    superscript: bool | None = None
    subscript: bool | None = None
    font_size: int | None = None
    font_rgb: tuple[int, int, int] | None = None


@dataclasses.dataclass
class ParagraphStyle:
    """Dataclass for paragraph style arguments."""

    bold: bool | None = None
    italic: bool | None = None
    font_size: int | None = None
    font_rgb: tuple[int, int, int] | None = None
    line_spacing: float | None = None
    space_before: int | None = None
    space_after: int | None = None
    alignment: text.WD_PARAGRAPH_ALIGNMENT | None = None


@dataclasses.dataclass
class CellBorder:
    """Dataclass for cell border style arguments.

    Attributes:
        sides: Tuple of sides for this to apply to, must be "top", "bottom", "insideH",
            "insideV", "start", or "end".
        sz: Size of the border, 1pt = 8/
        val: Whether to use a single or dashed line. None makes the line invisible.
        color: The color of the border as a hex code.
    """

    sides: tuple[str, ...]
    sz: int | None = None
    val: Literal["single", "dashed"] | None = "single"
    color: str | None = None


@dataclasses.dataclass
class CellStyle:
    """Dataclass for table style arguments."""

    paragraph: ParagraphStyle | None = None
    background_rgb: tuple[int, int, int] | None = None
    borders: list[CellBorder] | None = None


@dataclasses.dataclass
class TableSections:
    """Dataclass for enabling/disabling sections of a Word table."""

    first_column: bool | None = None
    first_row: bool | None = None
    last_column: bool | None = None
    last_row: bool | None = None
    no_h_band: bool | None = None
    no_v_band: bool | None = None

    # Alias snake case Python style to camelCase used in Word format.
    @property
    def firstColumn(self) -> bool | None:
        """Alias for first_column."""
        return self.first_column

    @property
    def firstRow(self) -> bool | None:
        """Alias for first_row."""
        return self.first_row

    @property
    def lastColumn(self) -> bool | None:
        """Alias for last_column."""
        return self.last_column

    @property
    def lastRow(self) -> bool | None:
        """Alias for last_row."""
        return self.last_row

    @property
    def noHBand(self) -> bool | None:
        """Alias for no_h_band."""
        return self.no_h_band

    @property
    def noVBand(self) -> bool | None:
        """Alias for no_v_band."""
        return self.no_v_band


@dataclasses.dataclass
class TableStyle:
    """Dataclass for table style arguments."""

    sections: TableSections | None = None
