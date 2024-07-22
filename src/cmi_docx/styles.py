"""Style interfaces for document properties."""

import dataclasses

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
    space_before: float | None = None
    space_after: float | None = None
    alignment: text.WD_PARAGRAPH_ALIGNMENT | None = None


@dataclasses.dataclass
class TableStyle:
    """Dataclass for table style arguments."""

    paragraph: ParagraphStyle | None = None
    space_before: float | None = None
    space_after: float | None = None
    background_rgb: tuple[int, int, int] | None = None
