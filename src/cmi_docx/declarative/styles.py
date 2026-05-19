"""Style definition components for declarative documents."""

import dataclasses

from docx.enum import text as docx_text


@dataclasses.dataclass
class TableSectionFormat:
    """Formatting options for a section of a table style.

    Attributes:
        font: Font name.
        font_size: Font size in points.
        bold: Apply bold formatting.
        italic: Apply italic formatting.
        underline: Apply underline formatting.
        color: Font color as (R, G, B) tuple (0-255).
        background: Cell shading fill color as (R, G, B) tuple (0-255).
        alignment: Paragraph alignment.
        spacing_before: Space before paragraph in points.
        spacing_after: Space after paragraph in points.
    """

    font: str | None = None
    font_size: int | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    color: tuple[int, int, int] | None = None
    background: tuple[int, int, int] | None = None
    alignment: docx_text.WD_PARAGRAPH_ALIGNMENT | None = None
    spacing_before: int | None = None
    spacing_after: int | None = None


@dataclasses.dataclass
class TableStyleDefinition:
    """Definition of a named table style with per-section formatting.

    Attributes:
        name: The style name.
        base_style: Name of the style to inherit from.
        whole_table: Formatting applied to all cells.
        first_row: Header row formatting.
        last_row: Footer row formatting.
        first_column: First column formatting.
        last_column: Last column formatting.
        banding_1_row: Alternating row stripe 1 formatting.
        banding_2_row: Alternating row stripe 2 formatting.
        banding_1_column: Alternating column stripe 1 formatting.
        banding_2_column: Alternating column stripe 2 formatting.
        top_left_cell: Top-left corner cell formatting (nwCell).
        top_right_cell: Top-right corner cell formatting (neCell).
        bottom_left_cell: Bottom-left corner cell formatting (swCell).
        bottom_right_cell: Bottom-right corner cell formatting (seCell).

    Note:
        Unlike ``ParagraphStyleDefinition``, this always creates a new style.
        A ``ValueError`` is raised if a style with the given ``name`` already
        exists in the document.
    """

    name: str
    base_style: str | None = None
    whole_table: TableSectionFormat | None = None
    first_row: TableSectionFormat | None = None
    last_row: TableSectionFormat | None = None
    first_column: TableSectionFormat | None = None
    last_column: TableSectionFormat | None = None
    banding_1_row: TableSectionFormat | None = None
    banding_2_row: TableSectionFormat | None = None
    banding_1_column: TableSectionFormat | None = None
    banding_2_column: TableSectionFormat | None = None
    top_left_cell: TableSectionFormat | None = None
    top_right_cell: TableSectionFormat | None = None
    bottom_left_cell: TableSectionFormat | None = None
    bottom_right_cell: TableSectionFormat | None = None


@dataclasses.dataclass
class ParagraphStyleDefinition:
    """Definition of a named paragraph style.

    If a style with ``name`` already exists in the document it will be modified
    in place; otherwise a new style is created.

    Attributes:
        name: The style name. If a style with this name already exists in the
            document it will be modified in place; otherwise a new style is
            created.
        base_style: Name of the style to inherit from. Only applied when
            creating a new style.
        next_paragraph_style: Name of the style to apply to the next paragraph.
        font: Font name.
        font_size: Font size in points.
        bold: Apply bold formatting.
        italic: Apply italic formatting.
        underline: Apply underline formatting.
        color: Font color as (R, G, B) tuple (0-255).
        alignment: Paragraph alignment.
        spacing_before: Space before paragraph in points.
        spacing_after: Space after paragraph in points.
        line_spacing: Line spacing multiplier.
        left_indent: Left indent in inches.
        right_indent: Right indent in inches.
        first_line_indent: First line indent in inches.
        keep_together: Keep all lines of paragraph on same page.
        keep_with_next: Keep paragraph with next paragraph.
        page_break_before: Start paragraph on new page.
        widow_control: Enable widow/orphan control.
    """

    name: str
    base_style: str | None = None
    next_paragraph_style: str | None = None
    font: str | None = None
    font_size: int | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    color: tuple[int, int, int] | None = None
    alignment: docx_text.WD_PARAGRAPH_ALIGNMENT | None = None
    spacing_before: int | None = None
    spacing_after: int | None = None
    line_spacing: float | None = None
    left_indent: float | None = None
    right_indent: float | None = None
    first_line_indent: float | None = None
    keep_together: bool | None = None
    keep_with_next: bool | None = None
    page_break_before: bool | None = None
    widow_control: bool | None = None
