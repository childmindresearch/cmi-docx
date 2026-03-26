"""Paragraph and text run components for declarative documents."""

from __future__ import annotations

import dataclasses
from typing import TYPE_CHECKING

from cmi_docx.declarative import base

if TYPE_CHECKING:
    from collections.abc import Awaitable, Coroutine

    from docx.enum import text as docx_text

    from cmi_docx.declarative import image

InlineElement = "TextRun | image.ImageRun | Tab | Break"


@dataclasses.dataclass
class TextRun(base.Component):
    """A run of text with formatting.

    Attributes:
        text: The text content.
        bold: Apply bold formatting.
        italic: Apply italic formatting.
        underline: Apply underline formatting.
        font: Font name.
        size: Font size in points.
        color: Text color as (R, G, B) tuple (0-255).
        superscript: Apply superscript formatting.
        subscript: Apply subscript formatting.
        strike: Apply strikethrough formatting.
        all_caps: Apply all caps formatting.
        small_caps: Apply small caps formatting.
    """

    text: Awaitable[str] | str
    comment_text: Awaitable[str] | str | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    font: str | None = None
    size: int | None = None
    color: tuple[int, int, int] | None = None
    superscript: bool | None = None
    subscript: bool | None = None
    strike: bool | None = None
    all_caps: bool | None = None
    small_caps: bool | None = None


@dataclasses.dataclass
class Tab(base.Component):
    """A tab character."""


@dataclasses.dataclass
class Break(base.Component):
    """A line or page break.

    Attributes:
        type: Type of break ('line', 'page', 'column', 'textWrapping').
    """

    type: str = "line"


@dataclasses.dataclass
class Paragraph(base.Component):
    """A paragraph with optional formatting and child runs.

    Attributes:
        text: Shorthand for a paragraph with a single text run.
        children: List of TextRun, ImageRun, Tab, Break, or coroutines
            that resolve to these types.
        heading: Heading level.
        style: Style name to apply.
        alignment: Paragraph alignment.
        spacing_before: Space before paragraph in points.
        spacing_after: Space after paragraph in points.
        line_spacing: Line spacing multiplier.
        left_indent: Left indent in points.
        right_indent: Right indent in points.
        first_line_indent: First line indent in points.
        keep_together: Keep all lines of paragraph on same page.
        keep_with_next: Keep paragraph with next paragraph.
        page_break_before: Start paragraph on new page.
        widow_control: Enable widow/orphan control.
    """

    text: Awaitable[str] | str | None = None
    comment_text: Awaitable[str] | str | None = None
    children: (
        list[
            TextRun
            | image.ImageRun
            | Tab
            | Break
            | Coroutine[None, None, TextRun | image.ImageRun | Tab | Break]
        ]
        | None
    ) = None
    heading: int | None = None
    style: str | None = None
    alignment: docx_text.WD_PARAGRAPH_ALIGNMENT | None = None
    spacing_before: int | None = None
    spacing_after: int | None = None
    line_spacing: float | None = None
    left_indent: int | None = None
    right_indent: int | None = None
    first_line_indent: int | None = None
    keep_together: bool | None = None
    keep_with_next: bool | None = None
    page_break_before: bool | None = None
    widow_control: bool | None = None

    def __post_init__(self) -> None:
        """Validate that either text or children is provided, not both.

        Raises:
            ValueError: If both text and children are provided, or neither is provided.
        """
        if self.text is not None and self.children is not None:
            msg = "Paragraph cannot have both 'text' and 'children'"
            raise ValueError(msg)
        if self.text is None and self.children is None:
            msg = "Paragraph must have either 'text' or 'children'"
            raise ValueError(msg)
