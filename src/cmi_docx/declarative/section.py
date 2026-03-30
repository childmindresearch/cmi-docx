"""Section, header, and footer components for declarative documents."""

import dataclasses
from collections.abc import Coroutine
from typing import Literal

from cmi_docx.declarative import base, paragraph, table

type BlockElement = paragraph.Paragraph | table.Table
type HeaderFooterType = Literal["default", "first", "even"]


@dataclasses.dataclass
class SectionProperties:
    """Configuration for a document section.

    Attributes:
        page_size: Page dimensions.
        page_orientation: 'portrait' or 'landscape'.
        page_numbering: Page numbering configuration.
        columns: Column configuration (dict with 'count', 'space', 'separator').
        vertical_align: Vertical alignment ('top', 'center', 'bottom', 'both').
        title_page: Use different header/footer on first page.
        type: Section break type ('nextPage', 'nextColumn', 'continuous',
            'evenPage', 'oddPage').
    """

    page_size: dict[Literal["width", "height"], int] | None = None
    page_orientation: Literal["portrait", "landscape"] | None = None
    page_numbering: dict[Literal["start", "format"], int | str] | None = None
    columns: dict[Literal["count", "space", "separator"], int | bool] | None = None
    vertical_align: Literal["top", "center", "bottom", "both"] | None = None
    title_page: bool | None = None
    type: (
        Literal["nextPage", "nextColumn", "continuous", "evenPage", "oddPage"] | None
    ) = None


@dataclasses.dataclass
class Header(base.Component):
    """A section header.

    Attributes:
        children: List of Paragraph or Table components, or coroutines that
            resolve to these types.
    """

    children: (
        list[
            paragraph.Paragraph
            | table.Table
            | Coroutine[None, None, paragraph.Paragraph | table.Table]
        ]
        | None
    ) = None


@dataclasses.dataclass
class Footer(base.Component):
    """A section footer.

    Attributes:
        children: List of Paragraph or Table components, or coroutines that
            resolve to these types.
    """

    children: (
        list[
            paragraph.Paragraph
            | table.Table
            | Coroutine[None, None, paragraph.Paragraph | table.Table]
        ]
        | None
    ) = None


@dataclasses.dataclass
class Section(base.Component):
    """A document section with optional headers and footers.

    Attributes:
        children: List of Paragraph or Table components, or coroutines that
            resolve to these types.
        properties: Section configuration.
        headers: Dictionary mapping header types ('default', 'first', 'even')
            to Header components.
        footers: Dictionary mapping footer types ('default', 'first', 'even')
            to Footer components.
    """

    children: (
        list[
            paragraph.Paragraph
            | table.Table
            | Coroutine[None, None, paragraph.Paragraph | table.Table]
        ]
        | None
    ) = None
    properties: SectionProperties | None = None
    headers: dict[HeaderFooterType, Header] | None = None
    footers: dict[HeaderFooterType, Footer] | None = None
