"""Section, header, and footer components for declarative documents."""

import dataclasses
from typing import Any

from cmi_docx.declarative.base import Component


@dataclasses.dataclass
class SectionProperties:
    """Configuration for a document section.

    Attributes:
        page_size: Page dimensions (dict with 'width' and 'height' in twips).
        page_margins: Page margins (dict with 'top', 'bottom', 'left', 'right',
            'header', 'footer', 'gutter' in twips).
        page_orientation: 'portrait' or 'landscape'.
        page_numbering: Page numbering configuration (dict with 'start', 'format').
        columns: Column configuration (dict with 'count', 'space', 'separator').
        vertical_align: Vertical alignment ('top', 'center', 'bottom', 'both').
        title_page: Use different header/footer on first page.
        type: Section break type ('nextPage', 'nextColumn', 'continuous',
            'evenPage', 'oddPage').
    """

    page_size: dict[str, int] | None = None
    page_margins: dict[str, int] | None = None
    page_orientation: str | None = None
    page_numbering: dict[str, Any] | None = None
    columns: dict[str, Any] | None = None
    vertical_align: str | None = None
    title_page: bool | None = None
    type: str | None = None


@dataclasses.dataclass
class Header(Component):
    """A section header.

    Attributes:
        children: List of Paragraph or Table components, or coroutines that
            resolve to these types.
    """

    children: list[Any] | None = None


@dataclasses.dataclass
class Footer(Component):
    """A section footer.

    Attributes:
        children: List of Paragraph or Table components, or coroutines that
            resolve to these types.
    """

    children: list[Any] | None = None


@dataclasses.dataclass
class Section(Component):
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

    children: list[Any] | None = None
    properties: SectionProperties | None = None
    headers: dict[str, Any] | None = None
    footers: dict[str, Any] | None = None
