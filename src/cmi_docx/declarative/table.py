"""Table components for declarative documents."""

import dataclasses
from typing import Any

from cmi_docx.declarative.base import Component


@dataclasses.dataclass
class TableCell(Component):
    """A table cell containing paragraphs or nested tables.

    Attributes:
        children: List of Paragraph or Table components, or coroutines that
            resolve to these types.
        width: Cell width configuration (dict with 'size' and 'type' keys).
        borders: Cell border configuration.
        shading: Cell shading/background color.
        vertical_align: Vertical alignment ('top', 'center', 'bottom').
        margins: Cell margins (dict with 'top', 'bottom', 'left', 'right').
    """

    children: list[Any] | None = None
    width: dict[str, Any] | None = None
    borders: dict[str, Any] | None = None
    shading: dict[str, Any] | None = None
    vertical_align: str | None = None
    margins: dict[str, int] | None = None


@dataclasses.dataclass
class TableRow(Component):
    """A table row containing cells.

    Attributes:
        children: List of TableCell components or coroutines that resolve to cells.
        height: Row height configuration (dict with 'value' and 'rule' keys).
        cant_split: Prevent row from splitting across pages.
        header: Mark row as header row.
    """

    children: list[Any]
    height: dict[str, Any] | None = None
    cant_split: bool | None = None
    header: bool | None = None


@dataclasses.dataclass
class Table(Component):
    """A table with rows and cells.

    Attributes:
        rows: List of TableRow components or coroutines that resolve to rows.
        column_widths: List of column widths in DXA units.
        width: Table width configuration (dict with 'size' and 'type' keys).
        borders: Table border configuration.
        alignment: Table alignment ('left', 'center', 'right').
        indent: Table indent from left margin.
        layout: Table layout type ('autofit', 'fixed').
        style: Table style name.
    """

    rows: list[Any]
    column_widths: list[int] | None = None
    width: dict[str, Any] | None = None
    borders: dict[str, Any] | None = None
    alignment: str | None = None
    indent: int | None = None
    layout: str | None = None
    style: str | None = None
