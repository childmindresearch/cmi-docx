"""Table components for declarative documents."""

from __future__ import annotations

import dataclasses
from typing import TYPE_CHECKING

from cmi_docx.declarative import base

if TYPE_CHECKING:
    from collections.abc import Coroutine, Iterable

    import cmi_docx.declarative


@dataclasses.dataclass
class TableCell(base.Component):
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

    children: (
        Iterable[
            cmi_docx.declarative.Paragraph
            | Table
            | Coroutine[None, None, cmi_docx.declarative.Paragraph | Table]
        ]
        | None
    ) = None
    width: dict[str, int | str] | None = None
    borders: dict[str, dict[str, str | int]] | None = None
    shading: dict[str, str] | None = None
    vertical_align: str | None = None
    margins: dict[str, int] | None = None


@dataclasses.dataclass
class TableRow(base.Component):
    """A table row containing cells.

    Attributes:
        children: List of TableCell components or coroutines that resolve to cells.
        height: Row height configuration (dict with 'value' and 'rule' keys).
        cant_split: Prevent row from splitting across pages.
        header: Mark row as header row.
    """

    children: list[TableCell | Coroutine[None, None, TableCell]]
    height: dict[str, int | str] | None = None
    cant_split: bool | None = None
    header: bool | None = None


@dataclasses.dataclass
class Table(base.Component):
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

    rows: list[TableRow | Coroutine[None, None, TableRow]]
    column_widths: list[int] | None = None
    width: dict[str, int | str] | None = None
    borders: dict[str, dict[str, str | int]] | None = None
    alignment: str | None = None
    indent: int | None = None
    layout: str | None = None
    style: str | None = None
