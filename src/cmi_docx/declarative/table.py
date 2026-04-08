"""Table components for declarative documents."""

from __future__ import annotations

import dataclasses
from typing import TYPE_CHECKING

from cmi_docx.declarative import base, paragraph

if TYPE_CHECKING:
    from collections.abc import Callable, Coroutine, Sequence


@dataclasses.dataclass
class TableCell(base.Component):
    """A table cell containing paragraphs or nested tables.

    Attributes:
        children: List of Paragraph or Table components, or coroutines that
            resolve to these types. May be a zero-argument callable for lazy
            evaluation (useful with ``condition``).
        borders: Cell border configuration.
    """

    children: (
        Sequence[
            paragraph.Paragraph
            | Table
            | Coroutine[None, None, paragraph.Paragraph | Table]
        ]
        | Callable[
            [],
            Sequence[
                paragraph.Paragraph
                | Table
                | Coroutine[None, None, paragraph.Paragraph | Table]
            ],
        ]
        | None
    ) = None
    borders: dict[str, dict[str, str | int]] | None = None


@dataclasses.dataclass
class TableRow(base.Component):
    """A table row containing cells.

    Attributes:
        children: List of TableCell components or coroutines that resolve to
            cells. May be a zero-argument callable for lazy evaluation (useful
            with ``condition``).
    """

    children: (
        Sequence[TableCell | Coroutine[None, None, TableCell]]
        | Callable[[], Sequence[TableCell | Coroutine[None, None, TableCell]]]
    )


@dataclasses.dataclass
class Table(base.Component):
    """A table with rows and cells.

    Attributes:
        rows: List of TableRow components or coroutines that resolve to rows.
            May be a zero-argument callable for lazy evaluation (useful with
            ``condition``).
        column_widths: List of column widths in DXA units.
        width: Table width configuration (dict with 'size' and 'type' keys).
        borders: Table border configuration.
        alignment: Table alignment ('left', 'center', 'right').
        indent: Table indent from left margin.
        layout: Table layout type ('autofit', 'fixed').
        style: Table style name.
    """

    rows: (
        Sequence[TableRow | Coroutine[None, None, TableRow]]
        | Callable[[], Sequence[TableRow | Coroutine[None, None, TableRow]]]
    )
    column_widths: list[int] | None = None
    width: dict[str, int | str] | None = None
    borders: dict[str, dict[str, str | int]] | None = None
    alignment: str | None = None
    indent: int | None = None
    layout: str | None = None
    style: str | None = None
