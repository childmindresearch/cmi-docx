"""Table components for declarative documents."""

from __future__ import annotations

import dataclasses
from typing import TYPE_CHECKING

from cmi_docx.declarative import base, paragraph

if TYPE_CHECKING:
    from collections.abc import Callable, Coroutine, MutableSequence, Sequence
    from typing import Literal


@dataclasses.dataclass
class TableCell(base.Component):
    """A table cell containing paragraphs or nested tables.

    Attributes:
        children: List of Paragraph or Table components, or coroutines that
            resolve to these types. May be a zero-argument callable for lazy
            evaluation (useful with ``condition``).
        borders: Cell border configuration.
        grid_span: Number of columns this cell spans (horizontal merge).
            Defaults to None (no spanning).
        vmerge: Vertical merge role. ``"restart"`` marks the top cell of a
            vertical merge group; ``"continue"`` marks subsequent cells in the
            group. Defaults to None.
    """

    children: (
        MutableSequence[
            paragraph.Paragraph
            | Table
            | Coroutine[None, None, paragraph.Paragraph | Table]
        ]
        | Callable[
            [],
            MutableSequence[
                paragraph.Paragraph
                | Table
                | Coroutine[None, None, paragraph.Paragraph | Table]
            ],
        ]
        | None
    ) = None
    borders: dict[str, dict[str, str | int]] | None = None
    grid_span: int | None = None
    vmerge: Literal["restart", "continue"] | None = None


@dataclasses.dataclass
class TableRow(base.Component):
    """A table row containing cells.

    Attributes:
        children: List of TableCell components or coroutines that resolve to
            cells. May be a zero-argument callable for lazy evaluation (useful
            with ``condition``).
    """

    children: (
        MutableSequence[TableCell | Coroutine[None, None, TableCell]]
        | Callable[[], MutableSequence[TableCell | Coroutine[None, None, TableCell]]]
    )


@dataclasses.dataclass
class Table(base.Component):
    """A table with rows and cells.

    Attributes:
        rows: List of TableRow components or coroutines that resolve to rows.
            May be a zero-argument callable for lazy evaluation (useful with
            ``condition``).
        column_widths: List of column widths in twips (DXA). 1440 twips equals
            1 inch; approximately 567 twips equals 1cm. Setting this implies
            fixed layout (autofit is disabled automatically).
        width: Table width configuration (dict with 'size' and 'type' keys).
        borders: Table border configuration.
        alignment: Table alignment ('left', 'center', 'right').
        indent: Table indent from left margin.
        layout: Table layout type. ``"fixed"`` sets fixed layout (autofit
            disabled). ``"autofit"`` sets autofit layout and suppresses column
            widths even if ``column_widths`` is also provided.
        style: Table style name.
    """

    rows: (
        MutableSequence[TableRow | Coroutine[None, None, TableRow]]
        | Callable[[], MutableSequence[TableRow | Coroutine[None, None, TableRow]]]
    )
    column_widths: Sequence[int] | None = None
    width: dict[str, int | str] | None = None
    borders: dict[str, dict[str, str | int]] | None = None
    alignment: str | None = None
    indent: int | None = None
    layout: Literal["autofit", "fixed"] | None = None
    style: str | None = None
