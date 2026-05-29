"""Table components for declarative documents.

See also https://ooxml.dev/docs/tables/ for details.
"""

from __future__ import annotations

import dataclasses
from typing import TYPE_CHECKING

from cmi_docx.declarative import base, paragraph

if TYPE_CHECKING:
    from collections.abc import Callable, Coroutine, MutableSequence, Sequence
    from typing import Literal


@dataclasses.dataclass
class TableBorder:
    """Defines a table border.

    Attributes:
        side: The side of the cell for the border.
        val: The type of border; only single is currently supported.
        sz: Size of the border.
        color: Color of the border (RGB); 0-255.
    """

    side: Literal["top", "left", "bottom", "right", "insideH", "insideV"]
    sz: int = 1
    color: tuple[int, int, int] = (0, 0, 0)
    val: Literal["single"] = "single"

    @property
    def hex_color(self) -> str:
        """Color as hexadecimal."""
        return f"{self.color[0]:02x}{self.color[1]:02x}{self.color[2]:02x}".upper()


@dataclasses.dataclass
class TableCell(base.Component):
    """A table cell containing paragraphs or nested tables.

    Attributes:
        children: List of Paragraph or Table components, or coroutines that
            resolve to these types. May be a zero-argument callable for lazy
            evaluation (useful with ``condition``).
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
        layout: Table layout type. ``"fixed"`` sets fixed layout (autofit
            disabled). ``"autofit"`` sets autofit layout and suppresses column
            widths even if ``column_widths`` is also provided.
        style: Table style name.
        borders: Cell border configuration.
    """

    rows: (
        MutableSequence[TableRow | Coroutine[None, None, TableRow]]
        | Callable[[], MutableSequence[TableRow | Coroutine[None, None, TableRow]]]
    )
    column_widths: Sequence[int] | None = None
    layout: Literal["autofit", "fixed"] | None = None
    style: str | None = None
    borders: MutableSequence[TableBorder] | None = None
