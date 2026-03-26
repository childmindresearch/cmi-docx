"""Image components for declarative documents."""

import dataclasses
import pathlib
from collections.abc import Coroutine
from typing import Literal

from cmi_docx.declarative import base


@dataclasses.dataclass
class ImageRun(base.Component):
    """An image embedded in the document.

    Attributes:
        data: Image data as bytes, file path, or a coroutine that resolves to bytes.
        type: Image type (e.g., 'png', 'jpg', 'jpeg', 'bmp', 'gif', 'svg').
        transformation: Dictionary with 'width' and/or 'height' in points,
            optional 'rotation' in degrees, optional 'flip' with 'horizontal'
            and/or 'vertical' boolean keys.
        alt_text: Alternative text for accessibility.
    """

    data: bytes | str | pathlib.Path | Coroutine[None, None, bytes]
    type: str | None = None
    transformation: dict[str, int | float | dict[str, bool]] | None = None
    alt_text: dict[Literal["title", "description", "name"], str] | None = None
