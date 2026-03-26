"""Image components for declarative documents."""

import dataclasses
import pathlib
from collections.abc import Coroutine
from typing import Any

from cmi_docx.declarative.base import Component


@dataclasses.dataclass
class ImageRun(Component):
    """An image embedded in the document.

    Attributes:
        data: Image data as bytes, file path (str or Path), or a coroutine
            that resolves to bytes.
        type: Image type ('png', 'jpg', 'jpeg', 'bmp', 'gif', 'svg').
        transformation: Dictionary with 'width' and/or 'height' in points,
            optional 'rotation' in degrees, optional 'flip' with 'horizontal'
            and/or 'vertical' boolean keys.
        alt_text: Alternative text for accessibility (dict with 'title',
            'description', 'name' keys).
        floating: Floating image positioning (dict with 'horizontalPosition',
            'verticalPosition', 'zIndex', etc.).
    """

    data: bytes | str | pathlib.Path | Coroutine[Any, Any, bytes]
    type: str | None = None
    transformation: dict[str, Any] | None = None
    alt_text: dict[str, str] | None = None
    floating: dict[str, Any] | None = None
