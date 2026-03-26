"""Top-level Document class for declarative API."""

import dataclasses
import io
import pathlib
from collections.abc import Coroutine

from docx import document

from cmi_docx.declarative import base, section


@dataclasses.dataclass
class Document(base.Component):
    """A Word document with sections.

    This is the top-level container for a declarative document. All operations
    are async - use await to resolve all async children concurrently.

    Attributes:
        sections: List of Section components or coroutines that resolve to sections.
        creator: Document creator metadata.
        title: Document title metadata.
        subject: Document subject metadata.
        description: Document description metadata.
        keywords: Document keywords metadata.
        category: Document category metadata.
        comments: Document comments metadata.
        styles: Document-level style definitions.
        numbering: Document-level numbering definitions.

    Example:
        >>> async def create_doc():
        ...     doc = Document(sections=[
        ...         Section(children=[
        ...             Paragraph(text="Hello World"),
        ...             fetch_paragraph(),  # async function
        ...         ]),
        ...     ])
        ...     await doc.save("output.docx")
    """

    sections: list[section.Section | Coroutine[None, None, section.Section]]
    creator: str | None = None
    title: str | None = None
    subject: str | None = None
    description: str | None = None
    keywords: str | None = None
    category: str | None = None
    comments: str | None = None
    styles: dict[str, str | int | bool] | None = None
    numbering: dict[str, str | int | list[dict[str, str | int]]] | None = None

    async def save(self, path_or_stream: str | pathlib.Path | io.BytesIO) -> None:
        """Save the document to a file or stream.

        Automatically resolves all async children before saving.

        Args:
            path_or_stream: File path (str or Path) or file-like object.
        """
        await self.resolve()

        from cmi_docx.declarative import pack  # noqa: PLC0415 # Circular import

        docx_doc = pack.pack(self)
        docx_doc.save(path_or_stream)

    async def to_docx(self) -> document.Document:
        """Convert to a python-docx Document for interop with Extend* API.

        Automatically resolves all async children before converting.

        Returns:
            A python-docx Document object.
        """
        await self.resolve()

        from cmi_docx.declarative import pack  # noqa: PLC0415 # Circular import

        return pack.pack(self)
