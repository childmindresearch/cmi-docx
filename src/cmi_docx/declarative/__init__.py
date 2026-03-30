"""Declarative API for creating Word documents."""

# sergey: disable-file: IMP001 # Allow barrel exports.

from cmi_docx.declarative.base import Component
from cmi_docx.declarative.document import Document
from cmi_docx.declarative.image import ImageRun
from cmi_docx.declarative.paragraph import Break, Paragraph, Tab, TextRun
from cmi_docx.declarative.section import Footer, Header, Section, SectionProperties
from cmi_docx.declarative.table import Table, TableCell, TableRow

__all__ = [
    "Break",
    "Component",
    "Document",
    "Footer",
    "Header",
    "ImageRun",
    "Paragraph",
    "Section",
    "SectionProperties",
    "Tab",
    "Table",
    "TableCell",
    "TableRow",
    "TextRun",
]
