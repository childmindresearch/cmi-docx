"""Basic tests for the declarative API."""

import io

from cmi_docx.declarative import (
    Break,
    Document,
    Paragraph,
    Section,
    Tab,
    TextRun,
)


def test_simple_document() -> None:
    """Test creating a simple document with text."""
    doc = Document(
        sections=[
            Section(
                children=[
                    Paragraph(text="Hello World"),
                ],
            ),
        ],
    )

    output = io.BytesIO()
    doc.save(output)
    assert output.tell() > 0


def test_document_with_text_runs() -> None:
    """Test creating a document with formatted text runs."""
    doc = Document(
        sections=[
            Section(
                children=[
                    Paragraph(
                        children=[
                            TextRun(text="Bold text", bold=True),
                            TextRun(text=" and "),
                            TextRun(text="italic text", italic=True),
                        ],
                    ),
                ],
            ),
        ],
    )

    output = io.BytesIO()
    doc.save(output)
    assert output.tell() > 0


def test_document_with_heading() -> None:
    """Test creating a document with headings."""
    doc = Document(
        sections=[
            Section(
                children=[
                    Paragraph(text="Main Heading", heading=1),
                    Paragraph(text="Subheading", heading=2),
                    Paragraph(text="Body text"),
                ],
            ),
        ],
    )

    output = io.BytesIO()
    doc.save(output)
    assert output.tell() > 0


def test_document_with_tabs_and_breaks() -> None:
    """Test creating a document with tabs and breaks."""
    doc = Document(
        sections=[
            Section(
                children=[
                    Paragraph(
                        children=[
                            TextRun(text="Before tab"),
                            Tab(),
                            TextRun(text="After tab"),
                        ],
                    ),
                    Paragraph(
                        children=[
                            TextRun(text="Before break"),
                            Break(type="line"),
                            TextRun(text="After break"),
                        ],
                    ),
                ],
            ),
        ],
    )

    output = io.BytesIO()
    doc.save(output)
    assert output.tell() > 0


def test_document_metadata() -> None:
    """Test creating a document with metadata."""
    doc = Document(
        sections=[
            Section(
                children=[
                    Paragraph(text="Content"),
                ],
            ),
        ],
        title="Test Document",
        creator="Test Author",
        subject="Testing",
    )

    output = io.BytesIO()
    doc.save(output)
    assert output.tell() > 0

    docx_doc = doc.to_docx()
    assert docx_doc.core_properties.title == "Test Document"
    assert docx_doc.core_properties.author == "Test Author"
    assert docx_doc.core_properties.subject == "Testing"
