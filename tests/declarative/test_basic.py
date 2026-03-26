"""Basic tests for the declarative API."""

import io

import cmi_docx


def test_simple_document() -> None:
    """Test creating a simple document with text."""
    doc = cmi_docx.declarative.Document(
        sections=[
            cmi_docx.declarative.Section(
                children=[
                    cmi_docx.declarative.Paragraph(text="Hello World"),
                ],
            ),
        ],
    )

    output = io.BytesIO()
    doc.save(output)
    assert output.tell() > 0


def test_document_with_text_runs() -> None:
    """Test creating a document with formatted text runs."""
    doc = cmi_docx.declarative.Document(
        sections=[
            cmi_docx.declarative.Section(
                children=[
                    cmi_docx.declarative.Paragraph(
                        children=[
                            cmi_docx.declarative.TextRun(text="Bold text", bold=True),
                            cmi_docx.declarative.TextRun(text=" and "),
                            cmi_docx.declarative.TextRun(
                                text="italic text", italic=True
                            ),
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
    doc = cmi_docx.declarative.Document(
        sections=[
            cmi_docx.declarative.Section(
                children=[
                    cmi_docx.declarative.Paragraph(text="Main Heading", heading=1),
                    cmi_docx.declarative.Paragraph(text="Subheading", heading=2),
                    cmi_docx.declarative.Paragraph(text="Body text"),
                ],
            ),
        ],
    )

    output = io.BytesIO()
    doc.save(output)
    assert output.tell() > 0


def test_document_with_tabs_and_breaks() -> None:
    """Test creating a document with tabs and breaks."""
    doc = cmi_docx.declarative.Document(
        sections=[
            cmi_docx.declarative.Section(
                children=[
                    cmi_docx.declarative.Paragraph(
                        children=[
                            cmi_docx.declarative.TextRun(text="Before tab"),
                            cmi_docx.declarative.Tab(),
                            cmi_docx.declarative.TextRun(text="After tab"),
                        ],
                    ),
                    cmi_docx.declarative.Paragraph(
                        children=[
                            cmi_docx.declarative.TextRun(text="Before break"),
                            cmi_docx.declarative.Break(type="line"),
                            cmi_docx.declarative.TextRun(text="After break"),
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
    doc = cmi_docx.declarative.Document(
        sections=[
            cmi_docx.declarative.Section(
                children=[
                    cmi_docx.declarative.Paragraph(text="Content"),
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
