"""Basic tests for the declarative API."""

import pytest

import cmi_docx


@pytest.mark.asyncio
async def test_simple_document() -> None:
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

    docx = await doc.to_docx()
    assert docx.paragraphs[0].text.startswith("Hello World")


@pytest.mark.asyncio
async def test_document_with_text_runs() -> None:
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

    docx = await doc.to_docx()
    assert docx.paragraphs[0].text.startswith("Bold text")
    assert docx.paragraphs[0].runs[0].bold
    assert docx.paragraphs[0].runs[2].italic


@pytest.mark.asyncio
async def test_document_with_heading() -> None:
    """Test creating a document with headings."""
    doc = cmi_docx.declarative.Document(
        sections=[
            cmi_docx.declarative.Section(
                children=[
                    cmi_docx.declarative.Paragraph(text="Main Heading", heading=1),
                    cmi_docx.declarative.Paragraph(text="Subheading", heading=2),
                ],
            ),
        ],
    )

    docx = await doc.to_docx()
    assert docx.paragraphs[0].text.startswith("Main Heading")
    assert docx.paragraphs[0].style is not None
    assert docx.paragraphs[0].style.name == "Heading 1"
    assert docx.paragraphs[1].text.startswith("Subheading")
    assert docx.paragraphs[1].style is not None
    assert docx.paragraphs[1].style.name == "Heading 2"


@pytest.mark.asyncio
async def test_document_metadata() -> None:
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

    docx_doc = await doc.to_docx()
    assert docx_doc.core_properties.title == "Test Document"
    assert docx_doc.core_properties.author == "Test Author"
    assert docx_doc.core_properties.subject == "Testing"
