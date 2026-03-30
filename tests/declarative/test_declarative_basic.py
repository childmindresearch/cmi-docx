"""Basic tests for the declarative API."""

import pytest

from cmi_docx import declarative


@pytest.mark.asyncio
async def test_simple_document() -> None:
    """Test creating a simple document with text."""
    doc = declarative.Document(
        sections=[
            declarative.Section(
                children=[
                    declarative.Paragraph(text="Hello World"),
                ],
            ),
        ],
    )

    docx = await doc.to_docx()
    assert docx.paragraphs[0].text.startswith("Hello World")


@pytest.mark.asyncio
async def test_document_with_text_runs() -> None:
    """Test creating a document with formatted text runs."""
    doc = declarative.Document(
        sections=[
            declarative.Section(
                children=[
                    declarative.Paragraph(
                        children=[
                            declarative.TextRun(text="Bold text", bold=True),
                            declarative.TextRun(text=" and "),
                            declarative.TextRun(text="italic text", italic=True),
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
    doc = declarative.Document(
        sections=[
            declarative.Section(
                children=[
                    declarative.Paragraph(text="Main Heading", heading=1),
                    declarative.Paragraph(text="Subheading", heading=2),
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
    doc = declarative.Document(
        sections=[
            declarative.Section(
                children=[
                    declarative.Paragraph(text="Content"),
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
