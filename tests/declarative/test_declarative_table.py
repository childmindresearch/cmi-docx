"""Table tests for the declarative API."""

import pytest

import cmi_docx


@pytest.mark.asyncio
async def test_simple_table() -> None:
    """Test creating a document with a simple table."""
    doc = cmi_docx.declarative.Document(
        sections=[
            cmi_docx.declarative.Section(
                children=[
                    cmi_docx.declarative.Table(
                        rows=[
                            cmi_docx.declarative.TableRow(
                                children=[
                                    cmi_docx.declarative.TableCell(
                                        children=[
                                            cmi_docx.declarative.Paragraph(
                                                text="Header 1"
                                            )
                                        ]
                                    ),
                                ],
                            ),
                        ],
                    ),
                ],
            ),
        ],
    )

    docx = await doc.to_docx()
    assert docx.tables[0].rows[0].cells[0].paragraphs[0].text == "Header 1"
