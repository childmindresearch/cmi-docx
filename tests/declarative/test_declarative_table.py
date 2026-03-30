"""Table tests for the declarative API."""

import pytest

from cmi_docx import declarative


@pytest.mark.asyncio
async def test_simple_table() -> None:
    """Test creating a document with a simple table."""
    doc = declarative.Document(
        sections=[
            declarative.Section(
                children=[
                    declarative.Table(
                        rows=[
                            declarative.TableRow(
                                children=[
                                    declarative.TableCell(
                                        children=[
                                            declarative.Paragraph(text="Header 1"),
                                            declarative.Paragraph(text="Paragraph 2"),
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
    assert docx.tables[0].rows[0].cells[0].paragraphs[1].text == "Paragraph 2"
