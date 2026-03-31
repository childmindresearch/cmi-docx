"""Tests for conditional rendering in the declarative API."""

import pytest

from cmi_docx import declarative


@pytest.mark.asyncio
async def test_section_condition_false() -> None:
    """Test that a section with condition=False is not rendered."""
    doc = declarative.Document(
        sections=[
            declarative.Section(
                children=[
                    declarative.Paragraph(text="Visible section"),
                ],
            ),
            declarative.Section(
                children=[
                    declarative.Paragraph(text="Hidden section"),
                ],
                condition=lambda: False,
            ),
        ],
    )

    docx = await doc.to_docx()
    assert len(docx.sections) == 2  # noqa: PLR2004
    content_paragraphs = [para for para in docx.paragraphs if para.text.strip()]
    assert len(content_paragraphs) == 1
    assert content_paragraphs[0].text.startswith("Visible section")


@pytest.mark.asyncio
async def test_paragraph_condition_false() -> None:
    """Test that a paragraph with condition=False is not rendered."""
    doc = declarative.Document(
        sections=[
            declarative.Section(
                children=[
                    declarative.Paragraph(text="First paragraph"),
                    declarative.Paragraph(
                        text="Hidden paragraph", condition=lambda: False
                    ),
                    declarative.Paragraph(text="Third paragraph"),
                ],
            ),
        ],
    )

    docx = await doc.to_docx()
    content_paragraphs = [para for para in docx.paragraphs if para.text.strip()]
    assert len(content_paragraphs) == 2  # noqa: PLR2004
    assert content_paragraphs[0].text.startswith("First paragraph")
    assert content_paragraphs[1].text.startswith("Third paragraph")


@pytest.mark.asyncio
async def test_text_run_condition_false() -> None:
    """Test that a text run with condition=False is not rendered."""
    doc = declarative.Document(
        sections=[
            declarative.Section(
                children=[
                    declarative.Paragraph(
                        children=[
                            declarative.TextRun(text="First "),
                            declarative.TextRun(
                                text="Hidden ", condition=lambda: False
                            ),
                            declarative.TextRun(text="Third"),
                        ],
                    ),
                ],
            ),
        ],
    )

    docx = await doc.to_docx()
    assert len(docx.paragraphs[0].runs) == 2  # noqa: PLR2004
    assert docx.paragraphs[0].text == "First Third"


@pytest.mark.asyncio
async def test_table_row_condition_false() -> None:
    """Test that a table row with condition=False is not rendered."""
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
                                            declarative.Paragraph(text="Row 1"),
                                        ]
                                    ),
                                ],
                            ),
                            declarative.TableRow(
                                children=[
                                    declarative.TableCell(
                                        children=[
                                            declarative.Paragraph(
                                                text="Row 2 - Hidden"
                                            ),
                                        ]
                                    ),
                                ],
                                condition=lambda: False,
                            ),
                            declarative.TableRow(
                                children=[
                                    declarative.TableCell(
                                        children=[
                                            declarative.Paragraph(text="Row 3"),
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
    assert len(docx.tables[0].rows) == 2  # noqa: PLR2004
    assert docx.tables[0].rows[0].cells[0].paragraphs[0].text == "Row 1"
    assert docx.tables[0].rows[1].cells[0].paragraphs[0].text == "Row 3"


@pytest.mark.asyncio
async def test_table_cell_condition_false() -> None:
    """Test that a table cell with condition=False is not rendered."""
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
                                            declarative.Paragraph(text="Cell 1"),
                                        ]
                                    ),
                                    declarative.TableCell(
                                        children=[
                                            declarative.Paragraph(
                                                text="Cell 2 - Hidden"
                                            ),
                                        ],
                                        condition=lambda: False,
                                    ),
                                    declarative.TableCell(
                                        children=[
                                            declarative.Paragraph(text="Cell 3"),
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
    assert len(docx.tables[0].rows[0].cells) == 2  # noqa: PLR2004
    assert docx.tables[0].rows[0].cells[0].paragraphs[0].text == "Cell 1"
    assert docx.tables[0].rows[0].cells[1].paragraphs[0].text == "Cell 3"


@pytest.mark.asyncio
async def test_table_condition_false() -> None:
    """Test that a table with condition=False is not rendered."""
    doc = declarative.Document(
        sections=[
            declarative.Section(
                children=[
                    declarative.Paragraph(text="Before table"),
                    declarative.Table(
                        rows=[
                            declarative.TableRow(
                                children=[
                                    declarative.TableCell(
                                        children=[
                                            declarative.Paragraph(text="Hidden table"),
                                        ]
                                    ),
                                ],
                            ),
                        ],
                        condition=lambda: False,
                    ),
                    declarative.Paragraph(text="After table"),
                ],
            ),
        ],
    )

    docx = await doc.to_docx()
    assert len(docx.tables) == 0
    content_paragraphs = [para for para in docx.paragraphs if para.text.strip()]
    assert len(content_paragraphs) == 2  # noqa: PLR2004
    assert content_paragraphs[0].text.startswith("Before table")
    assert content_paragraphs[1].text.startswith("After table")


@pytest.mark.asyncio
async def test_nested_condition_false() -> None:
    """Test that children of false-conditioned components are not rendered."""
    doc = declarative.Document(
        sections=[
            declarative.Section(
                children=[
                    declarative.Paragraph(text="Visible"),
                    declarative.Paragraph(
                        children=[
                            declarative.TextRun(text="This whole paragraph is hidden"),
                        ],
                        condition=lambda: False,
                    ),
                ],
            ),
        ],
    )

    docx = await doc.to_docx()
    content_paragraphs = [para for para in docx.paragraphs if para.text.strip()]
    assert len(content_paragraphs) == 1
    assert content_paragraphs[0].text.startswith("Visible")
