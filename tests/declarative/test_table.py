"""Table tests for the declarative API."""

import io

from cmi_docx.declarative import (
    Document,
    Paragraph,
    Section,
    Table,
    TableCell,
    TableRow,
)


def test_simple_table() -> None:
    """Test creating a document with a simple table."""
    doc = Document(
        sections=[
            Section(
                children=[
                    Paragraph(text="Table Example", heading=1),
                    Table(
                        rows=[
                            TableRow(
                                children=[
                                    TableCell(children=[Paragraph(text="Header 1")]),
                                    TableCell(children=[Paragraph(text="Header 2")]),
                                ],
                            ),
                            TableRow(
                                children=[
                                    TableCell(children=[Paragraph(text="Row 1 Col 1")]),
                                    TableCell(children=[Paragraph(text="Row 1 Col 2")]),
                                ],
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


def test_table_with_style() -> None:
    """Test creating a table with styling."""
    doc = Document(
        sections=[
            Section(
                children=[
                    Table(
                        rows=[
                            TableRow(
                                children=[
                                    TableCell(children=[Paragraph(text="Data")]),
                                ],
                            ),
                        ],
                        style="Light Grid",
                    ),
                ],
            ),
        ],
    )

    output = io.BytesIO()
    doc.save(output)
    assert output.tell() > 0
