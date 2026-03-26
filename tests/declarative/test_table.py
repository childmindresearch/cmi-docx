"""Table tests for the declarative API."""

import io

import cmi_docx


def test_simple_table() -> None:
    """Test creating a document with a simple table."""
    doc = cmi_docx.declarative.Document(
        sections=[
            cmi_docx.declarative.Section(
                children=[
                    cmi_docx.declarative.Paragraph(text="Table Example", heading=1),
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
                                    cmi_docx.declarative.TableCell(
                                        children=[
                                            cmi_docx.declarative.Paragraph(
                                                text="Header 2"
                                            )
                                        ]
                                    ),
                                ],
                            ),
                            cmi_docx.declarative.TableRow(
                                children=[
                                    cmi_docx.declarative.TableCell(
                                        children=[
                                            cmi_docx.declarative.Paragraph(
                                                text="Row 1 Col 1"
                                            )
                                        ]
                                    ),
                                    cmi_docx.declarative.TableCell(
                                        children=[
                                            cmi_docx.declarative.Paragraph(
                                                text="Row 1 Col 2"
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

    output = io.BytesIO()
    doc.save(output)
    assert output.tell() > 0


def test_table_with_style() -> None:
    """Test creating a table with styling."""
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
                                            cmi_docx.declarative.Paragraph(text="Data")
                                        ]
                                    ),
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
