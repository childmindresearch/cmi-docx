"""Tests for declarative template replacement functionality."""

import pathlib
import tempfile

import pytest

from cmi_docx import declarative


@pytest.mark.asyncio
async def test_declarative_template_replacement() -> None:
    """Test template replacement in declarative API."""
    with tempfile.TemporaryDirectory() as tmpdir:
        template_path = pathlib.Path(tmpdir) / "template.docx"

        template_doc = declarative.Document(
            sections=[
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Hello {{NAME}}"),
                        declarative.Paragraph(text="Age: {{AGE}}"),
                    ],
                ),
            ],
        )
        template_docx = await template_doc.to_docx()
        template_docx.save(str(template_path))

        doc = declarative.Document(sections=[])
        template = declarative.DocumentTemplate(
            path=template_path,
            replacements={"{{NAME}}": "Alice", "{{AGE}}": "25"},
        )

        result = await doc.to_docx(template=template)

        assert result.paragraphs[0].text == "Hello Alice"
        assert result.paragraphs[1].text == "Age: 25"


@pytest.mark.asyncio
async def test_declarative_template_with_table() -> None:
    """Test template replacement in tables using declarative API."""
    with tempfile.TemporaryDirectory() as tmpdir:
        template_path = pathlib.Path(tmpdir) / "template.docx"

        template_doc = declarative.Document(
            sections=[
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Report for {{COMPANY}}"),
                        declarative.Table(
                            rows=[
                                declarative.TableRow(
                                    children=[
                                        declarative.TableCell(
                                            children=[
                                                declarative.Paragraph(
                                                    text="{{METRIC1}}"
                                                ),
                                            ],
                                        ),
                                        declarative.TableCell(
                                            children=[
                                                declarative.Paragraph(
                                                    text="{{VALUE1}}"
                                                ),
                                            ],
                                        ),
                                    ],
                                ),
                                declarative.TableRow(
                                    children=[
                                        declarative.TableCell(
                                            children=[
                                                declarative.Paragraph(
                                                    text="{{METRIC2}}"
                                                ),
                                            ],
                                        ),
                                        declarative.TableCell(
                                            children=[
                                                declarative.Paragraph(
                                                    text="{{VALUE2}}"
                                                ),
                                            ],
                                        ),
                                    ],
                                ),
                            ],
                        ),
                    ],
                ),
            ],
        )
        template_docx = await template_doc.to_docx()
        template_docx.save(str(template_path))

        doc = declarative.Document(sections=[])
        template = declarative.DocumentTemplate(
            path=template_path,
            replacements={
                "{{COMPANY}}": "Acme Corp",
                "{{METRIC1}}": "Revenue",
                "{{VALUE1}}": "$1M",
                "{{METRIC2}}": "Profit",
                "{{VALUE2}}": "$200K",
            },
        )

        result = await doc.to_docx(template=template)

        assert result.paragraphs[0].text == "Report for Acme Corp"
        assert result.tables[0].rows[0].cells[0].text == "Revenue"
        assert result.tables[0].rows[0].cells[1].text == "$1M"
        assert result.tables[0].rows[1].cells[0].text == "Profit"
        assert result.tables[0].rows[1].cells[1].text == "$200K"


@pytest.mark.asyncio
async def test_declarative_template_no_replacements() -> None:
    """Test template without replacements."""
    with tempfile.TemporaryDirectory() as tmpdir:
        template_path = pathlib.Path(tmpdir) / "template.docx"

        template_doc = declarative.Document(
            sections=[
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Static content"),
                    ],
                ),
            ],
        )
        template_docx = await template_doc.to_docx()
        template_docx.save(str(template_path))

        doc = declarative.Document(sections=[])
        template = declarative.DocumentTemplate(path=template_path)

        result = await doc.to_docx(template=template)

        assert result.paragraphs[0].text == "Static content"
