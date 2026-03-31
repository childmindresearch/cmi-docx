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


@pytest.mark.asyncio
async def test_declarative_template_paragraph_index_insert_at_beginning() -> None:
    """Test inserting sections at the beginning of a template document."""
    with tempfile.TemporaryDirectory() as tmpdir:
        template_path = pathlib.Path(tmpdir) / "template.docx"

        template_doc = declarative.Document(
            sections=[
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Template First"),
                        declarative.Paragraph(text="Template Second"),
                    ],
                ),
            ],
        )
        template_docx = await template_doc.to_docx()
        template_docx.save(str(template_path))

        doc = declarative.Document(
            sections=[
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Inserted A"),
                        declarative.Paragraph(text="Inserted B"),
                    ],
                ),
            ],
        )
        template = declarative.DocumentTemplate(
            path=template_path,
            paragraph_index=0,
        )

        result = await doc.to_docx(template=template)

        assert result.paragraphs[0].text == "Inserted A"
        assert result.paragraphs[1].text == "Inserted B"
        assert result.paragraphs[2].text == "Template First"
        assert result.paragraphs[3].text == "Template Second"


@pytest.mark.asyncio
async def test_declarative_template_paragraph_index_insert_in_middle() -> None:
    """Test inserting sections in the middle of a template document."""
    with tempfile.TemporaryDirectory() as tmpdir:
        template_path = pathlib.Path(tmpdir) / "template.docx"

        template_doc = declarative.Document(
            sections=[
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Template First"),
                        declarative.Paragraph(text="Template Second"),
                        declarative.Paragraph(text="Template Third"),
                    ],
                ),
            ],
        )
        template_docx = await template_doc.to_docx()
        template_docx.save(str(template_path))

        doc = declarative.Document(
            sections=[
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Inserted A"),
                        declarative.Paragraph(text="Inserted B"),
                    ],
                ),
            ],
        )
        template = declarative.DocumentTemplate(
            path=template_path,
            paragraph_index=1,
        )

        result = await doc.to_docx(template=template)

        assert result.paragraphs[0].text == "Template First"
        assert result.paragraphs[1].text == "Inserted A"
        assert result.paragraphs[2].text == "Inserted B"
        assert result.paragraphs[3].text == "Template Second"
        assert result.paragraphs[4].text == "Template Third"


@pytest.mark.asyncio
async def test_declarative_template_paragraph_index_at_end() -> None:
    """Test inserting at an index past the last paragraph appends to the end."""
    with tempfile.TemporaryDirectory() as tmpdir:
        template_path = pathlib.Path(tmpdir) / "template.docx"

        template_doc = declarative.Document(
            sections=[
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Template First"),
                    ],
                ),
            ],
        )
        template_docx = await template_doc.to_docx()
        template_docx.save(str(template_path))

        doc = declarative.Document(
            sections=[
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Inserted A"),
                    ],
                ),
            ],
        )
        template = declarative.DocumentTemplate(
            path=template_path,
            paragraph_index=999,
        )

        result = await doc.to_docx(template=template)

        paragraph_texts = [para.text for para in result.paragraphs if para.text]
        assert paragraph_texts[0] == "Template First"
        assert paragraph_texts[1] == "Inserted A"


@pytest.mark.asyncio
async def test_declarative_template_paragraph_index_with_table() -> None:
    """Test inserting a table at a paragraph index in a template."""
    with tempfile.TemporaryDirectory() as tmpdir:
        template_path = pathlib.Path(tmpdir) / "template.docx"

        template_doc = declarative.Document(
            sections=[
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Template First"),
                        declarative.Paragraph(text="Template Second"),
                        declarative.Paragraph(text="Template Third"),
                    ],
                ),
            ],
        )
        template_docx = await template_doc.to_docx()
        template_docx.save(str(template_path))

        doc = declarative.Document(
            sections=[
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Before Table"),
                        declarative.Table(
                            rows=[
                                declarative.TableRow(
                                    children=[
                                        declarative.TableCell(
                                            children=[
                                                declarative.Paragraph(text="Cell 1"),
                                            ],
                                        ),
                                        declarative.TableCell(
                                            children=[
                                                declarative.Paragraph(text="Cell 2"),
                                            ],
                                        ),
                                    ],
                                ),
                            ],
                        ),
                        declarative.Paragraph(text="After Table"),
                    ],
                ),
            ],
        )
        template = declarative.DocumentTemplate(
            path=template_path,
            paragraph_index=1,
        )

        result = await doc.to_docx(template=template)

        assert result.paragraphs[0].text == "Template First"
        assert result.paragraphs[1].text == "Before Table"
        assert result.paragraphs[2].text == "After Table"
        assert result.paragraphs[3].text == "Template Second"
        assert result.paragraphs[4].text == "Template Third"

        assert len(result.tables) == 1
        assert result.tables[0].rows[0].cells[0].text == "Cell 1"
        assert result.tables[0].rows[0].cells[1].text == "Cell 2"


@pytest.mark.asyncio
async def test_declarative_template_paragraph_index_with_replacements() -> None:
    """Test paragraph_index works together with replacements."""
    with tempfile.TemporaryDirectory() as tmpdir:
        template_path = pathlib.Path(tmpdir) / "template.docx"

        template_doc = declarative.Document(
            sections=[
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Hello {{NAME}}"),
                        declarative.Paragraph(text="Template End"),
                    ],
                ),
            ],
        )
        template_docx = await template_doc.to_docx()
        template_docx.save(str(template_path))

        doc = declarative.Document(
            sections=[
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Inserted Content"),
                    ],
                ),
            ],
        )
        template = declarative.DocumentTemplate(
            path=template_path,
            replacements={"{{NAME}}": "Alice"},
            paragraph_index=1,
        )

        result = await doc.to_docx(template=template)

        assert result.paragraphs[0].text == "Hello Alice"
        assert result.paragraphs[1].text == "Inserted Content"
        assert result.paragraphs[2].text == "Template End"


@pytest.mark.asyncio
async def test_declarative_template_paragraph_index_multiple_sections() -> None:
    """Test paragraph_index with multiple declarative sections."""
    with tempfile.TemporaryDirectory() as tmpdir:
        template_path = pathlib.Path(tmpdir) / "template.docx"

        template_doc = declarative.Document(
            sections=[
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Template First"),
                        declarative.Paragraph(text="Template Second"),
                        declarative.Paragraph(text="Template Third"),
                    ],
                ),
            ],
        )
        template_docx = await template_doc.to_docx()
        template_docx.save(str(template_path))

        doc = declarative.Document(
            sections=[
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Section1 A"),
                    ],
                ),
                declarative.Section(
                    children=[
                        declarative.Paragraph(text="Section2 A"),
                    ],
                ),
            ],
        )
        template = declarative.DocumentTemplate(
            path=template_path,
            paragraph_index=1,
        )

        result = await doc.to_docx(template=template)

        assert result.paragraphs[0].text == "Template First"
        assert result.paragraphs[1].text == "Section1 A"
        assert result.paragraphs[2].text == "Section2 A"
        assert result.paragraphs[3].text == "Template Second"
        assert result.paragraphs[4].text == "Template Third"
