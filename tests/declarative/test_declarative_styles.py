"""Tests for declarative style definitions.

Covers ParagraphStyleDefinition, TableStyleDefinition, property application,
and XML inspection of per-section table formatting.
"""

import pytest
from docx import shared
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

from cmi_docx import declarative


def _make_doc(
    styles: list[
        declarative.ParagraphStyleDefinition | declarative.TableStyleDefinition
    ],
    *,
    sections: list[declarative.Section] | None = None,
) -> declarative.Document:
    """Build a minimal Document with the given style definitions.

    Args:
        styles: Style definitions to inject into the document.
        sections: Optional sections; defaults to a single empty section.

    Returns:
        A Document ready for ``await doc.to_docx()``.
    """
    if sections is None:
        sections = [
            declarative.Section(
                children=[declarative.Paragraph(text="placeholder")],
            )
        ]
    return declarative.Document(sections=sections, styles=styles)


@pytest.mark.asyncio
async def test_new_paragraph_style_is_created() -> None:
    """Test that a new ParagraphStyleDefinition creates a named paragraph style.

    After rendering, the style must exist in docx.styles with the correct name
    and type WD_STYLE_TYPE.PARAGRAPH.
    """
    doc = _make_doc([declarative.ParagraphStyleDefinition(name="MyCustomStyle")])

    docx = await doc.to_docx()

    style = docx.styles["MyCustomStyle"]
    assert style is not None
    assert style.name == "MyCustomStyle"
    assert style.type == WD_STYLE_TYPE.PARAGRAPH


@pytest.mark.asyncio
async def test_paragraph_style_font_properties() -> None:
    """Test that font, font_size, bold, and italic are applied to a paragraph style.

    The rendered style must expose the exact values passed in the definition via
    the python-docx ``font`` accessor.
    """
    doc = _make_doc(
        [
            declarative.ParagraphStyleDefinition(
                name="MyFontStyle",
                font="Arial",
                font_size=14,
                bold=True,
                italic=True,
            )
        ]
    )

    docx = await doc.to_docx()

    style = docx.styles["MyFontStyle"]
    assert style.font.name == "Arial"
    assert style.font.size == shared.Pt(14)
    assert style.font.bold is True
    assert style.font.italic is True


@pytest.mark.asyncio
async def test_new_table_style_is_created() -> None:
    """Test that a new TableStyleDefinition creates a named table style.

    After rendering, the style must exist in docx.styles with the correct name
    and type WD_STYLE_TYPE.TABLE.
    """
    doc = _make_doc([declarative.TableStyleDefinition(name="MyTableStyle")])

    docx = await doc.to_docx()

    style = docx.styles["MyTableStyle"]
    assert style is not None
    assert style.name == "MyTableStyle"
    assert style.type == WD_STYLE_TYPE.TABLE


@pytest.mark.asyncio
async def test_table_style_whole_table_font() -> None:
    """Test that whole_table formatting sets font properties on the table style.

    font, font_size, and bold on ``whole_table`` must be accessible via
    the style's ``font`` accessor after rendering.
    """
    doc = _make_doc(
        [
            declarative.TableStyleDefinition(
                name="WTStyle",
                whole_table=declarative.TableSectionFormat(
                    font="Calibri",
                    font_size=10,
                    bold=True,
                ),
            )
        ]
    )

    docx = await doc.to_docx()

    style = docx.styles["WTStyle"]
    assert style.font.name == "Calibri"
    assert style.font.size == shared.Pt(10)
    assert style.font.bold is True


@pytest.mark.asyncio
async def test_table_style_first_row_xml() -> None:
    """Test that first_row formatting produces the correct <w:tblStylePr> XML.

    The element with ``w:type="firstRow"`` must contain:
    - ``<w:rPr><w:b>`` (bold)
    - ``<w:rPr><w:color w:val="FFFFFF">`` (white text)
    - ``<w:tcPr><w:shd w:fill="003366">`` (dark blue background)
    """
    doc = _make_doc(
        [
            declarative.TableStyleDefinition(
                name="HeaderRowStyle",
                first_row=declarative.TableSectionFormat(
                    bold=True,
                    color=(255, 255, 255),
                    background=(0, 51, 102),
                ),
            )
        ]
    )

    docx = await doc.to_docx()

    style = docx.styles["HeaderRowStyle"]
    tbl_style_prs = style.element.findall(qn("w:tblStylePr"))
    first_row_pr = next(
        (el for el in tbl_style_prs if el.get(qn("w:type")) == "firstRow"), None
    )
    assert first_row_pr is not None

    r_pr = first_row_pr.find(qn("w:rPr"))
    assert r_pr is not None
    assert r_pr.find(qn("w:b")) is not None

    color_el = r_pr.find(qn("w:color"))
    assert color_el is not None
    assert color_el.get(qn("w:val")) == "FFFFFF"

    tc_pr = first_row_pr.find(qn("w:tcPr"))
    assert tc_pr is not None
    shd = tc_pr.find(qn("w:shd"))
    assert shd is not None
    assert shd.get(qn("w:fill")) == "003366"


@pytest.mark.asyncio
async def test_heading_style_font_overrides_theme() -> None:
    """Test that setting font on a built-in heading style removes theme font attributes.

    When font="Arial" is applied to the built-in "Heading 1" style via
    ParagraphStyleDefinition, the resolved style must:
    - Report style.font.name == "Arial" (the explicit font wins).
    - Have no w:asciiTheme attribute on the <w:rFonts> element, so the theme
      font cannot silently override the explicit font name at render time.
    """
    doc = _make_doc(
        [declarative.ParagraphStyleDefinition(name="Heading 1", font="Arial")],
        sections=[
            declarative.Section(
                children=[declarative.Paragraph(heading=1, text="Hello")],
            )
        ],
    )

    docx = await doc.to_docx()

    style = docx.styles["Heading 1"]
    assert style.font.name == "Arial"

    r_fonts = style.font.element.rPr.find(qn("w:rFonts"))
    assert r_fonts is not None
    assert r_fonts.get(qn("w:asciiTheme")) is None
