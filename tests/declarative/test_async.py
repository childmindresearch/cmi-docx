"""Async tests for the declarative API."""

import asyncio

import pytest

from cmi_docx import declarative


async def fetch_paragraph() -> declarative.Paragraph:
    """Simulate fetching a paragraph asynchronously."""
    await asyncio.sleep(0.01)
    return declarative.Paragraph(text="Async paragraph")


async def fetch_text_run() -> declarative.TextRun:
    """Simulate fetching a text run asynchronously."""
    await asyncio.sleep(0.01)
    return declarative.TextRun(text="async text", bold=True)


@pytest.mark.asyncio
async def test_async_paragraph() -> None:
    """Test creating a document with async paragraphs."""
    doc = declarative.Document(
        sections=[
            declarative.Section(
                children=[
                    declarative.Paragraph(text="Sync paragraph"),
                    fetch_paragraph(),
                ],
            ),
        ],
    )

    docx = await doc.to_docx()
    assert docx.paragraphs[0].text.startswith("Sync")
    assert docx.paragraphs[1].text.startswith("Async")


@pytest.mark.asyncio
async def test_async_text_run() -> None:
    """Test creating a document with async text runs."""
    doc = declarative.Document(
        sections=[
            declarative.Section(
                children=[
                    declarative.Paragraph(
                        children=[
                            declarative.TextRun(text="Sync text"),
                            fetch_text_run(),
                        ],
                    ),
                ],
            ),
        ],
    )

    docx = await doc.to_docx()
    assert docx.paragraphs[0].text.startswith("Sync text")
    assert "async text" in docx.paragraphs[0].text
