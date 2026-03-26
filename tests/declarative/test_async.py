"""Async tests for the declarative API."""

import asyncio
import io

import pytest

from cmi_docx.declarative import Document, Paragraph, Section, TextRun


async def fetch_paragraph() -> Paragraph:
    """Simulate fetching a paragraph asynchronously."""
    await asyncio.sleep(0.01)
    return Paragraph(text="Async paragraph")


async def fetch_text_run() -> TextRun:
    """Simulate fetching a text run asynchronously."""
    await asyncio.sleep(0.01)
    return TextRun(text="async text", bold=True)


@pytest.mark.asyncio
async def test_async_paragraph() -> None:
    """Test creating a document with async paragraphs."""
    doc = await Document(
        sections=[
            Section(
                children=[
                    Paragraph(text="Sync paragraph"),
                    fetch_paragraph(),
                ],
            ),
        ],
    )

    output = io.BytesIO()
    doc.save(output)
    assert output.tell() > 0


@pytest.mark.asyncio
async def test_async_text_run() -> None:
    """Test creating a document with async text runs."""
    doc = await Document(
        sections=[
            Section(
                children=[
                    Paragraph(
                        children=[
                            TextRun(text="Sync text "),
                            fetch_text_run(),
                        ],
                    ),
                ],
            ),
        ],
    )

    output = io.BytesIO()
    doc.save(output)
    assert output.tell() > 0


@pytest.mark.asyncio
async def test_async_section() -> None:
    """Test creating a document with async sections."""

    async def fetch_section() -> Section:
        await asyncio.sleep(0.01)
        return Section(
            children=[
                Paragraph(text="Async section content"),
            ],
        )

    doc = await Document(
        sections=[
            Section(
                children=[
                    Paragraph(text="First section"),
                ],
            ),
            fetch_section(),
        ],
    )

    output = io.BytesIO()
    doc.save(output)
    assert output.tell() > 0


@pytest.mark.asyncio
async def test_concurrent_resolution() -> None:
    """Test that async children are resolved concurrently."""
    call_times: list[float] = []

    async def fetch_with_timing(text: str) -> Paragraph:
        import time

        call_times.append(time.time())
        await asyncio.sleep(0.05)
        return Paragraph(text=text)

    doc = await Document(
        sections=[
            Section(
                children=[
                    fetch_with_timing("Para 1"),
                    fetch_with_timing("Para 2"),
                    fetch_with_timing("Para 3"),
                ],
            ),
        ],
    )

    output = io.BytesIO()
    doc.save(output)
    assert output.tell() > 0
    assert len(call_times) == 3
    time_diff = max(call_times) - min(call_times)
    assert time_diff < 0.02


def test_unresolved_document_raises_error() -> None:
    """Test that saving unresolved async document raises an error."""
    doc = Document(
        sections=[
            Section(
                children=[
                    fetch_paragraph(),
                ],
            ),
        ],
    )

    with pytest.raises(RuntimeError, match="unresolved async children"):
        output = io.BytesIO()
        doc.save(output)

    with pytest.raises(RuntimeError, match="unresolved async children"):
        doc.to_docx()
