"""Async tests for the declarative API."""

import asyncio
import io

import pytest

import cmi_docx


async def fetch_paragraph() -> cmi_docx.declarative.Paragraph:
    """Simulate fetching a paragraph asynchronously."""
    await asyncio.sleep(0.01)
    return cmi_docx.declarative.Paragraph(text="Async paragraph")


async def fetch_text_run() -> cmi_docx.declarative.TextRun:
    """Simulate fetching a text run asynchronously."""
    await asyncio.sleep(0.01)
    return cmi_docx.declarative.TextRun(text="async text", bold=True)


@pytest.mark.asyncio
async def test_async_paragraph() -> None:
    """Test creating a document with async paragraphs."""
    doc = await cmi_docx.declarative.Document(
        sections=[
            cmi_docx.declarative.Section(
                children=[
                    cmi_docx.declarative.Paragraph(text="Sync paragraph"),
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
    doc = await cmi_docx.declarative.Document(
        sections=[
            cmi_docx.declarative.Section(
                children=[
                    cmi_docx.declarative.Paragraph(
                        children=[
                            cmi_docx.declarative.TextRun(text="Sync text "),
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

    async def fetch_section() -> cmi_docx.declarative.Section:
        await asyncio.sleep(0.01)
        return cmi_docx.declarative.Section(
            children=[
                cmi_docx.declarative.Paragraph(text="Async section content"),
            ],
        )

    doc = await cmi_docx.declarative.Document(
        sections=[
            cmi_docx.declarative.Section(
                children=[
                    cmi_docx.declarative.Paragraph(text="First section"),
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

    async def fetch_with_timing(text: str) -> cmi_docx.declarative.Paragraph:
        import time

        call_times.append(time.time())
        await asyncio.sleep(0.05)
        return cmi_docx.declarative.Paragraph(text=text)

    doc = await cmi_docx.declarative.Document(
        sections=[
            cmi_docx.declarative.Section(
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
    doc = cmi_docx.declarative.Document(
        sections=[
            cmi_docx.declarative.Section(
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
