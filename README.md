# CMI-docx

[![Build](https://github.com/childmindresearch/cmi-docx/actions/workflows/test.yaml/badge.svg?branch=main)](https://github.com/childmindresearch/cmi-docx/actions/workflows/test.yaml?query=branch%3Amain)
[![codecov](https://codecov.io/gh/childmindresearch/cmi-docx/branch/main/graph/badge.svg?token=22HWWFWPW5)](https://codecov.io/gh/childmindresearch/cmi-docx)
[![Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)
![stability-stable](https://img.shields.io/badge/stability-stable-green.svg)
[![LGPL--2.1 License](https://img.shields.io/badge/license-LGPL--2.1-blue.svg)](https://github.com/childmindresearch/cmi-docx/blob/main/LICENSE)
[![pages](https://img.shields.io/badge/api-docs-blue)](https://childmindresearch.github.io/cmi-docx)

`cmi-docx` is a Python library by the [Child Mind Institute](https://childmind.org) that extends [`python-docx`](https://python-docx.readthedocs.io/) with higher-level tooling for `.docx` file manipulation. It provides two complementary APIs:

- **Imperative API** -- wrapper classes around `python-docx` objects for find/replace, formatting, insertion, and comments on existing documents.
- **Declarative API** -- an async-first, component-based system for constructing documents from scratch with conditional rendering, lazy evaluation, and template support.

## Features

- **Find and replace** across an entire document (body, headers, footers, and tables), even when the target text is split across multiple runs by Word's internal formatting.
- **Style-aware replacement** -- apply bold, italic, underline, font size, color, and more to replacement text.
- **Paragraph insertion** -- insert paragraphs by text, by object, or as images at any position in the document body.
- **Run-level formatting** -- read and write formatting on individual runs (bold, italic, underline, superscript, subscript, font size, font color).
- **Paragraph formatting** -- control alignment, line spacing, space before/after, and font properties for entire paragraphs.
- **Table and cell formatting** -- toggle table sections, set cell background colors, and configure cell borders.
- **Word comments** -- programmatically add comments to paragraphs, runs, or ranges of runs, with automatic comment preservation during text edits.
- **Declarative document construction** -- build documents as a tree of `Section`, `Paragraph`, `TextRun`, `Table`, and `ImageRun` components, then render to a `python-docx` `Document` with `await doc.to_docx()`.
- **Async and lazy evaluation** -- declare children as coroutines or callables; they are resolved concurrently via `asyncio.gather`.
- **Conditional rendering** -- attach a `condition` callable to any component to include or exclude it at render time without building its subtree.
- **Template support** -- open an existing `.docx` as a template, apply placeholder replacements, and insert new content at a specific paragraph index.

## Installation

Install from PyPI:

```sh
pip install cmi-docx
```

## Quick start

### Imperative API

The imperative API wraps `python-docx` objects with extension classes that add search, replace, formatting, and insertion capabilities.

#### Find and replace

```python
import docx
from cmi_docx import ExtendDocument, RunStyle

doc = docx.Document()
paragraph = doc.add_paragraph("Hello {{")
paragraph.add_run("FULL_NAME}}")

extend_doc = ExtendDocument(doc)
extend_doc.replace("{{FULL_NAME}}", "Jane Doe", RunStyle(bold=True))

print(doc.paragraphs[0].text)  # "Hello *Jane Doe*"
```

#### Paragraph formatting

```python
import docx
from cmi_docx import ExtendParagraph, ParagraphStyle
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = docx.Document()
paragraph = doc.add_paragraph("Formatted paragraph.")

ExtendParagraph(paragraph).format(
    ParagraphStyle(
        bold=True,
        italic=True,
        font_size=14,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )
)
```

#### Insert a run with formatting

```python
import docx
from cmi_docx import ExtendParagraph, RunStyle

doc = docx.Document()
paragraph = doc.add_paragraph("")
paragraph.add_run("Hello ")
paragraph.add_run("world!")

ExtendParagraph(paragraph).insert_run(1, "beautiful ", RunStyle(bold=True))

print(paragraph.text)  # "Hello beautiful world!"
```

#### Add a comment

```python
import docx
from cmi_docx import add_comment

document = docx.Document()
paragraph = document.add_paragraph("This needs review.")

add_comment(document, paragraph, "Reviewer", "Please check this section.")
```

### Declarative API

The declarative API lets you build documents as a tree of components. All children are resolved concurrently, and components can be conditionally included or lazily constructed.

#### Simple document

```python
import asyncio
from cmi_docx.declarative import Document, Section, Paragraph, TextRun

async def main() -> None:
    doc = Document(
        sections=[
            Section(
                children=[
                    Paragraph(text="Main Heading", heading=1),
                    Paragraph(
                        children=[
                            TextRun(text="Bold text", bold=True),
                            TextRun(text=" and "),
                            TextRun(text="italic text", italic=True),
                        ],
                    ),
                ],
            ),
        ],
        title="My Document",
        creator="Author Name",
    )

    docx_doc = await doc.to_docx()
    docx_doc.save("output.docx")

asyncio.run(main())
```

#### Async children

Components accept coroutines as children, which are resolved concurrently:

```python
import asyncio
from cmi_docx.declarative import Document, Section, Paragraph

async def fetch_paragraph() -> Paragraph:
    await asyncio.sleep(0.1)  # Simulate an API call
    return Paragraph(text="Content fetched asynchronously")

async def main() -> None:
    doc = Document(
        sections=[
            Section(
                children=[
                    Paragraph(text="Static content"),
                    fetch_paragraph(),
                ],
            ),
        ],
    )

    docx_doc = await doc.to_docx()
    docx_doc.save("output.docx")

asyncio.run(main())
```

#### Conditional rendering

Attach a `condition` callable to skip components without building their subtree:

```python
from cmi_docx.declarative import Document, Section, Paragraph

include_details = False

doc = Document(
    sections=[
        Section(
            children=[
                Paragraph(text="Always visible"),
                Paragraph(
                    text="Only shown when details are enabled",
                    condition=lambda: include_details,
                ),
            ],
        ),
    ],
)
```

#### Template-based documents

Open an existing `.docx` as a template, replace placeholders, and insert new content:

```python
import asyncio
from pathlib import Path
from cmi_docx.declarative import Document, DocumentTemplate, Section, Paragraph

async def main() -> None:
    doc = Document(
        sections=[
            Section(
                children=[Paragraph(text="Inserted content")],
            ),
        ],
    )

    template = DocumentTemplate(
        path=Path("template.docx"),
        replacements={"{{NAME}}": "Alice", "{{DATE}}": "2025-01-01"},
        paragraph_index=1,  # Insert after the first template paragraph
    )

    docx_doc = await doc.to_docx(template=template)
    docx_doc.save("output.docx")

asyncio.run(main())
```
