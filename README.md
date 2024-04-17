# CMI-docx

[![Build](https://github.com/childmindresearch/cmi-docx/actions/workflows/test.yaml/badge.svg?branch=main)](https://github.com/childmindresearch/cmi-docx/actions/workflows/test.yaml?query=branch%3Amain)
[![codecov](https://codecov.io/gh/childmindresearch/cmi-docx/branch/main/graph/badge.svg?token=22HWWFWPW5)](https://codecov.io/gh/childmindresearch/cmi-docx)
[![Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)
![stability-stable](https://img.shields.io/badge/stability-stable-green.svg)
[![LGPL--2.1 License](https://img.shields.io/badge/license-LGPL--2.1-blue.svg)](https://github.com/childmindresearch/cmi-docx/blob/main/LICENSE)
[![pages](https://img.shields.io/badge/api-docs-blue)](https://childmindresearch.github.io/cmi-docx)

`cmi-docx` is a Python library for manipulating .docx files built on top of `python-docx`. It extends the functionality of `python-docx` by providing additional features and utilities for working with .docx files.

## Features

- Find and replace.
- Insert paragraphs in the middle of a document.
- Manipulate styles and formatting of paragraphs' substrings.

## Installation

Install this package from pypi using your favorite package manager. For example, using `pip`:

```sh
pip install cmi-docx
```

## Quick start

The following example demonstratesa few features of cmi-docx:

```Python
import docx

from cmi_docx import document

doc = docx.Document()
paragraph = doc.add_paragraph("Hello, world!")
extend_document = document.ExtendDocument(doc)
extend_paragraph = document.ExtendParagraph(paragraph)

# Find and replace text.
extend_document.replace("Hello", "Hi", {"bold": True})

# Insert and image
extend_document.insert_image(index=1, image_path="path/to/image.png")

# Reformat a paragraph
extend_paragraph.format(italics=True)
```
