# Agent Instructions

## Project Overview

`cmi-docx` is a Python library (by Child Mind Institute) that extends
`python-docx` with additional tooling for `.docx` file manipulation: find/replace,
paragraph insertion, style formatting, comments, table formatting, and a
declarative async-first API for document construction.

## Repository Layout

```
src/cmi_docx/              # Main package (src layout)
  __init__.py              # Barrel exports for all public symbols
  document.py              # ExtendDocument: document-level find/replace/insert
  paragraph.py             # ExtendParagraph, FindParagraph: paragraph-level ops
  run.py                   # ExtendRun, FindRun: run-level find/replace/format
  table.py                 # ExtendTable, ExtendCell: table/cell formatting
  styles.py                # Style dataclasses (RunStyle, ParagraphStyle, CellStyle, etc.)
  comment.py               # add_comment(), CommentPreserver, CommentRange
  declarative/             # Declarative async component-based API
    base.py                # Component base class (async resolve, lazy fields, conditions)
    document.py            # Document, DocumentTemplate, packing functions
    paragraph.py           # Paragraph, TextRun, Tab, Break
    section.py             # Section, Header, Footer, SectionProperties
    table.py               # Table, TableRow, TableCell
    image.py               # ImageRun

tests/                     # pytest test suite
  test_document.py
  test_paragraph.py
  test_run.py
  test_table.py
  test_comment.py
  declarative/             # Tests for declarative API
    test_declarative_basic.py
    test_async.py
    test_condition.py
    test_declarative_table.py
    test_template_replacement.py

pyproject.toml             # Project metadata, dependencies, and all tool config
justfile                   # Task runner (agentcheck recipe)
.pre-commit-config.yaml    # Pre-commit hooks
docs/pdoc-theme/           # Custom CSS for pdoc-generated API docs
local/                     # Gitignored scratch space for local experiments
```

## Running Commands

**Always use `uv run` to execute any command.** The project uses `uv` for package
management. Do not use `pip`, `python`, or `pytest` directly.

```sh
# Run tests
uv run pytest

# Run tests with coverage
uv run pytest --cov=src tests --verbose

# Lint (with autofix)
uv run ruff check . --fix

# Format
uv run ruff format .

# Type check
uv run ty check .

# Run all agent checks (lint + typecheck + sergey)
just agentcheck
```

## Testing

- Framework: `pytest` with `pytest-asyncio` for async tests.
- Test path: `tests/` (configured in `pyproject.toml`).
- Source path is `src/` (configured via `pythonpath = ["src"]`).
- No `conftest.py`; fixtures are defined inline in test files.
- Run the full suite with `uv run pytest`. Run a single file with
  `uv run pytest tests/test_document.py`.

## Linting and Formatting

- **Ruff** handles both linting and formatting. Config selects `ALL` rules with
  specific ignores for formatter conflicts. Google-style docstrings are required.
- **ty** is the primary type checker (not mypy).
- **sergey-lint** runs additional custom checks.
- Target Python version: 3.12. The project supports 3.12, 3.13, and 3.14.

## Code Style

- **src layout**: all library code lives under `src/cmi_docx/`.
- **Docstrings**: Google convention. Every public function/class needs one.
- **Formatting**: double quotes, 4-space indent, 88-char line length.
- **Types**: use specific type hints everywhere; avoid `Any`.
- **Dataclasses**: used extensively for styles, find results, and declarative
  components.
- **Async**: the declarative API uses `asyncio` with `async`/`await` throughout.
  Components have an async `resolve()` method and children are gathered concurrently.
- **YAML files**: must use `.yaml` extension, never `.yml` (enforced by pre-commit).

## Key Patterns

- **Wrapper/extension classes**: `ExtendDocument`, `ExtendParagraph`, `ExtendRun`,
  `ExtendTable`, `ExtendCell` wrap `python-docx` objects and add functionality.
- **Declarative components**: `Component` base class in `declarative/base.py`
  supports lazy field evaluation, conditional rendering, and async resolution.
- **Barrel exports**: `__init__.py` re-exports all public symbols so consumers
  can do `from cmi_docx import ExtendDocument`.

## Dependencies

- **Runtime**: `python-docx` (>=1.1.2), `lxml` (>=6.0.2).
- **Dev**: `pytest`, `pytest-cov`, `pytest-asyncio`, `ruff`, `ty`, `pdoc`, `prek`,
  `sergey-lint`.
- Do not add new dependencies without justification.

## Before Submitting Changes

1. `uv run ruff check . --fix` -- fix lint issues.
2. `uv run ruff format .` -- format code.
3. `uv run ty check .` -- ensure no type errors.
4. `uv run pytest` -- ensure all tests pass.

Or run `just agentcheck` for the lint/typecheck checks, then
`uv run pytest` for tests.
