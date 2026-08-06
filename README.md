# CMI-docx

[![Build](https://github.com/childmindresearch/cmi-docx/actions/workflows/test.yaml/badge.svg?branch=main)](https://github.com/childmindresearch/cmi-docx/actions/workflows/test.yaml?query=branch%3Amain)
[![codecov](https://codecov.io/gh/childmindresearch/cmi-docx/branch/main/graph/badge.svg?token=22HWWFWPW5)](https://codecov.io/gh/childmindresearch/cmi-docx)
[![Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)
![stability-stable](https://img.shields.io/badge/stability-stable-green.svg)
[![LGPL--2.1 License](https://img.shields.io/badge/license-LGPL--2.1-blue.svg)](https://github.com/childmindresearch/cmi-docx/blob/main/LICENSE)
[![pages](https://img.shields.io/badge/api-docs-blue)](https://childmindresearch.github.io/cmi-docx)

`cmi-docx` is a Python library by the [Child Mind Institute](https://childmind.org)
that extends [`python-docx`](https://python-docx.readthedocs.io/) with higher-level
tooling for `.docx` files. It provides two APIs:

| API | Use it when | Entry point |
| --- | --- | --- |
| **Declarative** | You are *building* a document from data | `cmi_docx.declarative.Document` |
| **Imperative** | You are *editing* a document that already exists | `cmi_docx.ExtendDocument` |

Both can be combined: build with the declarative API, then post-process the
resulting `python-docx` object with the imperative wrappers.

## Installation

```sh
pip install cmi-docx
```

Requires Python 3.12 or newer.

---

# Tutorial: the declarative API

This section is a step-by-step tutorial. Each step builds on the previous one.
If you read it top to bottom you will have seen every feature the declarative
API supports.

## Step 1: the mental model

You describe *what* the document contains as a tree of plain dataclasses. You
never call `add_paragraph`, `add_run`, or touch XML. Then you render the tree
once:

```text
Document
└── Section            (page setup, headers, footers)
    ├── Paragraph      (a block)
    │   ├── TextRun    (an inline piece of text)
    │   ├── Tab
    │   ├── Break
    │   └── ImageRun
    └── Table          (a block)
        └── TableRow
            └── TableCell
                └── Paragraph / Table   (cells hold blocks, recursively)
```

Three rules cover almost everything:

1. **Blocks** (`Paragraph`, `Table`) go in the `children` of a `Section`,
   `Header`, `Footer`, or `TableCell`.
2. **Inline elements** (`TextRun`, `Tab`, `Break`, `ImageRun`) go in the
   `children` of a `Paragraph`.
3. Rendering is `await doc.to_docx()`, which returns an ordinary `python-docx`
   `Document` that you save yourself.

## Step 2: your first document

```python
import asyncio

from cmi_docx import declarative


async def main() -> None:
    doc = declarative.Document(
        sections=[
            declarative.Section(
                children=[
                    declarative.Paragraph(text="Annual Report", heading=1),
                    declarative.Paragraph(text="Written for the 2025 review cycle."),
                ],
            ),
        ],
        title="Annual Report",
        creator="Jane Doe",
    )

    docx_doc = await doc.to_docx()
    docx_doc.save("output.docx")


asyncio.run(main())
```

Things to notice:

- `to_docx()` is a coroutine, so it must be awaited. `Document` itself has no
  `save()` method -- you save the `python-docx` object it returns.
- `heading=1` applies the built-in `"Heading 1"` style. Use `style="..."` for
  any other named style.
- `Document(...)` metadata keyword arguments (`title`, `creator`, `subject`,
  `keywords`, `category`, `version`, `comments`, `description`) map onto the
  file's core properties. `created` and `modified` are always set to the current
  UTC time.

## Step 3: formatting text with `TextRun`

A `Paragraph` takes **either** `text` **or** `children` -- never both, never
neither. Passing both, or passing neither, raises `ValueError` immediately at
construction time.

Use `text` for a single uniform run. Use `children` as soon as one paragraph
mixes formatting:

```python
declarative.Paragraph(
    children=[
        declarative.TextRun(text="Result: "),
        declarative.TextRun(text="significant", bold=True, color=(0, 128, 0)),
        declarative.TextRun(text=" (p < .05)", italic=True, size=9),
    ],
)
```

`TextRun` formatting fields: `bold`, `italic`, `underline`, `font` (name),
`size` (points), `color` (`(r, g, b)`, each 0-255), `superscript`, `subscript`,
`strike`, `all_caps`, `small_caps`.

`Paragraph` block formatting fields:

| Field | Unit / type | Notes |
| --- | --- | --- |
| `heading` | `int` | Shorthand for `"Heading {n}"` |
| `style` | `str` | Any named paragraph style |
| `alignment` | `WD_PARAGRAPH_ALIGNMENT` | From `docx.enum.text` |
| `spacing_before`, `spacing_after` | points | |
| `line_spacing` | `float` | Multiplier, e.g. `1.5` |
| `left_indent`, `right_indent`, `first_line_indent` | **points** | Not inches |
| `keep_together`, `keep_with_next`, `page_break_before`, `widow_control` | `bool` | Pagination control |

```python
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

declarative.Paragraph(
    text="A centred, generously spaced quotation.",
    alignment=WD_PARAGRAPH_ALIGNMENT.CENTER,
    line_spacing=1.5,
    spacing_after=12,
    left_indent=36,  # points, = 0.5 inch
)
```

## Step 4: tabs, breaks, and images

These are inline elements, so they live in `Paragraph.children`:

```python
declarative.Paragraph(
    children=[
        declarative.TextRun(text="Name"),
        declarative.Tab(),
        declarative.TextRun(text="Score"),
        declarative.Break(),                    # line break (default)
        declarative.TextRun(text="Next line"),
        declarative.Break(type="page"),         # page break
        declarative.TextRun(text="Next page"),
    ],
)
```

`Break.type` accepts `"line"` (default), `"page"`, `"column"`, and
`"textWrapping"`. Anything other than `"page"` or `"column"` produces a plain
line break.

Images take bytes, a path, or a coroutine returning bytes:

```python
import pathlib

declarative.Paragraph(
    children=[
        declarative.ImageRun(
            data=pathlib.Path("chart.png"),
            transformation={"width": 300},  # points; omit height to keep aspect ratio
        ),
    ],
)
```

## Step 5: tables

Tables are strictly nested: `Table` -> `TableRow` -> `TableCell` -> blocks. A
cell's content is always a list of `Paragraph` (or nested `Table`) objects, not
a bare string.

```python
declarative.Table(
    rows=[
        declarative.TableRow(
            children=[
                declarative.TableCell(
                    children=[declarative.Paragraph(text="Metric")],
                    background_color=(0, 51, 102),
                ),
                declarative.TableCell(
                    children=[declarative.Paragraph(text="Value")],
                    background_color=(0, 51, 102),
                ),
            ],
        ),
        declarative.TableRow(
            children=[
                declarative.TableCell(children=[declarative.Paragraph(text="Revenue")]),
                declarative.TableCell(children=[declarative.Paragraph(text="$1M")]),
            ],
        ),
    ],
    style="Table Grid",
    column_widths=[2880, 1440],  # twips: 1440 = 1 inch
)
```

### Column widths and layout

- `column_widths` is a list in **twips** (1440 per inch, ~567 per cm). Its
  length must equal the table's total column count or `to_docx()` raises
  `ValueError`. Setting it implies fixed layout.
- `layout="fixed"` disables autofit without specifying widths.
- `layout="autofit"` wins over `column_widths`: autofit stays on and the widths
  are *not* applied.

### Merging cells

Horizontal merge uses `grid_span`; vertical merge uses `vmerge`. For a vertical
merge, the first cell is `"restart"` and each continuation is `"continue"`
(continuation cells normally have `children=None`).

```python
declarative.Table(
    rows=[
        declarative.TableRow(
            children=[
                declarative.TableCell(
                    children=[declarative.Paragraph(text="Spans two columns")],
                    grid_span=2,
                ),
                declarative.TableCell(
                    children=[declarative.Paragraph(text="Tall cell")],
                    vmerge="restart",
                ),
            ],
        ),
        declarative.TableRow(
            children=[
                declarative.TableCell(children=[declarative.Paragraph(text="A")]),
                declarative.TableCell(children=[declarative.Paragraph(text="B")]),
                declarative.TableCell(children=None, vmerge="continue"),
            ],
        ),
    ],
)
```

The table's column count is computed as the maximum of
`sum(grid_span or 1)` across all rendered rows, so rows may legitimately hold
different numbers of `TableCell` objects.

### Borders

Two border types exist, and they differ:

- `TableBorder` applies to the whole table. `side` is one of `"top"`,
  `"left"`, `"bottom"`, `"right"`, `"insideH"`, `"insideV"`. `val` is
  `"single"` only.
- `CellBorder` applies to one cell. `side` is one of `"top"`, `"bottom"`,
  `"start"` (left in LTR), `"end"` (right in LTR). `val` is `"single"` or
  `"dashed"`.

For both, `sz` is in eighths of a point (`sz=8` is 1pt) and `color` is an RGB
tuple.

```python
declarative.Table(
    rows=[...],
    borders=[
        declarative.TableBorder(side="top", sz=8, color=(0, 0, 0)),
        declarative.TableBorder(side="insideH", sz=4, color=(200, 200, 200)),
    ],
)

declarative.TableCell(
    children=[declarative.Paragraph(text="Highlighted")],
    borders=[declarative.CellBorder(side="bottom", sz=4, val="dashed", color=(255, 0, 0))],
    background_color=(255, 255, 200),
    vertical_alignment="center",  # "top" | "center" | "bottom"
)
```

## Step 6: sections, page setup, headers, and footers

A `Section` is a page-setup boundary. Each declarative `Section` maps to exactly
one Word section, so two declarative sections produce a two-section document.

```python
from docx.shared import Inches

declarative.Section(
    children=[declarative.Paragraph(text="Wide table on a landscape page")],
    properties=declarative.SectionProperties(
        page_orientation="landscape",
        page_size={"width": Inches(8.5), "height": Inches(11)},
        page_margins={"top": Inches(1), "bottom": Inches(1), "left": Inches(0.75)},
    ),
)
```

- `page_size` and `page_margins` values are in EMU. Always use
  `docx.shared.Inches`, `Cm`, or `Pt` instead of raw integers.
- `page_margins` keys are all optional; omitted sides keep their defaults.
- `page_orientation` swaps `page_size` dimensions as needed, so you can pass
  portrait dimensions together with `"landscape"` and get the expected result.

Headers and footers are keyed by type -- `"default"`, `"first"`, or `"even"`:

```python
declarative.Section(
    children=[declarative.Paragraph(text="Body text")],
    headers={
        "default": declarative.Header(
            children=[declarative.Paragraph(text="Confidential")],
        ),
    },
    footers={
        "default": declarative.Footer(
            children=[declarative.Paragraph(text="Child Mind Institute")],
        ),
    },
)
```

Word's default header/footer already contains one empty paragraph, so header
content is appended after it. Similarly, the section break between two
declarative sections introduces one empty body paragraph. When asserting on
output, filter out blank paragraphs:

```python
content = [para for para in docx_doc.paragraphs if para.text.strip()]
```

## Step 7: defining named styles

Rather than repeating formatting on every component, define a style once at the
document level and reference it by name.

```python
doc = declarative.Document(
    sections=[
        declarative.Section(
            children=[declarative.Paragraph(text="Body copy", style="Body")],
        ),
    ],
    styles=[
        declarative.ParagraphStyleDefinition(
            name="Body",
            font="Arial",
            font_size=11,
            line_spacing=1.15,
            spacing_after=6,
            left_indent=0.5,  # inches here, unlike Paragraph.left_indent
        ),
        declarative.ParagraphStyleDefinition(name="Heading 1", font="Arial"),
    ],
)
```

Two behaviours matter here:

- `ParagraphStyleDefinition` is **get-or-create**. If the name already exists
  (including built-ins such as `"Heading 1"`) it is modified in place, which is
  the supported way to restyle built-in headings. Setting `font` also strips the
  theme-font attributes so your explicit font is not silently overridden.
- Its indent fields (`left_indent`, `right_indent`, `first_line_indent`) are in
  **inches**, whereas the same-named fields on `Paragraph` are in **points**.
  This is the single easiest thing to get wrong.

`TableStyleDefinition` always creates a new style and raises `ValueError` if the
name is taken, so do not reuse built-in names like `"Table Grid"`. Per-region
formatting uses `TableSectionFormat`:

```python
declarative.TableStyleDefinition(
    name="ReportTable",
    base_style="Table Grid",
    whole_table=declarative.TableSectionFormat(font="Arial", font_size=10),
    first_row=declarative.TableSectionFormat(
        bold=True,
        color=(255, 255, 255),
        background=(0, 51, 102),
    ),
    banding_1_row=declarative.TableSectionFormat(background=(240, 240, 240)),
)
```

Available regions: `whole_table`, `first_row`, `last_row`, `first_column`,
`last_column`, `banding_1_row`, `banding_2_row`, `banding_1_column`,
`banding_2_column`, `top_left_cell`, `top_right_cell`, `bottom_left_cell`,
`bottom_right_cell`.

## Step 8: conditional rendering

Every component accepts a keyword-only `condition`: a zero-argument callable
returning `bool`. When it returns `False` the component and its entire subtree
are skipped.

```python
include_appendix = False

declarative.Section(
    children=[
        declarative.Paragraph(text="Always present"),
        declarative.Paragraph(
            text="Appendix A",
            heading=2,
            condition=lambda: include_appendix,
        ),
    ],
)
```

`condition` works on `Section`, `Paragraph`, `TextRun`, `Tab`, `Break`,
`ImageRun`, `Table`, `TableRow`, `TableCell`, `Header`, and `Footer`. A
false-conditioned `TableRow` or `TableCell` is removed from the rendered grid
rather than left blank.

Note that `condition` must be a **callable**, not a boolean:
`condition=lambda: flag`, not `condition=flag`. It is evaluated at resolve time,
so it sees the value of `flag` as of `to_docx()`, not as of construction.

`Document` is not a component and therefore has no `condition`.

## Step 9: lazy fields

Any field can be given a zero-argument callable instead of a value. The callable
is invoked during resolution -- but only if `condition` passed. This is how you
avoid paying for content that will not be rendered.

```python
def build_appendix() -> declarative.BlockChildren:
    # Expensive: never runs when the condition is False.
    return [declarative.Paragraph(text=row) for row in load_thousands_of_rows()]


declarative.Section(
    children=build_appendix,          # note: no parentheses
    condition=lambda: include_appendix,
)
```

Lazy values work for text too:

```python
declarative.Paragraph(text=lambda: f"Generated at {now()}")
```

`condition` itself is exempt from this materialisation, which is why it stays a
callable.

## Step 10: async content

The API is async-first so that content requiring I/O can be fetched
concurrently. Anywhere a component is accepted, a coroutine resolving to that
component is also accepted, and all of them are gathered with
`asyncio.gather`.

```python
import asyncio

from cmi_docx import declarative


async def fetch_summary() -> declarative.Paragraph:
    row = await database.fetch_summary()      # ~100 ms
    return declarative.Paragraph(text=row.text)


async def fetch_chart() -> bytes:
    return await http_client.get_png()        # ~100 ms


async def main() -> None:
    doc = declarative.Document(
        sections=[
            declarative.Section(
                children=[
                    declarative.Paragraph(text="Summary", heading=1),
                    fetch_summary(),          # coroutine as a block child
                    declarative.Paragraph(
                        children=[declarative.ImageRun(data=fetch_chart())],
                    ),
                ],
            ),
        ],
    )

    docx_doc = await doc.to_docx()             # both awaits overlap: ~100 ms total
    docx_doc.save("output.docx")


asyncio.run(main())
```

Coroutines are allowed for `Section.children`, `Header.children`,
`Footer.children`, `Paragraph.children`, `Paragraph.text`, `TextRun.text`,
`TextRun.comment_text`, `TextRun.comment_author`, `Table.rows`,
`TableRow.children`, `TableCell.children`, and `ImageRun.data`.

One exception: the `Document.sections` list must contain real `Section`
objects. A coroutine there fails with `AttributeError`. Await it first:

```python
sections = await asyncio.gather(build_intro(), build_body())
doc = declarative.Document(sections=list(sections))
```

## Step 11: Word comments

Set `comment_text` on a `Paragraph` (anchors the whole paragraph) or on a
`TextRun` (anchors just that run). `Document(comment_author=...)` sets the
default author; `comment_author` on the paragraph or run overrides it.

```python
doc = declarative.Document(
    sections=[
        declarative.Section(
            children=[
                declarative.Paragraph(
                    text="This figure needs a citation.",
                    comment_text="Add the 2024 source.",
                ),
                declarative.Paragraph(
                    children=[
                        declarative.TextRun(text="Sample size was "),
                        declarative.TextRun(
                            text="n = 42",
                            comment_text="Verify against the enrolment log.",
                            comment_author="QA Reviewer",
                        ),
                    ],
                ),
            ],
        ),
    ],
    comment_author="Jane Doe",
)
```

## Step 12: rendering into a template

`DocumentTemplate` opens an existing `.docx` as the starting point, applies
find/replace, and optionally inserts your declarative content at a specific
paragraph index instead of appending.

```python
import asyncio
import pathlib

from cmi_docx import declarative


async def main() -> None:
    doc = declarative.Document(
        sections=[
            declarative.Section(
                children=[declarative.Paragraph(text="Inserted content")],
            ),
        ],
    )

    template = declarative.DocumentTemplate(
        path=pathlib.Path("template.docx"),
        replacements={"{{NAME}}": "Alice", "{{DATE}}": "2025-01-01"},
        paragraph_index=1,
    )

    docx_doc = await doc.to_docx(template=template)
    docx_doc.save("output.docx")


asyncio.run(main())
```

- `paragraph_index` counts paragraphs in the *original* template, before any
  insertion. `0` prepends, `None` (the default) appends, and an index past the
  end appends.
- `replacements` runs before insertion and covers the body, headers, footers,
  and tables -- so placeholders inside template tables are substituted too.
- `Document(sections=[])` with a template is valid and gives you a pure
  find/replace pipeline.

## Reference: units at a glance

Mixed units are the most common source of surprising output.

| Where | Field | Unit |
| --- | --- | --- |
| `TextRun` | `size` | points |
| `Paragraph` | `spacing_before`, `spacing_after` | points |
| `Paragraph` | `left_indent`, `right_indent`, `first_line_indent` | points |
| `ParagraphStyleDefinition` | `font_size`, `spacing_before`, `spacing_after` | points |
| `ParagraphStyleDefinition` | `left_indent`, `right_indent`, `first_line_indent` | **inches** |
| `ImageRun` | `transformation["width"]`, `["height"]` | points |
| `SectionProperties` | `page_size`, `page_margins` | EMU (use `docx.shared.Inches`) |
| `Table` | `column_widths` | twips (1440 = 1 inch) |
| `TableBorder`, `CellBorder` | `sz` | eighths of a point |
| Everything | `color`, `background_color`, `background` | `(r, g, b)`, 0-255 |

## Reference: currently accepted but not applied

These fields type-check and will not raise, but have no effect on the rendered
document yet. Do not rely on them:

- `Document(numbering=...)`
- `SectionProperties.page_numbering`, `.columns`, `.vertical_align`,
  `.title_page`, `.type`
- `ImageRun.type`, `ImageRun.alt_text`

Also note `Document(description=...)` and `Document(comments=...)` both write to
the same core property, so setting both means the last one wins.

---

# The imperative API

Use these wrappers to modify a `python-docx` document you already have. Each
`Extend*` class takes the corresponding `python-docx` object and adds methods to
it; the original object is mutated in place.

## Find and replace

`ExtendDocument.replace` works even when Word has split the target string across
several runs, which is the usual reason naive `run.text.replace()` fails.

```python
import docx

from cmi_docx import ExtendDocument, RunStyle

doc = docx.Document()
paragraph = doc.add_paragraph("Hello {{")
paragraph.add_run("FULL_NAME}}")  # placeholder split across two runs

ExtendDocument(doc).replace("{{FULL_NAME}}", "Jane Doe", RunStyle(bold=True))

print(doc.paragraphs[0].text)  # "Hello Jane Doe"
```

`replace` covers the body, headers, footers, and tables. The optional
`RunStyle` styles the replacement text only.

Related methods on `ExtendDocument`: `find_in_paragraphs`, `find_in_runs`,
`insert_paragraph_by_text`, `insert_paragraph_by_object`, `insert_image`, and
`all_paragraphs`.

## Paragraph and run formatting

```python
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH

from cmi_docx import ExtendParagraph, ParagraphStyle, RunStyle

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

Insert a styled run at a specific run index:

```python
paragraph = doc.add_paragraph("")
paragraph.add_run("Hello ")
paragraph.add_run("world!")

ExtendParagraph(paragraph).insert_run(1, "beautiful ", RunStyle(bold=True))

print(paragraph.text)  # "Hello beautiful world!"
```

`ExtendRun` adds `format(RunStyle)` and `get_format()`, the latter returning the
run's current formatting as a `RunStyle` -- useful for copying formatting from
one run to another.

## Tables and cells

Note that imperative `CellBorder` differs from the declarative one: it takes a
`sides` **tuple** and a hex string `color`.

```python
import docx

from cmi_docx import CellBorder, CellStyle, ExtendCell, ExtendTable, TableSections, TableStyle

doc = docx.Document()
table = doc.add_table(rows=2, cols=2)

ExtendTable(table).format(TableStyle(sections=TableSections(first_row=True)))

ExtendCell(table.cell(0, 0)).format(
    CellStyle(
        background_rgb=(0, 51, 102),
        borders=[CellBorder(sides=("top", "bottom"), sz=8, color="000000")],
    )
)
```

## Comments

```python
import docx

from cmi_docx import add_comment

document = docx.Document()
paragraph = document.add_paragraph("This needs review.")

add_comment(document, paragraph, "Reviewer", "Please check this section.")
```

`add_comment` accepts a paragraph, a single run, or a list of runs to anchor the
comment to a range. `CommentPreserver` keeps existing comment anchors valid
while text around them is edited, and is applied automatically during
find/replace.

---

# Common mistakes

| Symptom | Cause |
| --- | --- |
| `ValueError: Paragraph must have either 'text' or 'children'` | You passed neither, or both |
| Indentation off by ~72x | `Paragraph` indents are points; `ParagraphStyleDefinition` indents are inches |
| `condition` never skips anything | You passed a `bool` instead of a callable |
| Lazy children still built | The callable was invoked -- you wrote `children=build()` instead of `children=build` |
| `AttributeError: 'coroutine' object has no attribute 'resolve'` | A coroutine was placed directly in `Document(sections=...)` |
| `ValueError: column_widths length (n) must match number of columns (m)` | Count columns as `sum(grid_span or 1)` for the widest row |
| `ValueError: document already contains style '...'` | `TableStyleDefinition` cannot reuse an existing name |
| Column widths ignored | `layout="autofit"` overrides `column_widths` |
| Unexpected empty paragraphs | Section breaks and default headers each contribute one; filter with `para.text.strip()` |
| `RuntimeWarning: coroutine was never awaited` | A coroutine was attached to a component that `condition` excluded |

# Further reading

- [Full API documentation](https://childmindresearch.github.io/cmi-docx)
- [`python-docx` documentation](https://python-docx.readthedocs.io/)
- [OOXML table reference](https://ooxml.dev/docs/tables/)
