""".. include:: ../../README.md"""  # noqa: D415

# sergey: disable-file: IMP001 # Allow importing non-modules for barrel export.

from cmi_docx.comment import add_comment  # noqa: F401
from cmi_docx.declarative import (  # noqa: F401
    Break,
    Component,
    Document,
    Footer,
    Header,
    ImageRun,
    Paragraph,
    Section,
    SectionProperties,
    Tab,
    Table,
    TableCell,
    TableRow,
    TextRun,
)
from cmi_docx.document import ExtendDocument  # noqa: F401
from cmi_docx.paragraph import ExtendParagraph, FindParagraph  # noqa: F401
from cmi_docx.run import ExtendRun, FindRun  # noqa: F401
from cmi_docx.styles import (  # noqa: F401
    CellBorder,
    CellStyle,
    ParagraphStyle,
    RunStyle,
    TableSections,
    TableStyle,
)
from cmi_docx.table import ExtendCell, ExtendTable  # noqa: F401
