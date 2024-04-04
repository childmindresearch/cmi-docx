"""Module for extending python-docx Run objects."""

from docx import shared
from docx.text import paragraph as docx_paragraph


class FindRun:
    """Data class for maintaing find results in runs.

    Attributes:
        paragraph: The paragraph containing the text.
        run_indices: The run indices of the text needle's start and end.
        character_indices: The character indices of the text in the runs.
            The first index is the start of the text in the first run containing
            the text. The second index is the end of the text in the last run.
    """

    def __init__(
        self,
        paragraph: docx_paragraph.Paragraph,
        run_indices: tuple[int, int],
        character_indices: tuple[int, int],
    ) -> None:
        """Initializes a FindRun object.

        Args:
            paragraph: The paragraph containing the text.
            run_indices: The run indices of the text needle's start and end.
            character_indices: The character indices of the text in the runs.
        """
        self.paragraph = paragraph
        self.run_indices = run_indices
        self.character_indices = character_indices

    @property
    def runs(self) -> list[docx_paragraph.Run]:
        """Returns the runs containing the text."""
        return self.paragraph.runs[self.run_indices[0] : self.run_indices[1] + 1]

    def replace(self, replace: str) -> None:
        """Replaces the text in the runs with the replacement text.

        Args:
            replace: The text to replace.
        """
        start = self.character_indices[0]
        end = self.character_indices[1]

        if len(self.runs) == 1:
            self.runs[0].text = (
                self.runs[0].text[:start] + replace + self.runs[0].text[end:]
            )
        else:
            self.runs[0].text = self.runs[0].text[:start] + replace
            for run in self.runs[1:-1]:
                run.clear()
            self.runs[-1].text = self.runs[-1].text[end:]
        self.character_indices = (start, start + len(replace))
        self.run_indices = (self.run_indices[0], self.run_indices[0])

    def __lt__(self, other: "FindRun") -> bool:
        """Sorts FindRun in order of appearance in the paragraph.

        Makes FindRun objects sortable.

        Args:
            other: The other FindRun object.

        Returns:
            True if the character index of the first run is less than the
            character index of the other run.
        """
        if self.paragraph != other.paragraph:
            msg = "Cannot compare FindRun objects from different paragraphs."
            raise ValueError(msg)

        if self.run_indices[0] == other.run_indices[0]:
            return self.character_indices[0] < other.character_indices[0]
        return self.run_indices[0] < other.run_indices[0]


class ExtendRun:
    """Extends a python-docx Word run with additional functionality."""

    def __init__(self, run: docx_paragraph.Run) -> None:
        """Initializes an ExtendRun object.

        Args:
            run: The run to extend.
        """
        self.run = run

    def format(
        self,
        *,
        bold: bool | None = None,
        italics: bool | None = None,
        underline: bool | None = None,
        strike: bool | None = None,
        superscript: bool | None = None,
        subscript: bool | None = None,
        font_size: int | None = None,
        font_rgb: tuple[int, int, int] | None = None,
    ) -> None:
        """Formats a run in a Word document.

        Args:
            bold: Whether to bold the run.
            italics: Whether to italicize the run.
            underline: Whether to underline the run.
            strike: Whether to strike through the run.
            superscript: Whether to superscript the run.
            subscript: Whether to subscript the run.
            font_size: The font size of the run.
            font_rgb: The font color of the run.
        """
        if superscript and subscript:
            msg = "Cannot have superscript and subscript at the same time."
            raise ValueError(msg)

        if bold is not None:
            self.run.bold = bold
        if italics is not None:
            self.run.italic = italics
        if underline is not None:
            self.run.underline = underline
        if strike is not None:
            self.run.strike = strike
        if superscript is not None:
            self.run.font.superscript = superscript
        if subscript is not None:
            self.run.font.subscript = subscript
        if font_size is not None:
            self.run.font.size = font_size
        if font_rgb is not None:
            self.run.font.color.rgb = shared.RGBColor(*font_rgb)
