"""Module for extending python-docx Run objects."""

from typing import Any

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
        self._replacement_done = False

    @property
    def runs(self) -> list[docx_paragraph.Run]:
        """Returns the runs containing the text."""
        return self.paragraph.runs[self.run_indices[0] : self.run_indices[1] + 1]

    def replace(self, replace: str, style: dict[str, Any] | None = None) -> None:
        """Replaces the text in the runs with the replacement text.

        Args:
            replace: The text to replace.
            style: The style to apply to the replacement text. If None, the
                replacement text will have the same style as the first character of the
                original text.
        """
        if self._replacement_done:
            msg = "Cannot use a FindRun replacement more than once."
            raise ValueError(msg)

        if style is None:
            self._replace_without_style(replace)
        else:
            self._replace_with_style(replace, style)
        self._replacement_done = True
        return

    def _replace_without_style(self, replace: str) -> None:
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

    def _replace_with_style(self, replace: str, style: dict[str, Any]) -> None:
        """Replaces the text in the runs with the replacement text and style.

        Args:
            replace: The text to replace.
            style: The style to apply to the replacement text.
        """
        start = self.character_indices[0]
        end = self.character_indices[1]

        pre, post = self.runs[0].text[:start], self.runs[0].text[end:]
        self.runs[0].text = pre

        new_run = self.paragraph._element._new_r()
        new_run.text = replace
        self.paragraph.runs[self.run_indices[0]]._element.addnext(new_run)
        ExtendRun(self.paragraph.runs[self.run_indices[0] + 1]).format(**style)

        post_run = self.paragraph._element._new_r()
        post_run.text = post
        self.paragraph.runs[self.run_indices[0] + 1]._element.addnext(post_run)
        pre_style = ExtendRun(self.paragraph.runs[self.run_indices[0]]).get_format()
        ExtendRun(self.paragraph.runs[self.run_indices[0] + 2]).format(**pre_style)

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
        if superscript is not None:
            self.run.font.superscript = superscript
        if subscript is not None:
            self.run.font.subscript = subscript
        if font_size is not None:
            self.run.font.size = font_size
        if font_rgb is not None:
            self.run.font.color.rgb = shared.RGBColor(*font_rgb)

    def get_format(self) -> dict[str, Any]:
        """Returns the formatting of the run.

        Returns:
            The formatting of the run.
        """
        return {
            "bold": self.run.bold,
            "italics": self.run.italic,
            "underline": self.run.underline,
            "superscript": self.run.font.superscript,
            "subscript": self.run.font.subscript,
            "font_size": self.run.font.size,
            "font_rgb": self.run.font.color.rgb,
        }
