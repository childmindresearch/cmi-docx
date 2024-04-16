"""Module for extending python-docx Paragraph objects."""

import bisect
import dataclasses
import itertools
import re
from typing import Any

from docx.enum import text
from docx.text import paragraph as docx_paragraph
from docx.text import run as docx_run

from cmi_docx import run


@dataclasses.dataclass
class FindParagraph:
    """Data class for maintaing find results in paragraphs.

    Attributes:
        paragraph: The paragraph containing the text.
        character_indices: A list of matching character indices of the text in the
            paragraph.
    """

    paragraph: docx_paragraph.Paragraph
    character_indices: list[tuple[int, int]]


class ExtendParagraph:
    """Extends a python-docx Word paragraph with additional functionality."""

    def __init__(self, paragraph: docx_paragraph.Paragraph) -> None:
        """Initializes an ExtendParagraph object.

        Args:
            paragraph: The paragraph to extend.
        """
        self.paragraph = paragraph

    def find_in_paragraph(self, needle: str) -> FindParagraph:
        """Finds the indices of a text relative to the paragraph.

        Args:
            needle: The text to find.

        Returns:
            The indices of the text in the paragraph.
        """
        within_paragraph_indices = [
            (match.start(), match.end())
            for match in re.finditer(re.escape(needle), self.paragraph.text)
        ]

        return FindParagraph(
            paragraph=self.paragraph,
            character_indices=within_paragraph_indices,
        )

    def find_in_runs(self, needle: str) -> list[run.FindRun]:
        """Finds the indices of a text relative to the paragraph's runs.

        Args:
            needle: The text to find.

        Returns:
            The indices of the text in the paragraph.
        """
        if len(needle) == 0:
            return []

        run_finds: list[run.FindRun] = []
        run_lengths = [len(run.text) for run in self.paragraph.runs]
        cumulative_run_lengths = list(itertools.accumulate(run_lengths))
        for occurence in self.find_in_paragraph(needle).character_indices:
            start_run = bisect.bisect_right(cumulative_run_lengths, occurence[0])
            end_run = bisect.bisect_right(
                cumulative_run_lengths[:-1], occurence[1], lo=start_run
            )

            start_index = (
                occurence[0] - cumulative_run_lengths[start_run - 1]
                if start_run > 0
                else occurence[0]
            )
            end_index = (
                occurence[1] - cumulative_run_lengths[end_run - 1]
                if end_run > 0
                else occurence[1]
            )

            run_finds.append(
                run.FindRun(
                    paragraph=self.paragraph,
                    run_indices=(start_run, end_run),
                    character_indices=(start_index, end_index),
                )
            )
        return run_finds

    def replace(
        self, needle: str, replace: str, style: dict[str, Any] | None = None
    ) -> None:
        """Finds and replaces text in a Word paragraph.

        Args:
            needle: The text to find.
            replace: The text to replace.
            style: The style to apply to the replacement text.
        """
        run_finder = self.find_in_runs(needle)
        run_finder.sort(
            key=lambda x: (x.run_indices[0], x.character_indices[0]), reverse=True
        )

        for run_find in run_finder:
            run_find.replace(replace, style)

    def insert_run(self, index: int, text: str, style: dict[str, Any]) -> docx_run.Run:
        """Inserts a run into a paragraph.

        Args:
            index: The index to insert the run at.
            text: The text of the run.
            style: The style of the run, see run.ExtendRun.format for more details.

        Returns:
            The inserted run.
        """
        if index == len(self.paragraph.runs):
            self.paragraph.add_run(text)
        else:
            new_run = self.paragraph._element._new_r()
            new_run.text = text
            if index < 0:
                self.paragraph.runs[index]._element.addnext(new_run)
            else:
                self.paragraph.runs[index]._element.addprevious(new_run)

        run.ExtendRun(self.paragraph.runs[index]).format(**style)
        return self.paragraph.runs[index]

    def format(
        self,
        *,
        bold: bool | None = None,
        italics: bool | None = None,
        font_size: int | None = None,
        font_rgb: tuple[int, int, int] | None = None,
        line_spacing: float | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        alignment: text.WD_PARAGRAPH_ALIGNMENT | None = None,
    ) -> None:
        """Formats a paragraph in a Word document.

        Args:
            bold: Whether to bold the paragraph.
            italics: Whether to italicize the paragraph.
            font_size: The font size of the paragraph.
            font_rgb: The font color of the paragraph.
            line_spacing: The line spacing of the paragraph.
            space_before: The spacing before the paragraph.
            space_after: The spacing after the paragraph.
            alignment: The alignment of the paragraph.
        """
        if line_spacing is not None:
            self.paragraph.paragraph_format.line_spacing = line_spacing

        if alignment is not None:
            self.paragraph.alignment = alignment

        if space_before is not None:
            self.paragraph.paragraph_format.space_before = space_before

        if space_after is not None:
            self.paragraph.paragraph_format.space_after = space_after

        for paragraph_run in self.paragraph.runs:
            run.ExtendRun(paragraph_run).format(
                bold=bold,
                italics=italics,
                font_size=font_size,
                font_rgb=font_rgb,
            )
