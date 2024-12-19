"""Module for extending python-docx Paragraph objects."""

import bisect
import dataclasses
import itertools
import re

from docx.text import paragraph as docx_paragraph
from docx.text import run as docx_run

from cmi_docx import run, styles


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
                cumulative_run_lengths[:-1],
                occurence[1] - 1,  # -1 as the range does not include the last character
                lo=start_run,
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
        self, needle: str, replace: str, style: styles.RunStyle | None = None
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

    def insert_run(self, index: int, text: str, style: styles.RunStyle) -> docx_run.Run:
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

        run.ExtendRun(self.paragraph.runs[index]).format(style)
        return self.paragraph.runs[index]

    def format(
        self,
        style: styles.ParagraphStyle,
    ) -> None:
        """Formats a paragraph in a Word document.

        Args:
            style: The style to apply to the paragraph.
        """
        if style.line_spacing is not None:
            self.paragraph.paragraph_format.line_spacing = style.line_spacing

        if style.alignment is not None:
            self.paragraph.alignment = style.alignment

        if style.space_before is not None:
            self.paragraph.paragraph_format.space_before = style.space_before

        if style.space_after is not None:
            self.paragraph.paragraph_format.space_after = style.space_after

        for paragraph_run in self.paragraph.runs:
            run.ExtendRun(paragraph_run).format(
                styles.RunStyle(
                    bold=style.bold,
                    italic=style.italic,
                    font_size=style.font_size,
                    font_rgb=style.font_rgb,
                )
            )
