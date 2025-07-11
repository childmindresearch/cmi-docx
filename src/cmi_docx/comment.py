"""Module for handling Word comments.

Code based on sample code in https://github.com/python-openxml/python-docx/issues/93.
"""

import dataclasses
import datetime
from xml.etree import ElementTree

from docx import document, oxml
from docx.opc import constants as docx_constants
from docx.opc import packuri, part
from docx.oxml import ns
from docx.oxml.text import run as docx_run
from docx.text import paragraph, run
from lxml import etree

_COMMENTS_PART_DEFAULT_XML_BYTES = (
    b"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r
<w:comments
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
    xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:lc="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"
    xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
    xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
    xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
    xmlns:cr="http://schemas.microsoft.com/office/comments/2020/reactions">
</w:comments>
"""
).strip()


def add_comment(
    docx_doc: document.Document,
    location: tuple[paragraph.Paragraph | run.Run, paragraph.Paragraph | run.Run]
    | paragraph.Paragraph
    | run.Run,
    author: str,
    text: str,
) -> None:
    """Adds a comment to Word document.

    There is a known bug where a range of locations can be provided
    where the start comes after the end.

    Args:
        docx_doc: A Word document.
        location: The paragraph and/or run object to place the comment on.
            May also be a tuple of these where the first element is the start
            and the second element is the end.
        author: Name of the comment author.
        text: Content of the comment.
    """
    if not isinstance(location, tuple):
        elements = (location._element, location._element)
    elif len(location) > 2:
        raise ValueError("Location must be a single element or a tuple of two.")
    else:
        elements = (location[0]._element, location[1]._element)

    try:
        comments_part = docx_doc.part.part_related_by(
            docx_constants.RELATIONSHIP_TYPE.COMMENTS
        )
    except KeyError:
        # No comments part found.
        comments_part = part.Part(
            partname=packuri.PackURI("/word/comments.xml"),
            content_type=docx_constants.CONTENT_TYPE.WML_COMMENTS,
            blob=_COMMENTS_PART_DEFAULT_XML_BYTES,
            package=docx_doc.part.package,
        )
        docx_doc.part.relate_to(
            comments_part, docx_constants.RELATIONSHIP_TYPE.COMMENTS
        )

    comments_xml = oxml.parse_xml(comments_part.blob)

    # Create the comment
    comment_id = str(len(comments_xml.findall(ns.qn("w:comment"))))
    comment_element = oxml.OxmlElement("w:comment")
    comment_element.set(ns.qn("w:id"), comment_id)
    comment_element.set(ns.qn("w:author"), author)
    comment_element.set(ns.qn("w:date"), datetime.datetime.now().isoformat())

    # Create the text element for the comment
    for para in text.split("\n"):
        comment_paragraph = oxml.OxmlElement("w:p")
        comment_run = oxml.OxmlElement("w:r")
        comment_text_element = oxml.OxmlElement("w:t")
        comment_text_element.text = para
        comment_run.append(comment_text_element)
        comment_paragraph.append(comment_run)
        comment_element.append(comment_paragraph)
        comments_xml.append(comment_element)
    comments_part._blob = ElementTree.tostring(comments_xml)

    # Create the commentRangeStart and commentRangeEnd elements
    comment_range_start = oxml.OxmlElement("w:commentRangeStart")
    comment_range_start.set(ns.qn("w:id"), comment_id)
    comment_range_end = oxml.OxmlElement("w:commentRangeEnd")
    comment_range_end.set(ns.qn("w:id"), comment_id)

    # Add the commentRangeStart to the first element and commentRangeEnd to
    # the last element
    elements[0].insert(0, comment_range_start)
    elements[1].append(comment_range_end)

    # Add the comment reference to each element in the range
    # for element in elements:
    comment_reference = oxml.OxmlElement("w:r")
    comment_reference_run = oxml.OxmlElement("w:r")
    comment_reference_run_properties = oxml.OxmlElement("w:rPr")
    comment_reference_run_properties.append(
        oxml.OxmlElement("w:rStyle", {ns.qn("w:val"): "CommentReference"})
    )
    comment_reference_run.append(comment_reference_run_properties)
    comment_reference_element = oxml.OxmlElement("w:commentReference")
    comment_reference_element.set(ns.qn("w:id"), comment_id)
    comment_reference_run.append(comment_reference_element)
    comment_reference.append(comment_reference_run)

    elements[0].append(comment_reference)


@dataclasses.dataclass
class CommentRange:
    """All data representing a comment range.

    Used for temporal removal and reinsertion via `CommentPreserver`.
    """

    id: str
    start_index: int
    end_index: int


class CommentPreserver:
    """Utility to remove and reinsert comment ranges.

    As some document modifications will affect comment ranges
    in weird ways this allows to preserve them.
    """

    def __init__(self, paragraph_element: etree._Element) -> None:
        """Create CommentPreserver utility."""
        self.paragraph = paragraph_element
        self.ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    def _get_text_length_before_element(self, element: etree._Element) -> int:
        """Calculate total text length before given element."""
        length = 0
        for elem in self.paragraph.iter():
            if elem == element:
                break
            if elem.tag == f"{{{self.ns['w']}}}t":
                length += len(elem.text or "")
        return length

    def extract_comments(self) -> list[CommentRange]:
        """Extract all comment ranges from the paragraph."""
        comments = []
        start_elements = {}

        for elem in self.paragraph.iter():
            if elem.tag == f"{{{self.ns['w']}}}commentRangeStart":
                comment_id = elem.get(f"{{{self.ns['w']}}}id")
                start_pos = self._get_text_length_before_element(elem)
                start_elements[comment_id] = start_pos

            if elem.tag == f"{{{self.ns['w']}}}commentRangeEnd":
                comment_id = elem.get(f"{{{self.ns['w']}}}id")
                if comment_id in start_elements:
                    end_pos = self._get_text_length_before_element(elem)
                    comments.append(
                        CommentRange(
                            id=comment_id,
                            start_index=start_elements[comment_id],
                            end_index=end_pos,
                        )
                    )

        return comments

    def strip_comments(self) -> None:
        """Remove all comment-related elements from the paragraph."""
        to_remove = []

        def should_remove_run(run: docx_run.CT_R) -> bool:
            """Helper to determine if a run should be removed."""
            # Check for nested runs
            if any(child.tag == f"{{{self.ns['w']}}}r" for child in run):
                return False
            # Check if empty
            if len(run) == 0:
                return True
            # Check if it's a comment reference run
            if (
                len(run) == 1
                and run[0].tag == f"{{{self.ns['w']}}}rPr"
                and any(
                    "CommentReference" in e.get(f"{{{self.ns['w']}}}val", "")
                    for e in run[0].iter()
                )
            ):
                return True
            # Check if it only contains rPr and no text
            has_text = any(child.tag == f"{{{self.ns['w']}}}t" for child in run)
            has_only_props = all(
                child.tag
                in [f"{{{self.ns['w']}}}rPr", f"{{{self.ns['w']}}}commentReference"]
                for child in run
            )
            return not has_text and has_only_props

        for elem in self.paragraph.iter():
            # Check for comment range elements
            if any(tag in elem.tag for tag in ["commentRange", "commentReference"]):
                to_remove.append(elem)
            # Check for runs that should be removed
            elif elem.tag == f"{{{self.ns['w']}}}r" and should_remove_run(elem):
                to_remove.append(elem)

        for elem in to_remove:
            parent = elem.getparent()
            if parent is not None:
                parent.remove(elem)

        # Clean up any remaining empty runs after removing elements
        for elem in list(self.paragraph):
            if elem.tag == f"{{{self.ns['w']}}}r" and should_remove_run(elem):
                self.paragraph.remove(elem)

    @staticmethod
    def adjust_range_positions(
        comments: list[CommentRange],
        edit_start: int,
        edit_end: int,
        new_text_length: int,
    ) -> list[CommentRange]:
        """Adjust comment range positions based on text edit."""
        adjusted = []
        change_in_length = new_text_length - (edit_end - edit_start)

        for comment in comments:
            new_comment = CommentRange(
                id=comment.id,
                start_index=comment.start_index,
                end_index=comment.end_index,
            )

            # Comment completely before edit
            if comment.end_index <= edit_start:
                adjusted.append(new_comment)
                continue

            # Comment completely after edit
            if comment.start_index >= edit_end:
                new_comment.start_index += change_in_length
                new_comment.end_index += change_in_length
                adjusted.append(new_comment)
                continue

            # Comment overlaps with edit
            if comment.start_index < edit_start:
                if comment.end_index > edit_end:
                    # Edit is fully within comment range - adjust only the end position
                    new_comment.end_index = comment.end_index + change_in_length
                else:
                    # Edit starts in comment but extends beyond
                    new_comment.end_index = comment.end_index + change_in_length
            elif comment.end_index > edit_end:
                # Edit starts before comment but ends within - adjust both positions
                new_comment.start_index = edit_start
                new_comment.end_index = comment.end_index + change_in_length
            else:
                # Edit fully encompasses comment - maintain relative positions
                new_comment.start_index = edit_start
                new_comment.end_index = edit_start + new_text_length

            adjusted.append(new_comment)

        return adjusted

    def restore_comments(self, comments: list[CommentRange]) -> None:
        """Restore comment ranges to the paragraph at adjusted positions."""
        # First, ensure any leftover comment references are removed
        self.strip_comments()

        for comment in comments:
            # Create and insert comment range start
            start_elem = etree.Element(f"{{{self.ns['w']}}}commentRangeStart")
            start_elem.set(f"{{{self.ns['w']}}}id", comment.id)
            self._insert_at_position(start_elem, comment.start_index)

            # Create and insert comment range end
            end_elem = etree.Element(f"{{{self.ns['w']}}}commentRangeEnd")
            end_elem.set(f"{{{self.ns['w']}}}id", comment.id)
            self._insert_at_position(end_elem, comment.end_index)

            # Create and insert comment reference as last element
            ref_run = etree.Element(f"{{{self.ns['w']}}}r")
            ref_props = etree.SubElement(ref_run, f"{{{self.ns['w']}}}rPr")
            style = etree.SubElement(ref_props, f"{{{self.ns['w']}}}rStyle")
            style.set(f"{{{self.ns['w']}}}val", "CommentReference")
            ref = etree.SubElement(ref_run, f"{{{self.ns['w']}}}commentReference")
            ref.set(f"{{{self.ns['w']}}}id", comment.id)

            # Clean up any empty runs before adding the comment reference
            for elem in list(self.paragraph):
                if elem.tag == f"{{{self.ns['w']}}}r" and len(elem) == 0:
                    self.paragraph.remove(elem)

            # Always append reference at the end of paragraph
            self.paragraph.append(ref_run)

    def _insert_at_position(self, elem: etree._Element, text_pos: int) -> None:
        """Insert element at the specified text position."""
        current_pos = 0

        # First collect all text-containing runs and their positions
        runs_with_text = []
        for child in self.paragraph:
            if child.tag == f"{{{self.ns['w']}}}r":
                text_elems = child.findall(f".//{{{self.ns['w']}}}t")
                if text_elems:
                    text_length = sum(len(t.text or "") for t in text_elems)
                    if text_length > 0:
                        runs_with_text.append((child, current_pos, text_length))
                        current_pos += text_length

        # Handle insertion based on text position
        if not runs_with_text:
            # If no text runs, insert at beginning
            if self.paragraph:
                self.paragraph.insert(0, elem)
            else:
                self.paragraph.append(elem)
            return

        # Find the appropriate run based on text position
        insert_after = None
        for i, (run, start_pos, length) in enumerate(runs_with_text):  # noqa: F402
            if start_pos + length >= text_pos:
                if start_pos + length == text_pos:
                    # Exact match at end of run - insert after this run
                    insert_after = run
                    break
                # Position is within or at start of run - insert before this run
                insert_idx = self.paragraph.index(run)
                self.paragraph.insert(insert_idx, elem)
                return
            insert_after = run

        # If we get here, either:
        # 1. We found an exact match at end of a run and should insert after it
        # 2. Position is past all runs, so insert after last text run
        if insert_after is not None:
            insert_idx = self.paragraph.index(insert_after) + 1
            self.paragraph.insert(insert_idx, elem)
        else:
            # Fallback - append at end
            self.paragraph.append(elem)
