"""Module for handling Word comments.

Code based on sample code in https://github.com/python-openxml/python-docx/issues/93.
"""

import datetime
from xml.etree import ElementTree

from docx import document, oxml
from docx.opc import constants as docx_constants
from docx.opc import packuri, part
from docx.oxml import ns
from docx.text import paragraph, run

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
