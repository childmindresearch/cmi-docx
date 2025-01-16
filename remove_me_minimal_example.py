"""Minimal example of comment bug.

Notice how in 'after-replace' the comment is only on the snippet before the
replacement.
"""

import docx

import cmi_docx

document = docx.Document()
paragraph = document.add_paragraph("Hello Florian, how are you.")

cmi_docx.add_comment(
    document,
    (paragraph, paragraph),
    "Reinder",
    "Help us Flori-wan, you're our only hope.",
)

extended_paragraph = cmi_docx.ExtendParagraph(paragraph)

document.save("before_replace.docx")
extended_paragraph.replace_between(6, 13, "Elon")

document.save("after_replace.docx")
