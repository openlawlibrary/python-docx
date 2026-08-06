"""Custom element classes related to footnote/endnote references (CT_FtnEdnRef)."""

from __future__ import annotations

from docx.oxml.simpletypes import ST_DecimalNumber, ST_OnOff
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute


class CT_FtnEdnRef(BaseOxmlElement):
    """`w:footnoteReference` or `w:endnoteReference` element.

    Appears in a run to mark the point in the document text a footnote or endnote is
    anchored, referencing the footnote/endnote content by `@w:id`.
    """

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    customMarkFollows: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:customMarkFollows", ST_OnOff
    )
