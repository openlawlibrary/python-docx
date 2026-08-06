"""Custom element classes for bookmark-related elements."""

from __future__ import annotations

from docx.oxml.simpletypes import ST_String, XsdUnsignedInt
from docx.oxml.xmlchemy import BaseOxmlElement, RequiredAttribute


class CT_BookmarkStart(BaseOxmlElement):
    """`<w:bookmarkStart>` element, specifying the id and name of a bookmark start."""

    id: int = RequiredAttribute("w:id", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]
    name: str = RequiredAttribute("w:name", ST_String)  # pyright: ignore[reportAssignmentType]


class CT_BookmarkEnd(BaseOxmlElement):
    """`<w:bookmarkEnd>` element, specifying the id of the bookmark it closes."""

    id: int = RequiredAttribute("w:id", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]
