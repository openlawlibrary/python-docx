"""The |Endnotes| object and related proxy classes."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.blkcntnr import BlockItemContainer

if TYPE_CHECKING:
    from docx.oxml.endnote import CT_Endnotes, CT_FtnEdn
    from docx.parts.endnotes import EndnotesPart


class Endnotes:
    """Collection containing the endnotes added to this document."""

    def __init__(self, endnotes: CT_Endnotes, endnotes_part: EndnotesPart):
        self._element = self._endnotes = endnotes
        self._endnotes_part = endnotes_part

    def __getitem__(self, endnote_reference_id: int) -> Endnote:
        """The |Endnote| identified by `endnote_reference_id` (the `w:id` value of its
        `w:endnote` element, matching the `w:id` of the `w:endnoteReference` that
        anchors it in the document text).

        Raises |IndexError| if no endnote has that id.
        """
        endnote = self._endnotes.get_by_id(endnote_reference_id)
        if endnote is None:
            raise IndexError
        return Endnote(endnote, self._endnotes_part)

    def __len__(self) -> int:
        """The count of `w:endnote` elements in this collection.

        Note this includes the two `separator`/`continuationSeparator` elements every
        endnotes part starts with, in addition to any real endnotes.
        """
        return len(self._endnotes)

    def add_endnote(self, endnote_reference_id: int) -> Endnote:
        """Return a newly created |Endnote| having `endnote_reference_id`.

        The new endnote is inserted at the position that keeps endnotes ordered by id;
        when an endnote with `endnote_reference_id` already exists (because a new
        endnote is being inserted ahead of it in the document text), that endnote and
        every one after it are shifted up by one id to make room.
        """
        elements = self._endnotes
        if elements.get_by_id(endnote_reference_id) is None:
            return Endnote(elements.add_endnote(endnote_reference_id), self._endnotes_part)

        for index in reversed(range(len(elements))):
            element = cast("CT_FtnEdn", elements[index])
            if element.id == endnote_reference_id:
                element.id += 1
                new_endnote = element.add_endnote_before(endnote_reference_id)
                return Endnote(new_endnote, self._endnotes_part)
            element.id += 1

        raise AssertionError("unreachable: id was confirmed present above")


class Endnote(BlockItemContainer):
    """Proxy for a single endnote in the document.

    An endnote is also a block-item container, similar to a table cell, so it can
    contain both paragraphs and tables and its paragraphs can contain rich text,
    hyperlinks, and images.
    """

    def __init__(self, e: CT_FtnEdn, endnotes_part: EndnotesPart):
        super().__init__(e, endnotes_part)
        self._e = e

    @property
    def id(self) -> int:
        """The `w:id` value uniquely identifying this endnote."""
        return self._e.id
