"""The |Footnotes| object and related proxy classes."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.blkcntnr import BlockItemContainer

if TYPE_CHECKING:
    from docx.oxml.footnote import CT_Footnotes, CT_FtnEnd
    from docx.parts.footnotes import FootnotesPart


class Footnotes:
    """Collection containing the footnotes added to this document."""

    def __init__(self, footnotes: CT_Footnotes, footnotes_part: FootnotesPart):
        self._element = self._footnotes = footnotes
        self._footnotes_part = footnotes_part

    def __getitem__(self, footnote_reference_id: int) -> Footnote:
        """The |Footnote| identified by `footnote_reference_id` (the `w:id` value of its
        `w:footnote` element, matching the `w:id` of the `w:footnoteReference` that
        anchors it in the document text).

        Raises |IndexError| if no footnote has that id.
        """
        footnote = self._footnotes.get_by_id(footnote_reference_id)
        if footnote is None:
            raise IndexError
        return Footnote(footnote, self._footnotes_part)

    def __len__(self) -> int:
        """The count of `w:footnote` elements in this collection.

        Note this includes the two `separator`/`continuationSeparator` elements every
        footnotes part starts with, in addition to any real footnotes.
        """
        return len(self._footnotes)

    def add_footnote(self, footnote_reference_id: int) -> Footnote:
        """Return a newly created |Footnote| having `footnote_reference_id`.

        The new footnote is inserted at the position that keeps footnotes ordered by
        id; when a footnote with `footnote_reference_id` already exists (because a new
        footnote is being inserted ahead of it in the document text), that footnote and
        every one after it are shifted up by one id to make room.
        """
        elements = self._footnotes
        if elements.get_by_id(footnote_reference_id) is None:
            return Footnote(elements.add_footnote(footnote_reference_id), self._footnotes_part)

        for index in reversed(range(len(elements))):
            element = cast("CT_FtnEnd", elements[index])
            if element.id == footnote_reference_id:
                element.id += 1
                new_footnote = element.add_footnote_before(footnote_reference_id)
                return Footnote(new_footnote, self._footnotes_part)
            element.id += 1

        raise AssertionError("unreachable: id was confirmed present above")


class Footnote(BlockItemContainer):
    """Proxy for a single footnote in the document.

    A footnote is also a block-item container, similar to a table cell, so it can
    contain both paragraphs and tables and its paragraphs can contain rich text,
    hyperlinks, and images.
    """

    def __init__(self, f: CT_FtnEnd, footnotes_part: FootnotesPart):
        super().__init__(f, footnotes_part)
        self._f = f

    @property
    def id(self) -> int:
        """The `w:id` value uniquely identifying this footnote."""
        return self._f.id
