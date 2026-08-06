"""Custom element classes related to endnotes (CT_Endnotes, CT_FtnEdn)."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List, cast

from docx.oxml.parser import OxmlElement
from docx.oxml.simpletypes import ST_DecimalNumber
from docx.oxml.xmlchemy import BaseOxmlElement, OneOrMore, RequiredAttribute, ZeroOrMore

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P


class CT_Endnotes(BaseOxmlElement):
    """`w:endnotes` element, the root element for the endnotes part.

    Contains a sequence of `w:endnote` elements. The first two are always the
    `separator`/`continuationSeparator` marks Word uses to render the endnote
    separator line; real endnote content follows those.
    """

    add_endnote_sequence: Callable[[], CT_FtnEdn]
    endnote_sequence_lst: List[CT_FtnEdn]

    endnote_sequence = OneOrMore("w:endnote")

    def add_endnote(self, endnote_reference_id: int) -> CT_FtnEdn:
        """Return a newly appended `w:endnote` element having `endnote_reference_id`."""
        new_e = self.add_endnote_sequence()
        new_e.id = endnote_reference_id
        return new_e

    def get_by_id(self, id: int) -> CT_FtnEdn | None:
        """The `w:endnote` element having `@w:id` equal to `id`, or |None| if not found."""
        found = self.xpath(f'w:endnote[@w:id="{id}"]')
        return found[0] if found else None


class CT_FtnEdn(BaseOxmlElement):
    """`w:endnote` element, containing the content of a single endnote.

    An endnote is a so-called "story" and can contain paragraphs and tables much like a
    table cell.
    """

    add_p: Callable[[], CT_P]
    p_lst: List[CT_P]
    tbl_lst: List[CT_Tbl]
    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    def add_endnote_before(self, endnote_reference_id: int) -> CT_FtnEdn:
        """Return a newly created `w:endnote` element inserted immediately before this
        one, having `endnote_reference_id`."""
        new_endnote = cast("CT_FtnEdn", OxmlElement("w:endnote"))
        new_endnote.id = endnote_reference_id
        self.addprevious(new_endnote)
        return new_endnote

    @property
    def inner_content_elements(self) -> List[CT_P | CT_Tbl]:
        """Generate all `w:p` and `w:tbl` elements in this endnote."""
        return self.xpath("./w:p | ./w:tbl")
