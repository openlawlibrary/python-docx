"""Custom element classes related to footnotes (CT_Footnotes, CT_FtnEnd)."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List, cast

from docx.oxml.parser import OxmlElement
from docx.oxml.simpletypes import ST_DecimalNumber
from docx.oxml.xmlchemy import BaseOxmlElement, OneOrMore, RequiredAttribute, ZeroOrMore

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P


class CT_Footnotes(BaseOxmlElement):
    """`w:footnotes` element, the root element for the footnotes part.

    Contains a sequence of `w:footnote` elements. The first two are always the
    `separator`/`continuationSeparator` marks Word uses to render the footnote
    separator line; real footnote content follows those.
    """

    add_footnote_sequence: Callable[[], CT_FtnEnd]
    footnote_sequence_lst: List[CT_FtnEnd]

    footnote_sequence = OneOrMore("w:footnote")

    def add_footnote(self, footnote_reference_id: int) -> CT_FtnEnd:
        """Return a newly appended `w:footnote` element having `footnote_reference_id`."""
        new_f = self.add_footnote_sequence()
        new_f.id = footnote_reference_id
        return new_f

    def get_by_id(self, id: int) -> CT_FtnEnd | None:
        """The `w:footnote` element having `@w:id` equal to `id`, or |None| if not found."""
        found = self.xpath(f'w:footnote[@w:id="{id}"]')
        return found[0] if found else None


class CT_FtnEnd(BaseOxmlElement):
    """`w:footnote` element, containing the content of a single footnote.

    A footnote is a so-called "story" and can contain paragraphs and tables much like a
    table cell.
    """

    add_p: Callable[[], CT_P]
    p_lst: List[CT_P]
    tbl_lst: List[CT_Tbl]
    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    def add_footnote_before(self, footnote_reference_id: int) -> CT_FtnEnd:
        """Return a newly created `w:footnote` element inserted immediately before this
        one, having `footnote_reference_id`."""
        new_footnote = cast("CT_FtnEnd", OxmlElement("w:footnote"))
        new_footnote.id = footnote_reference_id
        self.addprevious(new_footnote)
        return new_footnote

    @property
    def inner_content_elements(self) -> List[CT_P | CT_Tbl]:
        """Generate all `w:p` and `w:tbl` elements in this footnote."""
        return self.xpath("./w:p | ./w:tbl")
