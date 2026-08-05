"""Custom element classes for structured document tags (`w:sdt`, content controls)."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator, List

from docx.oxml.ns import qn
from docx.oxml.shared import CT_OnOff, CT_String
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore, ZeroOrOne

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.font import CT_RPr
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.text.run import CT_R


class CT_SdtBase(BaseOxmlElement):
    """`<w:sdt>` element, a structured document tag (content control).

    Can appear at inline (run) or block (paragraph/table) level; both use this same
    element type in this simplified content-control model.
    """

    get_or_add_sdtPr: Callable[[], CT_SdtPr]
    get_or_add_sdtContent: Callable[[], CT_SdtContentBase]

    sdtPr: CT_SdtPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:sdtPr", successors=("w:sdtContent",)
    )
    sdtContent: CT_SdtContentBase | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:sdtContent"
    )

    @property
    def name(self) -> str | None:
        """Tag name of this content control.

        |None| if `w:sdtPr` is not present, which does not occur for a `w:sdt` created
        via `add_sdt()`.
        """
        sdtPr = self.sdtPr
        return None if sdtPr is None else sdtPr.name

    @property
    def text(self) -> str:  # pyright: ignore[reportIncompatibleMethodOverride]
        """The textual content of this content control.

        `CT_SdtBase` stores its text in zero-or-more `w:r` descendants, at any depth,
        of its `w:sdtContent` child. This allows the text of a content control to
        participate in the text of the paragraph or run that contains it.
        """
        sdtContent = self.sdtContent
        return "" if sdtContent is None else "".join(r.text for r in sdtContent.iter_runs())


class CT_SdtPr(BaseOxmlElement):
    """`<w:sdtPr>` element, the properties of a content control."""

    get_or_add_rPr: Callable[[], CT_RPr]
    get_or_add_alias: Callable[[], CT_String]
    get_or_add_tagElm: Callable[[], CT_String]
    get_or_add_showingPlcHdr: Callable[[], CT_OnOff]
    _remove_showingPlcHdr: Callable[[], None]

    rPr: CT_RPr | None = ZeroOrOne("w:rPr")  # pyright: ignore[reportAssignmentType]
    alias: CT_String | None = ZeroOrOne("w:alias")  # pyright: ignore[reportAssignmentType]
    # -- named `tagElm`, not `tag`, to avoid colliding with `_Element.tag` --
    tagElm: CT_String | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:tag"
    )
    showingPlcHdr: CT_OnOff | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:showingPlcHdr"
    )

    @property
    def name(self) -> str | None:
        """Tag name of this content control, `./w:tag/@w:val`.

        |None| if `w:tag` is not present.
        """
        tagElm = self.tagElm
        return None if tagElm is None else tagElm.val

    @name.setter
    def name(self, value: str) -> None:
        self.get_or_add_tagElm().val = value

    @property
    def alias_val(self) -> str | None:
        """`./w:alias/@w:val`, or |None| if `w:alias` is not present."""
        alias = self.alias
        return None if alias is None else alias.val

    @alias_val.setter
    def alias_val(self, value: str) -> None:
        self.get_or_add_alias().val = value

    @property
    def active_placeholder(self) -> bool:
        """True when this content control is showing its placeholder text."""
        return self.showingPlcHdr is not None

    @active_placeholder.setter
    def active_placeholder(self, value: bool) -> None:
        if value:
            self.get_or_add_showingPlcHdr().val = True
        else:
            self._remove_showingPlcHdr()


class CT_SdtContentBase(BaseOxmlElement):
    """`<w:sdtContent>` element, the content of a content control."""

    add_p: Callable[[], CT_P]
    _add_r: Callable[[], CT_R]
    _new_sdt: Callable[[], CT_SdtBase]
    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    p = ZeroOrMore("w:p")
    p_lst: List[CT_P]
    r = ZeroOrMore("w:r")
    sdt = ZeroOrMore("w:sdt")
    sdt_lst: List[CT_SdtBase]
    tbl = ZeroOrMore("w:tbl")
    tbl_lst: List[CT_Tbl]

    @property
    def inner_content_elements(self) -> List[CT_P | CT_Tbl]:
        """Generate all `w:p` and `w:tbl` elements directly in this content.

        Elements appear in document order.
        """
        return self.xpath("./w:p | ./w:tbl")

    def iter_runs(self) -> Iterator[CT_R]:
        """Generate each `w:r` descendant of this element, in document order."""
        return self.iterdescendants(qn("w:r"))  # pyright: ignore[reportReturnType]
