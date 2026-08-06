# pyright: reportPrivateUsage=false

"""Custom element classes related to paragraphs (CT_P)."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Callable, List, cast

from docx.oxml.parser import OxmlElement
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore, ZeroOrOne

if TYPE_CHECKING:
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml.bookmark import CT_BookmarkEnd, CT_BookmarkStart
    from docx.oxml.numbering import CT_Lvl, CT_Numbering
    from docx.oxml.sdts import CT_SdtBase
    from docx.oxml.section import CT_SectPr
    from docx.oxml.text.hyperlink import CT_Hyperlink
    from docx.oxml.text.pagebreak import CT_LastRenderedPageBreak
    from docx.oxml.text.parfmt import CT_PPr
    from docx.oxml.text.run import CT_R


class CT_P(BaseOxmlElement):
    """`<w:p>` element, containing the properties and text for a paragraph."""

    add_r: Callable[[], CT_R]
    get_or_add_pPr: Callable[[], CT_PPr]
    bookmarkStart_lst: List[CT_BookmarkStart]
    bookmarkEnd_lst: List[CT_BookmarkEnd]
    hyperlink_lst: List[CT_Hyperlink]
    r_lst: List[CT_R]
    sdt_lst: List[CT_SdtBase]
    _new_sdt: Callable[[], CT_SdtBase]

    bookmarkStart = ZeroOrMore("w:bookmarkStart", successors=("w:pPr", "w:hyperlink", "w:r"))
    pPr: CT_PPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:pPr", successors=("w:bookmarkEnd",)
    )
    hyperlink = ZeroOrMore("w:hyperlink", successors=("w:bookmarkEnd",))
    r = ZeroOrMore("w:r", successors=("w:bookmarkEnd",))
    sdt = ZeroOrMore("w:sdt", successors=("w:bookmarkEnd",))
    bookmarkEnd = ZeroOrMore("w:bookmarkEnd")

    def add_p_before(self) -> CT_P:
        """Return a new `<w:p>` element inserted directly prior to this one."""
        new_p = cast(CT_P, OxmlElement("w:p"))
        self.addprevious(new_p)
        return new_p

    @property
    def alignment(self) -> WD_PARAGRAPH_ALIGNMENT | None:
        """The value of the `<w:jc>` grandchild element or |None| if not present."""
        pPr = self.pPr
        if pPr is None:
            return None
        return pPr.jc_val

    @alignment.setter
    def alignment(self, value: WD_PARAGRAPH_ALIGNMENT):
        pPr = self.get_or_add_pPr()
        pPr.jc_val = value

    def clear_content(self):
        """Remove all child elements, except the `<w:pPr>` element if present."""
        for child in self.xpath("./*[not(self::w:pPr)]"):
            self.remove(child)

    @property
    def endnote_reference_ids(self) -> List[int]:
        """The `@w:id` of each `w:endnoteReference` in this paragraph's runs."""
        return [id for r in self.r_lst for id in r.endnote_reference_ids]

    @property
    def footnote_reference_ids(self) -> List[int]:
        """The `@w:id` of each `w:footnoteReference` in this paragraph's runs."""
        return [id for r in self.r_lst for id in r.footnote_reference_ids]

    @property
    def inner_content_elements(self) -> List[CT_R | CT_Hyperlink]:
        """Run and hyperlink children of the `w:p` element, in document order."""
        return self.xpath("./w:r | ./w:hyperlink")

    @property
    def lastRenderedPageBreaks(self) -> List[CT_LastRenderedPageBreak]:
        """All `w:lastRenderedPageBreak` descendants of this paragraph.

        Rendered page-breaks commonly occur in a run but can also occur in a run inside
        a hyperlink. This returns both.
        """
        return self.xpath(
            "./w:r/w:lastRenderedPageBreak | ./w:hyperlink/w:r/w:lastRenderedPageBreak"
        )

    def lvl_from_para_props(self, numbering_el: CT_Numbering) -> CT_Lvl | None:
        """The `w:lvl` numbering-level definition for this paragraph, resolved via its
        own direct numbering properties (not those inherited from its style)."""
        return numbering_el.get_lvl_from_props(self)

    def lvl_from_style_props(
        self, numbering_el: CT_Numbering, styles_cache: dict[str, Any]
    ) -> CT_Lvl | None:
        """The `w:lvl` numbering-level definition for this paragraph, resolved via its
        paragraph style's numbering properties."""
        return numbering_el.get_lvl_from_props(self, styles_cache)

    def number(self, numbering_el: CT_Numbering, styles_cache: dict[str, Any]) -> str | None:
        """This paragraph's list-item label (e.g. `"1)\\t"`), or |None| if this
        paragraph is not part of a numbered list."""
        return numbering_el.get_num_for_p(self, styles_cache)

    def set_li_lvl(
        self,
        numbering_el: CT_Numbering,
        styles_cache: dict[str, Any],
        prev_el: CT_P | None,
        ilvl: int | None,
    ) -> None:
        """Set this paragraph's list-item indentation level."""
        numbering_el.set_li_lvl(self, styles_cache, prev_el, ilvl)

    def set_sectPr(self, sectPr: CT_SectPr):
        """Unconditionally replace or add `sectPr` as grandchild in correct sequence."""
        pPr = self.get_or_add_pPr()
        pPr._remove_sectPr()
        pPr._insert_sectPr(sectPr)

    @property
    def style(self) -> str | None:
        """String contained in `w:val` attribute of `./w:pPr/w:pStyle` grandchild.

        |None| if not present.
        """
        pPr = self.pPr
        if pPr is None:
            return None
        return pPr.style

    @style.setter
    def style(self, style: str | None):
        pPr = self.get_or_add_pPr()
        pPr.style = style

    @property
    def text(self):  # pyright: ignore[reportIncompatibleMethodOverride]
        """The textual content of this paragraph.

        Inner-content child elements like `w:r` and `w:hyperlink` are translated to
        their text equivalent.
        """
        return "".join(e.text for e in self.xpath("w:r | w:hyperlink | w:sdt"))

    def _insert_pPr(self, pPr: CT_PPr) -> CT_PPr:
        self.insert(0, pPr)
        return pPr
