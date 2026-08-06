# pyright: reportPrivateUsage=false

"""Custom element classes related to paragraphs (CT_P)."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List, cast

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore, ZeroOrOne

_FLD_CHAR_CONTAINER_TAGS = (qn("w:hyperlink"), qn("w:sdt"), qn("w:sdtContent"), qn("w:smartTag"))

if TYPE_CHECKING:
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml.section import CT_SectPr
    from docx.oxml.text.hyperlink import CT_Hyperlink
    from docx.oxml.text.pagebreak import CT_LastRenderedPageBreak
    from docx.oxml.text.parfmt import CT_PPr
    from docx.oxml.text.run import CT_R


class CT_P(BaseOxmlElement):
    """`<w:p>` element, containing the properties and text for a paragraph."""

    add_r: Callable[[], CT_R]
    get_or_add_pPr: Callable[[], CT_PPr]
    hyperlink_lst: List[CT_Hyperlink]
    r_lst: List[CT_R]

    pPr: CT_PPr | None = ZeroOrOne("w:pPr")  # pyright: ignore[reportAssignmentType]
    hyperlink = ZeroOrMore("w:hyperlink")
    r = ZeroOrMore("w:r")

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
    def inner_content_elements(self) -> List[CT_R | CT_Hyperlink]:
        """Run and hyperlink children of the `w:p` element, in document order."""
        self.strip_hidden_fld_char_content()
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

    def set_sectPr(self, sectPr: CT_SectPr):
        """Unconditionally replace or add `sectPr` as grandchild in correct sequence."""
        pPr = self.get_or_add_pPr()
        pPr._remove_sectPr()
        pPr._insert_sectPr(sectPr)

    def strip_hidden_fld_char_content(self, container: BaseOxmlElement | None = None) -> None:
        """Remove markup Word hides inside a complex-field (`w:fldChar`) sequence.

        A complex field -- used for a table of contents, cross-reference, computed
        date, and the like -- is represented as a run sequence: a `begin` marker, the
        field-code runs (an `instrText` element, not meant to be visible), optionally a
        `separate` marker followed by the cached field *result* (the only part meant to
        be visible), and an `end` marker. This walks this paragraph's own runs, and
        those nested in each of its hyperlinks/content-controls, removing every run
        child that falls inside a hidden span along with the now-redundant `w:fldChar`
        markers themselves, dropping any run left empty as a result.

        This mutates the paragraph tree permanently and is idempotent -- once run, no
        `w:fldChar` markers remain to process on a later call. This matches Word's own
        model of a field's cached "result" as the authoritative value on open, but it
        does mean a round-tripped save will no longer contain the original field code.
        """
        container = self if container is None else container
        ignore_depth = 0
        has_separate = False
        for child in list(container):
            if child.tag == qn("w:r"):
                mutated = False
                for grandchild in list(child):
                    if ignore_depth != 0:
                        child.remove(grandchild)
                        mutated = True
                    if grandchild.tag == qn("w:fldChar"):
                        mutated = True
                        fldCharType = grandchild.get(qn("w:fldCharType"))
                        if fldCharType == "begin":
                            ignore_depth += 1
                        elif fldCharType == "separate":
                            ignore_depth -= 1
                            has_separate = True
                        elif not has_separate:
                            ignore_depth -= 1
                        else:
                            has_separate = False
                for fldChar in child.findall(qn("w:fldChar")):
                    child.remove(fldChar)
                # -- only drop a run left empty BY this scrub -- a run that was
                # -- already empty (no fldChar involvement) is left as-is.
                if mutated and len(child) == 0:
                    container.remove(child)
            elif child.tag in _FLD_CHAR_CONTAINER_TAGS:
                self.strip_hidden_fld_char_content(cast(BaseOxmlElement, child))

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
        self.strip_hidden_fld_char_content()
        return "".join(e.text for e in self.xpath("w:r | w:hyperlink"))

    def _insert_pPr(self, pPr: CT_PPr) -> CT_PPr:
        self.insert(0, pPr)
        return pPr
