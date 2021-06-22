# encoding: utf-8

"""
Custom element classes related to paragraphs (CT_P).
"""
from ..ns import qn
from ..sdts import CT_SdtBase
from ..xmlchemy import BaseOxmlElement, OxmlElement, ZeroOrMore, ZeroOrOne


class CT_P(BaseOxmlElement):
    """
    ``<w:p>`` element, containing the properties, text for a paragraph and content controls.
    """
    pPr = ZeroOrOne('w:pPr')
    r = ZeroOrMore('w:r')
    sdt = ZeroOrMore('w:sdt', CT_SdtBase)
    bookmarkStart = ZeroOrMore('w:bookmarkStart')
    bookmarkEnd = ZeroOrMore('w:bookmarkEnd')

    def _insert_pPr(self, pPr):
        self.insert(0, pPr)
        return pPr

    def add_p_before(self):
        """
        Return a new ``<w:p>`` element inserted directly prior to this one.
        """
        new_p = OxmlElement('w:p')
        self.addprevious(new_p)
        return new_p

    @property
    def alignment(self):
        """
        The value of the ``<w:jc>`` grandchild element or |None| if not
        present.
        """
        pPr = self.pPr
        if pPr is None:
            return None
        return pPr.jc_val

    @alignment.setter
    def alignment(self, value):
        pPr = self.get_or_add_pPr()
        pPr.jc_val = value

    def clear_content(self):
        """
        Remove all child elements, except the ``<w:pPr>`` element if present.
        """
        for child in self[:]:
            if child.tag == qn('w:pPr'):
                continue
            self.remove(child)

    def lvl(self, numbering_el, styles_cache):
        """
        Returns ``<w:lvl>`` element formatting for the current paragraph.
        """
        return numbering_el.get_lvl_for_p(self, styles_cache)

    def lvl_from_props(self, numbering_el):
        """
        Returns ``<w:lvl>`` element formatting for the current paragraph.
        """
        return numbering_el.get_lvl_from_properties(self)

    def lvl_from_style_props(self, numbering_el, styles_cache):
        """
        Returns ``<w:lvl>`` element formatting for the current paragraph.
        """
        return numbering_el.get_lvl_from_style_properties(self, styles_cache)

    def number(self, numbering_el, styles_cache):
        """
        Returns numbering part of the paragraph if any, else returns None.
        """
        return numbering_el.get_num_for_p(self, styles_cache)

    def set_sectPr(self, sectPr):
        """
        Unconditionally replace or add *sectPr* as a grandchild in the
        correct sequence.
        """
        pPr = self.get_or_add_pPr()
        pPr._remove_sectPr()
        pPr._insert_sectPr(sectPr)

    def set_li_lvl(self, numbering_el, styles_el, prev_el, ilvl):
        """
        Sets list indentation level for this paragraph.
        """
        numbering_el.set_li_lvl(self, styles_el, prev_el, ilvl)

    @property
    def style(self):
        """
        String contained in w:val attribute of ./w:pPr/w:pStyle grandchild,
        or |None| if not present.
        """
        pPr = self.pPr
        if pPr is None:
            return None
        return pPr.style

    @style.setter
    def style(self, style):
        pPr = self.get_or_add_pPr()
        pPr.style = style

    def iter_r_lst_recursive(self):
        """
        Override xmlchemy generated list of runs to include runs from
        hyperlinks and content controls.
        """

        def get_runs(elem):
            for child in elem:
                if child.tag == qn('w:r'):
                    yield child
                elif child.tag in (qn('w:hyperlink'), qn('w:sdt'), qn('w:sdtContent')):
                    yield from get_runs(child)
        yield from get_runs(self)
