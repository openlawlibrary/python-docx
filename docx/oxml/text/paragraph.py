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
    sdt = ZeroOrMore('w:sdt', CT_SdtBase)
    bookmarkStart = ZeroOrMore('w:bookmarkStart', successors=('w:pPr', 'w:r',))
    pPr = ZeroOrOne('w:pPr', successors=('w:bookmarkEnd',))
    r = ZeroOrMore('w:r', successors=('w:bookmarkEnd',))
    bookmarkEnd = ZeroOrMore('w:bookmarkEnd')
    hyperlink = ZeroOrMore('w:hyperlink')

    def _insert_pPr(self, pPr):
        self.insert(0, pPr)
        return pPr

    def add_hyperlink(self, reference, text):
        """
        Return a new ``<w:hyperlink>`` element inserted at the end of this paragraph.
        The `reference` can be a valid URL or an bookmark name.
        """
        new_h = self._add_hyperlink()
        r = new_h._add_r()
        r.text = text
        r.style = 'Hyperlink'
        if reference.startswith('rId'):
            # reference is an bookmark name, stored with an relationship
            new_h.relationship_id = reference
        else:
            new_h.anchor = reference
        return new_h

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

    @property
    def footnote_reference_ids(self):
        """
        Return all footnote reference ids (``<w:footnoteReference>``) form the paragraph,
        or |None| if not present.
        """
        footnote_ids = []
        for run in self.r_lst:
            new_footnote_ids = run.footnote_reference_ids
            if new_footnote_ids:
                footnote_ids.extend(new_footnote_ids)
        if footnote_ids == []:
            footnote_ids = None
        return footnote_ids

    def lvl_from_para_props(self, numbering_el):
        """
        Returns ``<w:lvl>`` numbering level paragraph formatting for the current paragraph using
        numbering linked via the direct paragraph formatting.
        """
        return numbering_el.get_lvl_from_props(self)

    def lvl_from_style_props(self, numbering_el, styles_cache):
        """
        Returns ``<w:lvl>`` numbering level paragraph formatting for the current paragraph using
        numbering linked via the paragraph style formatting.
        """
        return numbering_el.get_lvl_from_props(self, styles_cache)

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

    def iter_r_lst_recursive(self, return_hyperlinks=False):
        """
        Override xmlchemy generated list of runs to include runs from
        hyperlinks and content controls.
        If the argument `return_hyperlinks` is `True` then the hyperlinks
        will be returned as `CT_Hyperlink`.
        """

        def get_runs(elem):
            # Two flags used to remove hidden parts of complex field characters.
            ignoreRun = 0 # used to count nesting of ``<w:fldChar>``, if it's 0 then the run property is not inside a hidden part of ``<w:fldChar>``
            hasSeparate = False
            for child in elem:
                if child.tag == qn('w:r'):
                    # Remove run's elements that are not visible.
                    # Complex field characters have not visible part and they are removed.
                    # The code below removes hidden run elements and the ``<w:fldChar>`` tags as well.
                    for child_r in child:
                        if ignoreRun != 0:
                            child.remove(child_r)
                        if child_r.tag == qn('w:fldChar'):
                            t = child_r.fldCharType
                            if t == 'begin':
                                ignoreRun += 1
                            elif t == 'separate':
                                ignoreRun -= 1
                                # should ignore the next `end`, because text
                                # is shown from `separate` to `end`
                                hasSeparate = True
                                # We know this `fldCharType == 'end'`, so we check
                                # if this `fldChar` has 'separate' in it.
                            elif hasSeparate is False:
                                ignoreRun -= 1
                            else:
                                hasSeparate = False
                    # removes ``<w:fldChar>`` form run because the hidden elements are removed and this tag is obsolete
                    for child_r in child:
                        if child_r.tag == qn('w:fldChar'):
                            child.remove(child_r)
                    # yields runs that have at least one visible element
                    if len(child) > 0:
                        yield child
                elif child.tag == qn('w:hyperlink'):
                    if return_hyperlinks:
                        yield child
                    else:
                        yield from get_runs(child)
                elif child.tag in (qn('w:sdt'), qn('w:sdtContent'), qn('w:smartTag'),):
                    yield from get_runs(child)
        yield from get_runs(self)
