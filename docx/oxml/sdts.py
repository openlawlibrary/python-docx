"""
Custom element classes that represent content control
elements within the document. Since there are four
types of `sdt` and `sdtContent` elements and types
for each part of the document (block, cell, row and run).
`Base` represents single element and element type covering
all occurrences of `w:sdt` within the document.
"""

from .xmlchemy import BaseOxmlElement, ZeroOrOne, ZeroOrMore
from .ns import nsmap, qn

class CT_SdtBase(BaseOxmlElement):
    """
    ``<w:sdt>`` structured document tag element (content control)
    specifies content control elements on any level of document.
    """
    _tag_seq = ('w:sdtPr', 'w:sdtEndPr', 'w:sdtContent')

    sdtPr = ZeroOrOne('w:sdtPr', successors=_tag_seq[1:])
    sdtEndPr = ZeroOrOne('w:sdtEndPr', successors=_tag_seq[2:])
    sdtContent = ZeroOrOne('w:sdtContent', successors=())

    del _tag_seq

    @property
    def name(self):
        return self.sdtPr.name

class CT_SdtPr(BaseOxmlElement):
    """
    ``<w:sdtPr>`` represents property element of ``<w:sdt>`` (content control).
    """
    _alias = ZeroOrOne('w:alias')
    lock = ZeroOrOne('w:lock')
    placeholder = ZeroOrOne('w:placeholder')
    _temporary = ZeroOrOne('w:temporary')
    tag_name = ZeroOrOne('w:tag')
    active_placeholder = ZeroOrOne('w:showingPlcHdr')
    rPr = ZeroOrOne('w:rPr')

    @property
    def name(self):
        try:
            return self.tag_name.get('{%s}val' % nsmap['w'])
        except AttributeError:
            return None

    @name.setter
    def name(self, tag_name):
        tag = self._add_tag_name()
        tag.set('{%s}val' % nsmap['w'], tag_name)

    @property
    def alias(self):
        return self._alias.get('{%s}val' % nsmap['w'])

    @alias.setter
    def alias(self, alias_name):
        alias = self._add__alias()
        alias.set('{%s}val' % nsmap['w'], alias_name)

    @property
    def temporary(self):
        return self._temporary.get('{%s}val' % nsmap['w'])

    @temporary.setter
    def temporary(self, tmp):
        temp = self._add__temporary()
        temp.set(nsmap['w'], tmp)

class CT_SdtContentBase(BaseOxmlElement):
    """
    ``<w:sdtContent>`` represents content within ``<w:sdt>`` (content control).'
    It contains all paragraphs within the content control.
    """
    p = ZeroOrMore('w:p')
    r = ZeroOrMore('w:r')
    sdt = ZeroOrMore('w:sdt')
    tbl = ZeroOrMore('w:tbl')

    def iter_runs(self):
        def walk(el):
            for child in el:
                if child.tag == qn('w:r'):
                    yield child
                elif child.tag in (qn('w:smartTag'),):
                    yield from get_runs(child)
                else:
                    yield from walk(child)
        yield from walk(self)
