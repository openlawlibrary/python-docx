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
    alias = ZeroOrOne('w:alias')
    lock = ZeroOrOne('w:lock')
    placeholder = ZeroOrOne('w:placeholder')
    temporary = ZeroOrOne('w:temporary')
    tag_name = ZeroOrOne('w:tag')
    active_placeholder = ZeroOrOne('w:showingPlcHdr')
    rPr = ZeroOrOne('w:rPr')

    @property
    def name(self):
        try:
            return self.tag_name.get('{%s}val' % nsmap['w'])
        except AttributeError:
            return None


class CT_SdtContentBase(BaseOxmlElement):
    """
    ``<w:sdtContent>`` represents content within ``<w:sdt>`` (content control).'
    It contains all paragraphs within the content control.
    """
    p = ZeroOrMore('w:p')
    r = ZeroOrMore('w:r')

    def iter_runs(self):
        def walk(el):
            for child in el:
                if child.tag == qn('w:r'):
                    yield child
                else:
                    yield from walk(child)
        yield from walk(self)
