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
    _tag_seq = ('w:sdtPr', 'w:stEndPr', 'w:sdtContent')

    sdtPr = ZeroOrOne('w:sdtPr', successors=_tag_seq[1:])
    stEndPr = ZeroOrOne('w:stEndPr', successors=_tag_seq[2:])
    sdtContent = ZeroOrOne('w:sdtContent', successors=())

    del _tag_seq

    @property
    def name(self):
        return self.sdtPr.name

class CT_SdtPr(BaseOxmlElement):
    """
    ``<w:sdtPr>`` represents property element of ``<w:sdt>`` (content control).
    """
    tag = ZeroOrOne('w:tag')
    date = ZeroOrOne('w:date')
    active_placeholder = ZeroOrOne('w:showingPlcHdr')

    @property
    def name(self):
        try:
            return self.tag.get('{%s}val' % nsmap['w'])
        except AttributeError:
            return None


class CT_SdtContentBase(BaseOxmlElement):
    """
    ``<w:sdtContent>`` represents content within ``<w:sdt>`` (content control).'
    It contains all paragraphs within the content control.
    """
    p = ZeroOrMore('w:p')

    def iter_runs(self):
        def walk(el):
            for child in el:
                if child.tag == qn('w:r'):
                    yield child
                else:
                    yield from walk(child)
        yield from walk(self)
