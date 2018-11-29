"""
Custom element classes that represent content control
elements within the document part.
"""

from .xmlchemy import BaseOxmlElement, ZeroOrOne, ZeroOrMore
from .ns import qn

class CT_SdtBase(BaseOxmlElement):
    """
    ``<w:sdt>`` structured document tag element specifying
    a content control elements.
    """
    _tag_seq = ('w:sdtPr', 'w:stEndPr', 'w:sdtContent')

    sdtPr = ZeroOrOne('w:sdtPr', successors=_tag_seq[1:])
    stEndPr = ZeroOrOne('w:stEndPr', successors=_tag_seq[2:])
    sdtContent = ZeroOrOne('w:sdtContent', successors=())

    del _tag_seq

class CT_SdtContentBase(BaseOxmlElement):
    """
    ``<w:sdtContent>`` represents content within ``<w:sdt>``
    """
    r = ZeroOrMore('w:r')

    @property
    def runs(self):
        def get_runs(el):
            for child in el:
                if child.tag == qn('w:r'):
                    return child
                else:
                    yield from get_runs(child)
        yield from get_runs(self)
