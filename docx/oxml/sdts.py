"""
Custom element classes that represent content control
elements within the document part.
"""

from .xmlchemy import BaseOxmlElement, ZeroOrOne, ZeroOrMore
from .ns import nsmap

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

    @property
    def name(self):
        return self.sdtPr.name

class CT_SdtPr(BaseOxmlElement):
    tag = ZeroOrOne('w:tag')
    date = ZeroOrOne('w:date')

    @property
    def name(self):
        try:
            return self.tag.get('{%s}val' % nsmap['w'])
        except:
            raise Exception('All content controls should be named (having '
            'set unique content control tag name).')


class CT_SdtContentBase(BaseOxmlElement):
    """
    ``<w:sdtContent>`` represents content within ``<w:sdt>``
    """
    p = ZeroOrMore('w:p')
