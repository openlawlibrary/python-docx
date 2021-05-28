"""
Custom element class related to tracking changes.
"""

from ..simpletypes import ST_DecimalNumber, ST_String
from ..xmlchemy import BaseOxmlElement, RequiredAttribute, ZeroOrMore

class CT_TrackChange(BaseOxmlElement):
    """
    ``<w:ins>`` and ``<w:del>`` elements with id and author attributes.
    Elements contain runs, content controls, and other insertions and deletions.
    """
    id = RequiredAttribute('w:id', ST_DecimalNumber)
    author = RequiredAttribute('w:author', ST_String)

    r = ZeroOrMore('w:r')
    sdt = ZeroOrMore('w:sdt')
    ins_ = ZeroOrMore('w:ins')
    del_ = ZeroOrMore('w:del')

    @property
    def text(self):
        text = ''
        for r in self.r_lst:
            text += r.text
        return text
