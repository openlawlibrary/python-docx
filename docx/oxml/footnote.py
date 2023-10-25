# encoding: utf-8

"""
Custom element classes related to footnote (CT_FtnEnd, CT_Footnotes).
"""

from .ns import qn
from .xmlchemy import (
    BaseOxmlElement, RequiredAttribute, ZeroOrMore, OneOrMore
)
from .simpletypes import (
    ST_DecimalNumber
)

class CT_Footnotes(BaseOxmlElement):
    """
    ``<w:footnotes>`` element, containing a sequence of footnote (w:footnote) elements
    """
    footnote_sequence = OneOrMore('w:footnote')

    def get_by_id(self, id):
        found = self.xpath('w:footnote[@w:id="%s"]' % id)
        if not found:
            return None
        return found[0]


class CT_FtnEnd(BaseOxmlElement):
    """
    ``<w:footnote>`` element, containing the properties for a specific footnote
    """
    id = RequiredAttribute('w:id', ST_DecimalNumber)
    p = ZeroOrMore('w:p')

    @property
    def paragraphs(self):
        """
        Returns a list of paragraphs |CT_P|, or |None| if none paragraph is present.
        """
        paragraphs = []
        for child in self:
            if child.tag == qn('w:p'):
                paragraphs.append(child)
        if paragraphs == []:
            paragraphs = None
        return paragraphs
