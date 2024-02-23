# encoding: utf-8

"""
Custom element classes related to footnote (CT_FtnEnd, CT_Footnotes).
"""

from .ns import qn
from .xmlchemy import (
    BaseOxmlElement, OxmlElement, RequiredAttribute, ZeroOrMore, OneOrMore
)
from .simpletypes import (
    ST_DecimalNumber
)

class CT_Footnotes(BaseOxmlElement):
    """
    ``<w:footnotes>`` element, containing a sequence of footnote (w:footnote) elements
    """
    footnote_sequence = OneOrMore('w:footnote')

    def add_footnote(self, footnote_reference_id):
        """
        Create a ``<w:footnote>`` element with `footnote_reference_id`.
        """
        new_f = self.add_footnote_sequence()
        new_f.id = footnote_reference_id
        return new_f

    def get_by_id(self, id):
        found = self.xpath(f'w:footnote[@w:id="{id}"]')
        if not found:
            return None
        return found[0]


class CT_FtnEnd(BaseOxmlElement):
    """
    ``<w:footnote>`` element, containing the properties for a specific footnote
    """
    id = RequiredAttribute('w:id', ST_DecimalNumber)
    p = ZeroOrMore('w:p')

    def add_footnote_before(self, footnote_reference_id):
        """
        Create a ``<w:footnote>`` element with `footnote_reference_id`
        and insert it before the current element.
        """
        new_footnote = OxmlElement('w:footnote')
        new_footnote.id = footnote_reference_id
        self.addprevious(new_footnote)
        return new_footnote
