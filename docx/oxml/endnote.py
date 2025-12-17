# encoding: utf-8

"""
Custom element classes related to endnote (CT_FtnEdn, CT_Endnotes).
"""

from .ns import qn
from .xmlchemy import (
    BaseOxmlElement, OxmlElement, RequiredAttribute, ZeroOrMore, OneOrMore
)
from .simpletypes import (
    ST_DecimalNumber
)

class CT_Endnotes(BaseOxmlElement):
    """
    ``<w:endnotes>`` element, containing a sequence of endnote (w:endnote) elements
    """
    endnote_sequence = OneOrMore('w:endnote')

    def add_endnote(self, endnote_reference_id):
        """
        Create a ``<w:endnote>`` element with `endnote_reference_id`.
        """
        new_e = self.add_endnote_sequence()
        new_e.id = endnote_reference_id
        return new_e

    def get_by_id(self, id):
        found = self.xpath(f'w:endnote[@w:id="{id}"]')
        if not found:
            return None
        return found[0]


class CT_FtnEdn(BaseOxmlElement):
    """
    ``<w:endnote>`` element, containing the properties for a specific endnote.
    Note: CT_FtnEdn is the correct OOXML type name for both footnotes and endnotes.
    """
    id = RequiredAttribute('w:id', ST_DecimalNumber)
    p = ZeroOrMore('w:p')

    def add_endnote_before(self, endnote_reference_id):
        """
        Create a ``<w:endnote>`` element with `endnote_reference_id`
        and insert it before the current element.
        """
        new_endnote = OxmlElement('w:endnote')
        new_endnote.id = endnote_reference_id
        self.addprevious(new_endnote)
        return new_endnote
