"""Custom element classes related to footnote (CT_FtnEnd, CT_Footnotes)."""

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.xmlchemy import (
    BaseOxmlElement, RequiredAttribute, ZeroOrMore, OneOrMore
)
from docx.oxml.simpletypes import (
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

    @property
    def paragraphs(self):
        """
        Returns a list of paragraphs |CT_P|, or |None| if none paragraph is present.
        """
        paragraphs = []
        for child in self:
            if child.tag == qn('w:p'):
                paragraphs.append(child)
        return paragraphs
