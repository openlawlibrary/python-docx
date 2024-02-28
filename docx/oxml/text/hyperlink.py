# encoding: utf-8

"""
Custom element classes related to text runs (CT_Hyperlink).
"""


from ..simpletypes import ST_String
from ..xmlchemy import (
    BaseOxmlElement, OptionalAttribute, ZeroOrMore
)

class CT_Hyperlink(BaseOxmlElement):
    """
    ``<w:hyperlink>`` element, containing properties related to field.
    """
    anchor = OptionalAttribute('w:anchor', ST_String)
    relationship_id = OptionalAttribute('r:id', ST_String)
    r = ZeroOrMore('w:r')
