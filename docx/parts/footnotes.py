# encoding: utf-8

"""
Provides FootnotesPart and related objects
"""

import os

from ..opc.constants import CONTENT_TYPE as CT
from ..opc.packuri import PackURI
from docx.opc.part import XmlPart
from ..oxml import parse_xml
from ..footnotes import Footnotes


class FootnotesPart(XmlPart):
    """
    Proxy for the footnotes.xml part containing footnotes definitions for a document.
    """
    @classmethod
    def default(cls, package):
        """
        Return a newly created footnote part, containing a default set of elements.
        """
        partname = PackURI('/word/footnotes.xml')
        content_type = CT.WML_FOOTNOTES
        element = parse_xml(cls._default_footnote_xml())
        return cls(partname, content_type, element, package)

    @property
    def footnotes(self):
        """
        The |Footnotes| instance containing the footnotes (<w:footnotes> element
        proxies) for this footnotes part.
        """
        return Footnotes(self.element)

    @classmethod
    def _default_footnote_xml(cls):
        """
        Return a bytestream containing XML for a default styles part.
        """
        path = os.path.join(
            os.path.split(__file__)[0], '..', 'templates',
            'default-footnotes.xml'
        )
        with open(path, 'rb') as f:
            xml_bytes = f.read()
        return xml_bytes
