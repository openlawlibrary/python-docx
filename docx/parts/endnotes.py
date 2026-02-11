# encoding: utf-8

"""
Provides EndnotesPart and related objects
"""

import os

from ..opc.constants import CONTENT_TYPE as CT
from ..opc.packuri import PackURI
from ..oxml import parse_xml
from ..endnotes import Endnotes
from docx.parts.story import BaseStoryPart


class EndnotesPart(BaseStoryPart):
    """
    Proxy for the endnotes.xml part containing endnotes definitions for a document.
    """
    @classmethod
    def default(cls, package):
        """
        Return a newly created endnote part, containing a default set of elements.
        """
        partname = PackURI('/word/endnotes.xml')
        content_type = CT.WML_ENDNOTES
        element = parse_xml(cls._default_endnote_xml())
        return cls(partname, content_type, element, package)

    @property
    def endnotes(self):
        """
        The |Endnotes| instance containing the endnotes (<w:endnotes> element
        proxies) for this endnotes part.
        """
        return Endnotes(self.element, self)

    @classmethod
    def _default_endnote_xml(cls):
        """
        Return a bytestream containing XML for a default endnotes part.
        """
        path = os.path.join(
            os.path.split(__file__)[0], '..', 'templates',
            'default-endnotes.xml'
        )
        with open(path, 'rb') as f:
            xml_bytes = f.read()
        return xml_bytes
