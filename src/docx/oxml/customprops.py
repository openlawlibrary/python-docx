"""Custom element class for the custom-properties part root element."""

from __future__ import annotations

from typing import cast

from docx.oxml.ns import nsdecls
from docx.oxml.parser import parse_xml
from docx.oxml.xmlchemy import BaseOxmlElement


class CT_CustomProperties(BaseOxmlElement):
    """`<Properties>` element, the root element of the Custom Properties part.

    Stored as `/docProps/custom.xml`. Each child `<property>` element holds one custom
    document property.
    """

    _customProperties_tmpl = (
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006'
        '/custom-properties" %s/>\n' % nsdecls("vt")
    )

    @classmethod
    def new(cls) -> CT_CustomProperties:
        """Return a new `<Properties>` element."""
        return cast(CT_CustomProperties, parse_xml(cls._customProperties_tmpl))
