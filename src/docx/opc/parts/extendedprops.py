"""Extended properties part, corresponds to `/docProps/app.xml` part in package."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.extendedprops import ExtendedProperties
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml.extendedprops import CT_ExtendedProperties

if TYPE_CHECKING:
    from docx.opc.package import OpcPackage


class ExtendedPropertiesPart(XmlPart):
    """Corresponds to part named `/docProps/app.xml`.

    Contains the extended document properties for this document package.
    """

    @classmethod
    def default(cls, package: OpcPackage) -> ExtendedPropertiesPart:
        """Return a new |ExtendedPropertiesPart| object initialized with default
        (empty) extended properties."""
        return cls._new(package)

    @property
    def extended_properties(self) -> ExtendedProperties:
        """An |ExtendedProperties| object providing read/write access to the extended
        properties contained in this extended properties part."""
        return ExtendedProperties(self.element)

    @classmethod
    def _new(cls, package: OpcPackage) -> ExtendedPropertiesPart:
        partname = PackURI("/docProps/app.xml")
        content_type = CT.OFC_EXTENDED_PROPERTIES
        extendedProperties = CT_ExtendedProperties.new()
        return ExtendedPropertiesPart(partname, content_type, extendedProperties, package)
