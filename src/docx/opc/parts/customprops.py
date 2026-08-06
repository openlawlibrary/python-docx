"""Custom properties part, corresponds to `/docProps/custom.xml` part in package."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.customprops import CustomProperties
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml.customprops import CT_CustomProperties

if TYPE_CHECKING:
    from docx.opc.package import OpcPackage


class CustomPropertiesPart(XmlPart):
    """Corresponds to part named `/docProps/custom.xml`.

    Contains the custom document properties for this document package.
    """

    @classmethod
    def default(cls, package: OpcPackage) -> CustomPropertiesPart:
        """Return a new |CustomPropertiesPart| object initialized with default
        (empty) custom properties."""
        return cls._new(package)

    @property
    def custom_properties(self) -> CustomProperties:
        """A |CustomProperties| object providing read/write access to the custom
        properties contained in this custom properties part."""
        return CustomProperties(self.element)

    @classmethod
    def _new(cls, package: OpcPackage) -> CustomPropertiesPart:
        partname = PackURI("/docProps/custom.xml")
        content_type = CT.OFC_CUSTOM_PROPERTIES
        customProperties = CT_CustomProperties.new()
        return CustomPropertiesPart(partname, content_type, customProperties, package)
