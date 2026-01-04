# encoding: utf-8

"""
Extended properties part, corresponds to ``/docProps/app.xml`` part in package.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..constants import CONTENT_TYPE as CT
from ..extendedprops import ExtendedProperties
from ...oxml.extendedprops import CT_ExtendedProperties
from ..packuri import PackURI
from ..part import XmlPart


class ExtendedPropertiesPart(XmlPart):
    """
    Corresponds to part named ``/docProps/app.xml``, containing the extended
    document properties for this document package.
    """
    @classmethod
    def default(cls, package):
        """
        Return a new |ExtendedPropertiesPart| object initialized with default
        values for its base properties.
        """
        extended_properties_part = cls._new(package)
        extended_properties = extended_properties_part.extended_properties
        extended_properties.template = 'Normal.dotm'
        extended_properties.total_time = 0
        extended_properties.application = 'python-docx'
        return extended_properties_part

    @property
    def extended_properties(self):
        """
        An |ExtendedProperties| object providing read/write access to the
        extended properties contained in this extended properties part.
        """
        return ExtendedProperties(self.element)

    @classmethod
    def _new(cls, package):
        partname = PackURI('/docProps/app.xml')
        content_type = CT.OFC_EXTENDED_PROPERTIES
        extendedProperties = CT_ExtendedProperties.new()
        return ExtendedPropertiesPart(
            partname, content_type, extendedProperties, package
        )
