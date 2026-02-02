# encoding: utf-8

"""
Unit test suite for the docx.opc.parts.extendedprops module
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from docx.opc.extendedprops import ExtendedProperties
from docx.opc.parts.extendedprops import ExtendedPropertiesPart
from docx.oxml.extendedprops import CT_ExtendedProperties

from ...unitutil.mock import class_mock, instance_mock


class DescribeExtendedPropertiesPart(object):

    def it_provides_access_to_its_extended_props_object(self, extprops_fixture):
        extended_properties_part, ExtendedProperties_ = extprops_fixture
        extended_properties = extended_properties_part.extended_properties
        ExtendedProperties_.assert_called_once_with(extended_properties_part.element)
        assert isinstance(extended_properties, ExtendedProperties)

    def it_can_create_a_default_extended_properties_part(self):
        extended_properties_part = ExtendedPropertiesPart.default(None)
        assert isinstance(extended_properties_part, ExtendedPropertiesPart)
        extended_properties = extended_properties_part.extended_properties
        assert extended_properties.template == 'Normal.dotm'
        assert extended_properties.total_time == 0
        assert extended_properties.application == 'python-docx'

    # fixtures ---------------------------------------------

    @pytest.fixture
    def extprops_fixture(self, element_, ExtendedProperties_):
        extended_properties_part = ExtendedPropertiesPart(None, None, element_, None)
        return extended_properties_part, ExtendedProperties_

    # fixture components -----------------------------------

    @pytest.fixture
    def ExtendedProperties_(self, request):
        return class_mock(request, 'docx.opc.parts.extendedprops.ExtendedProperties')

    @pytest.fixture
    def element_(self, request):
        return instance_mock(request, CT_ExtendedProperties)
