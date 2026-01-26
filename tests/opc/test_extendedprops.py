# encoding: utf-8

"""
Unit test suite for the docx.opc.extendedprops module
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from docx.opc.extendedprops import ExtendedProperties
from docx.oxml import parse_xml


class DescribeExtendedProperties(object):

    def it_knows_the_integer_property_values(self, int_prop_get_fixture):
        extended_properties, prop_name, expected_value = int_prop_get_fixture
        actual_value = getattr(extended_properties, prop_name)
        assert actual_value == expected_value

    def it_can_change_the_integer_property_values(self, int_prop_set_fixture):
        extended_properties, prop_name, value, expected_xml = int_prop_set_fixture
        setattr(extended_properties, prop_name, value)
        assert extended_properties._element.xml == expected_xml

    def it_knows_the_string_property_values(self, text_prop_get_fixture):
        extended_properties, prop_name, expected_value = text_prop_get_fixture
        actual_value = getattr(extended_properties, prop_name)
        assert actual_value == expected_value

    def it_can_change_the_string_property_values(self, text_prop_set_fixture):
        extended_properties, prop_name, value, expected_xml = text_prop_set_fixture
        setattr(extended_properties, prop_name, value)
        assert extended_properties._element.xml == expected_xml

    def it_raises_on_invalid_total_time(self, extended_properties_empty):
        with pytest.raises(ValueError):
            extended_properties_empty.total_time = -1
        with pytest.raises(ValueError):
            extended_properties_empty.total_time = 'not an int'

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('total_time', 120),
        ('pages', 5),
        ('words', 1000),
        ('characters', 5000),
        ('lines', 50),
        ('paragraphs', 10),
    ])
    def int_prop_get_fixture(self, request, extended_properties):
        prop_name, expected_value = request.param
        return extended_properties, prop_name, expected_value

    @pytest.fixture(params=[
        ('total_time', 'TotalTime', 60),
        ('pages', 'Pages', 10),
        ('words', 'Words', 500),
    ])
    def int_prop_set_fixture(self, request):
        prop_name, tagname, value = request.param
        properties = self.properties(None, None)
        extended_properties = ExtendedProperties(parse_xml(properties))
        expected_xml = self.properties(tagname, str(value))
        return extended_properties, prop_name, value, expected_xml

    @pytest.fixture(params=[
        ('template', 'Normal.dotm'),
        ('manager', 'John Doe'),
        ('company', 'Acme Corp'),
        ('application', 'Microsoft Word'),
        ('app_version', '16.0'),
    ])
    def text_prop_get_fixture(self, request, extended_properties):
        prop_name, expected_value = request.param
        return extended_properties, prop_name, expected_value

    @pytest.fixture(params=[
        ('template', 'Template', 'Custom.dotm'),
        ('manager', 'Manager', 'Jane Smith'),
        ('company', 'Company', 'New Corp'),
        ('application', 'Application', 'python-docx'),
        ('app_version', 'AppVersion', '1.0.0'),
    ])
    def text_prop_set_fixture(self, request):
        prop_name, tagname, value = request.param
        properties = self.properties(None, None)
        extended_properties = ExtendedProperties(parse_xml(properties))
        expected_xml = self.properties(tagname, value)
        return extended_properties, prop_name, value, expected_xml

    # fixture components ---------------------------------------------

    def properties(self, tagname, str_val):
        tmpl = (
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDo'
            'cument/2006/extended-properties" xmlns:vt="http://schemas.ope'
            'nxmlformats.org/officeDocument/2006/docPropsVTypes">%s</Prope'
            'rties>\n'
        )
        if not tagname:
            child_element = ''
        elif not str_val:
            child_element = '\n  <%s/>\n' % tagname
        else:
            child_element = (
                '\n  <%s>%s</%s>\n' % (tagname, str_val, tagname)
            )
        return tmpl % child_element

    @pytest.fixture
    def extended_properties(self):
        element = parse_xml(
            b'<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>'
            b'\n<Properties xmlns="http://schemas.openxmlformats.org/office'
            b'Document/2006/extended-properties" xmlns:vt="http://schemas.o'
            b'penxmlformats.org/officeDocument/2006/docPropsVTypes">\n'
            b'  <Template>Normal.dotm</Template>\n'
            b'  <TotalTime>120</TotalTime>\n'
            b'  <Pages>5</Pages>\n'
            b'  <Words>1000</Words>\n'
            b'  <Characters>5000</Characters>\n'
            b'  <Application>Microsoft Word</Application>\n'
            b'  <Lines>50</Lines>\n'
            b'  <Paragraphs>10</Paragraphs>\n'
            b'  <Manager>John Doe</Manager>\n'
            b'  <Company>Acme Corp</Company>\n'
            b'  <AppVersion>16.0</AppVersion>\n'
            b'</Properties>\n'
        )
        return ExtendedProperties(element)

    @pytest.fixture
    def extended_properties_empty(self):
        element = parse_xml(
            b'<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>'
            b'\n<Properties xmlns="http://schemas.openxmlformats.org/office'
            b'Document/2006/extended-properties" xmlns:vt="http://schemas.o'
            b'penxmlformats.org/officeDocument/2006/docPropsVTypes">\n'
            b'</Properties>\n'
        )
        return ExtendedProperties(element)
