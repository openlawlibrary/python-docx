# encoding: utf-8

"""
Unit test suite for the docx.oxml.extendedprops module.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from docx.oxml import parse_xml
from docx.oxml.extendedprops import CT_ExtendedProperties
from docx.oxml.ns import nsmap


class DescribeCT_ExtendedProperties(object):

    def it_can_construct_a_new_element(self):
        props = CT_ExtendedProperties.new()
        assert props.tag == '{%s}Properties' % nsmap['ep']

    # --- Integer property getter tests ---

    def it_gets_integer_property_value(self, int_prop_get_fixture):
        props, prop_name, expected_value = int_prop_get_fixture
        actual_value = getattr(props, prop_name)
        assert actual_value == expected_value

    def it_returns_zero_when_integer_element_is_missing(self, int_prop_missing_fixture):
        props, prop_name = int_prop_missing_fixture
        assert getattr(props, prop_name) == 0

    def it_returns_zero_when_integer_element_has_no_text(self, int_prop_empty_text_fixture):
        props, prop_name = int_prop_empty_text_fixture
        assert getattr(props, prop_name) == 0

    def it_returns_zero_when_integer_element_has_invalid_text(self, int_prop_invalid_fixture):
        props, prop_name = int_prop_invalid_fixture
        assert getattr(props, prop_name) == 0

    def it_returns_zero_when_integer_element_has_negative_value(self, int_prop_negative_fixture):
        props, prop_name = int_prop_negative_fixture
        assert getattr(props, prop_name) == 0

    # --- Integer property setter tests ---

    def it_can_set_integer_property_value(self, int_prop_set_fixture):
        props, prop_name, xml_tagname, value = int_prop_set_fixture
        setattr(props, prop_name, value)
        # Verify XML element was actually created with correct value
        element = getattr(props, xml_tagname)
        assert element is not None, "XML element was not created"
        assert element.text == str(value), "XML element text doesn't match"

    def it_raises_on_negative_integer_value(self, int_prop_set_negative_fixture):
        props, prop_name = int_prop_set_negative_fixture
        with pytest.raises(ValueError):
            setattr(props, prop_name, -1)

    def it_raises_on_non_integer_value(self, int_prop_set_non_int_fixture):
        props, prop_name = int_prop_set_non_int_fixture
        with pytest.raises(ValueError):
            setattr(props, prop_name, 'not an int')

    # --- String property getter tests ---

    def it_gets_string_property_value(self, text_prop_get_fixture):
        props, prop_name, expected_value = text_prop_get_fixture
        actual_value = getattr(props, prop_name)
        assert actual_value == expected_value

    def it_returns_empty_string_when_text_element_is_missing(self, text_prop_missing_fixture):
        props, prop_name = text_prop_missing_fixture
        assert getattr(props, prop_name) == ''

    def it_returns_empty_string_when_text_element_has_no_text(self, text_prop_empty_fixture):
        props, prop_name = text_prop_empty_fixture
        assert getattr(props, prop_name) == ''

    # --- String property setter tests ---

    def it_can_set_string_property_value(self, text_prop_set_fixture):
        props, prop_name, xml_tagname, value = text_prop_set_fixture
        setattr(props, prop_name, value)
        # Verify XML element was actually created with correct value
        element = getattr(props, xml_tagname)
        assert element is not None, "XML element was not created"
        assert element.text == value, "XML element text doesn't match"

    def it_converts_none_to_empty_string_on_set(self, text_prop_set_none_fixture):
        props, prop_name, xml_tagname = text_prop_set_none_fixture
        setattr(props, prop_name, None)
        # Verify XML element was created with empty string
        element = getattr(props, xml_tagname)
        assert element is not None, "XML element was not created"
        assert element.text == '', "XML element text should be empty string"

    def it_converts_non_string_to_string_on_set(self, text_prop_set_non_string_fixture):
        props, prop_name, xml_tagname = text_prop_set_non_string_fixture
        setattr(props, prop_name, 123)
        # Verify XML element was created with string conversion
        element = getattr(props, xml_tagname)
        assert element is not None, "XML element was not created"
        assert element.text == '123', "XML element text should be '123'"

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def empty_props(self):
        return CT_ExtendedProperties.new()

    @pytest.fixture
    def props_with_values(self):
        return parse_xml(
            b'<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n'
            b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument'
            b'/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats'
            b'.org/officeDocument/2006/docPropsVTypes">\n'
            b'  <TotalTime>120</TotalTime>\n'
            b'  <Pages>5</Pages>\n'
            b'  <Words>1000</Words>\n'
            b'  <Characters>5000</Characters>\n'
            b'  <Lines>50</Lines>\n'
            b'  <Paragraphs>10</Paragraphs>\n'
            b'  <CharactersWithSpaces>6000</CharactersWithSpaces>\n'
            b'  <DocSecurity>4</DocSecurity>\n'
            b'  <Template>Normal.dotm</Template>\n'
            b'  <Manager>John Doe</Manager>\n'
            b'  <Company>Acme Corp</Company>\n'
            b'  <Application>Microsoft Word</Application>\n'
            b'  <AppVersion>16.0</AppVersion>\n'
            b'  <HyperlinkBase>http://example.com</HyperlinkBase>\n'
            b'</Properties>\n'
        )

    # --- Integer property fixtures ---

    @pytest.fixture(params=[
        ('totalTime_int', 120),
        ('pages_int', 5),
        ('words_int', 1000),
        ('characters_int', 5000),
        ('lines_int', 50),
        ('paragraphs_int', 10),
        ('charactersWithSpaces_int', 6000),
        ('docSecurity_int', 4),
    ])
    def int_prop_get_fixture(self, request, props_with_values):
        prop_name, expected_value = request.param
        return props_with_values, prop_name, expected_value

    @pytest.fixture(params=[
        'totalTime_int',
        'pages_int',
        'words_int',
        'characters_int',
        'lines_int',
        'paragraphs_int',
        'charactersWithSpaces_int',
        'docSecurity_int',
    ])
    def int_prop_missing_fixture(self, request, empty_props):
        return empty_props, request.param

    @pytest.fixture(params=[
        ('TotalTime', 'totalTime_int'),
        ('Pages', 'pages_int'),
        ('Words', 'words_int'),
    ])
    def int_prop_empty_text_fixture(self, request):
        tagname, prop_name = request.param
        props = parse_xml(
            ('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument'
             '/2006/extended-properties"><%s/></Properties>' % tagname).encode('utf-8')
        )
        return props, prop_name

    @pytest.fixture(params=[
        ('TotalTime', 'totalTime_int'),
        ('Pages', 'pages_int'),
        ('Words', 'words_int'),
    ])
    def int_prop_invalid_fixture(self, request):
        tagname, prop_name = request.param
        props = parse_xml(
            ('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument'
             '/2006/extended-properties"><%s>not a number</%s></Properties>'
             % (tagname, tagname)).encode('utf-8')
        )
        return props, prop_name

    @pytest.fixture(params=[
        ('TotalTime', 'totalTime_int'),
        ('Pages', 'pages_int'),
        ('Words', 'words_int'),
    ])
    def int_prop_negative_fixture(self, request):
        tagname, prop_name = request.param
        props = parse_xml(
            ('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument'
             '/2006/extended-properties"><%s>-5</%s></Properties>'
             % (tagname, tagname)).encode('utf-8')
        )
        return props, prop_name

    @pytest.fixture(params=[
        ('totalTime_int', 'TotalTime'),
        ('pages_int', 'Pages'),
        ('words_int', 'Words'),
        ('characters_int', 'Characters'),
        ('lines_int', 'Lines'),
        ('paragraphs_int', 'Paragraphs'),
    ])
    def int_prop_set_fixture(self, request, empty_props):
        prop_name, xml_tagname = request.param
        return empty_props, prop_name, xml_tagname, 42

    @pytest.fixture(params=[
        'totalTime_int',
        'pages_int',
        'words_int',
    ])
    def int_prop_set_negative_fixture(self, request, empty_props):
        return empty_props, request.param

    @pytest.fixture(params=[
        'totalTime_int',
        'pages_int',
        'words_int',
    ])
    def int_prop_set_non_int_fixture(self, request, empty_props):
        return empty_props, request.param

    # --- String property fixtures ---

    @pytest.fixture(params=[
        ('template_text', 'Normal.dotm'),
        ('manager_text', 'John Doe'),
        ('company_text', 'Acme Corp'),
        ('application_text', 'Microsoft Word'),
        ('appVersion_text', '16.0'),
        ('hyperlinkBase_text', 'http://example.com'),
    ])
    def text_prop_get_fixture(self, request, props_with_values):
        prop_name, expected_value = request.param
        return props_with_values, prop_name, expected_value

    @pytest.fixture(params=[
        'template_text',
        'manager_text',
        'company_text',
        'application_text',
        'appVersion_text',
        'hyperlinkBase_text',
    ])
    def text_prop_missing_fixture(self, request, empty_props):
        return empty_props, request.param

    @pytest.fixture(params=[
        ('Template', 'template_text'),
        ('Manager', 'manager_text'),
        ('Company', 'company_text'),
    ])
    def text_prop_empty_fixture(self, request):
        tagname, prop_name = request.param
        props = parse_xml(
            ('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument'
             '/2006/extended-properties"><%s/></Properties>' % tagname).encode('utf-8')
        )
        return props, prop_name

    @pytest.fixture(params=[
        ('template_text', 'Template'),
        ('manager_text', 'Manager'),
        ('company_text', 'Company'),
        ('application_text', 'Application'),
        ('appVersion_text', 'AppVersion'),
        ('hyperlinkBase_text', 'HyperlinkBase'),
    ])
    def text_prop_set_fixture(self, request, empty_props):
        prop_name, xml_tagname = request.param
        return empty_props, prop_name, xml_tagname, 'Test Value'

    @pytest.fixture(params=[
        ('template_text', 'Template'),
        ('manager_text', 'Manager'),
        ('company_text', 'Company'),
    ])
    def text_prop_set_none_fixture(self, request, empty_props):
        prop_name, xml_tagname = request.param
        return empty_props, prop_name, xml_tagname

    @pytest.fixture(params=[
        ('template_text', 'Template'),
        ('manager_text', 'Manager'),
        ('company_text', 'Company'),
    ])
    def text_prop_set_non_string_fixture(self, request, empty_props):
        prop_name, xml_tagname = request.param
        return empty_props, prop_name, xml_tagname
