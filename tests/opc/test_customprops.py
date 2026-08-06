"""Unit test suite for the `docx.opc.customprops` module."""

from __future__ import annotations

from docx.opc.customprops import CustomProperties
from docx.oxml.customprops import CT_CustomProperties


class DescribeCustomProperties:
    """Unit-test suite for `docx.opc.customprops.CustomProperties`."""

    def it_starts_empty_on_a_new_custom_properties_part(self):
        custom_properties = CustomProperties(CT_CustomProperties.new())
        assert len(custom_properties) == 0

    def it_can_add_and_read_back_a_string_property(self):
        custom_properties = CustomProperties(CT_CustomProperties.new())
        custom_properties["str_test"] = "test123"
        assert custom_properties["str_test"] == "test123"
        assert len(custom_properties) == 1

    def it_can_add_and_read_back_a_bool_property(self):
        custom_properties = CustomProperties(CT_CustomProperties.new())
        custom_properties["bool_test"] = False
        assert custom_properties["bool_test"] is False

    def it_can_add_and_read_back_an_int_property(self):
        custom_properties = CustomProperties(CT_CustomProperties.new())
        custom_properties["num_test"] = 777
        assert custom_properties["num_test"] == 777

    def it_can_update_an_existing_property(self):
        custom_properties = CustomProperties(CT_CustomProperties.new())
        custom_properties["num_test"] = 777
        custom_properties["num_test"] = 42
        assert custom_properties["num_test"] == 42
        assert len(custom_properties) == 1

    def it_returns_None_for_a_property_not_present(self):
        custom_properties = CustomProperties(CT_CustomProperties.new())
        assert custom_properties["does-not-exist"] is None
