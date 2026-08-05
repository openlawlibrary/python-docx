"""Unit test suite for the `docx.opc.extendedprops` module."""

from __future__ import annotations

import pytest

from docx.opc.extendedprops import ExtendedProperties
from docx.oxml.extendedprops import CT_ExtendedProperties


class DescribeExtendedProperties:
    """Unit-test suite for `docx.opc.extendedprops.ExtendedProperties`."""

    @pytest.mark.parametrize(
        ("prop_name", "default_value", "new_value"),
        [
            ("total_time", 0, 42),
            ("pages", 0, 4),
            ("words", 0, 640),
            ("characters", 0, 3457),
            ("lines", 0, 100),
            ("paragraphs", 0, 157),
            ("characters_with_spaces", 0, 4029),
            ("doc_security", 0, 1),
        ],
    )
    def it_reads_and_writes_integer_properties(
        self, prop_name: str, default_value: int, new_value: int
    ):
        extended_properties = ExtendedProperties(CT_ExtendedProperties.new())
        assert getattr(extended_properties, prop_name) == default_value
        setattr(extended_properties, prop_name, new_value)
        assert getattr(extended_properties, prop_name) == new_value

    @pytest.mark.parametrize(
        "prop_name",
        [
            "template",
            "manager",
            "company",
            "application",
            "app_version",
            "hyperlink_base",
        ],
    )
    def it_reads_and_writes_string_properties(self, prop_name: str):
        extended_properties = ExtendedProperties(CT_ExtendedProperties.new())
        assert getattr(extended_properties, prop_name) == ""
        setattr(extended_properties, prop_name, "some value")
        assert getattr(extended_properties, prop_name) == "some value"

    def it_raises_setting_a_negative_integer_property(self):
        extended_properties = ExtendedProperties(CT_ExtendedProperties.new())
        with pytest.raises(ValueError, match="requires non-negative int"):
            extended_properties.words = -1
