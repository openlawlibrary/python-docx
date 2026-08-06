"""Unit-test suite for the `docx.oxml.sdts` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.sdts import CT_SdtBase, CT_SdtPr

from ..unitutil.cxml import element


class DescribeCT_SdtBase:
    """Unit-test suite for `docx.oxml.sdts.CT_SdtBase`."""

    def it_knows_its_name(self):
        sdt = cast(CT_SdtBase, element("w:sdt/w:sdtPr/w:tag{w:val=bmk1}"))
        assert sdt.name == "bmk1"

    def it_has_no_name_when_sdtPr_is_not_present(self):
        sdt = cast(CT_SdtBase, element("w:sdt"))
        assert sdt.name is None

    @pytest.mark.parametrize(
        ("cxml", "expected_text"),
        [
            ("w:sdt", ""),
            ('w:sdt/w:sdtContent/w:r/w:t"foobar"', "foobar"),
            (
                'w:sdt/w:sdtContent/(w:p/w:r/w:t"foo",w:p/w:r/w:t"bar")',
                "foobar",
            ),
        ],
    )
    def it_knows_its_text(self, cxml: str, expected_text: str):
        sdt = cast(CT_SdtBase, element(cxml))
        assert sdt.text == expected_text


class DescribeCT_SdtPr:
    """Unit-test suite for `docx.oxml.sdts.CT_SdtPr`."""

    def it_can_change_its_name(self):
        sdtPr = cast(CT_SdtPr, element("w:sdtPr"))
        sdtPr.name = "bmk1"
        assert sdtPr.name == "bmk1"

    def it_can_change_its_alias(self):
        sdtPr = cast(CT_SdtPr, element("w:sdtPr"))
        sdtPr.alias_val = "Bookmark One"
        assert sdtPr.alias_val == "Bookmark One"

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:sdtPr", False),
            ("w:sdtPr/w:showingPlcHdr", True),
        ],
    )
    def it_knows_whether_its_placeholder_is_active(self, cxml: str, expected_value: bool):
        sdtPr = cast(CT_SdtPr, element(cxml))
        assert sdtPr.active_placeholder is expected_value

    def it_can_activate_and_deactivate_its_placeholder(self):
        sdtPr = cast(CT_SdtPr, element("w:sdtPr"))

        sdtPr.active_placeholder = True
        assert sdtPr.active_placeholder is True

        sdtPr.active_placeholder = False
        assert sdtPr.active_placeholder is False
