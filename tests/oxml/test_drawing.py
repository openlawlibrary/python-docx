"""Test suite for the docx.oxml.drawing module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.drawing import CT_Drawing
from docx.oxml.shape import CT_Blip

from ..unitutil.cxml import element


class DescribeCT_Drawing:
    """Unit-test suite for the CT_Drawing (`<w:drawing>`) element."""

    @pytest.mark.parametrize(
        ("drawing_cxml", "expected_value"),
        [
            (
                "w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill"
                "/a:blip{r:embed=rId1}",
                "rId1",
            ),
            (
                "w:drawing/wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill"
                "/a:blip{r:link=rId2}",
                None,
            ),
        ],
    )
    def it_finds_its_blip_regardless_of_inline_or_floating_nesting(
        self, drawing_cxml: str, expected_value: str | None
    ):
        drawing = cast(CT_Drawing, element(drawing_cxml))

        blip = drawing.blip

        assert isinstance(blip, CT_Blip)
        assert blip.embed == expected_value

    def it_has_no_blip_when_it_contains_no_picture(self):
        drawing = cast(CT_Drawing, element("w:drawing/wp:inline/a:graphic/a:graphicData/a:grpSp"))

        assert drawing.blip is None
