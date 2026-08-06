"""Test suite for the docx.oxml.text.paragraph module."""

from __future__ import annotations

from typing import cast

from docx.oxml.text.paragraph import CT_P

from ...unitutil.cxml import element


class DescribeCT_P:
    """Unit-test suite for the CT_P (paragraph, <w:p>) element."""

    def it_can_add_a_hyperlink_referencing_a_relationship(self):
        p = cast(CT_P, element("w:p"))

        hyperlink = p.add_hyperlink("open oll", rId="rId6")

        assert hyperlink.rId == "rId6"
        assert hyperlink.anchor is None
        assert hyperlink.r_lst[0].text == "open oll"
        assert hyperlink.r_lst[0].style == "Hyperlink"
        assert p.hyperlink_lst == [hyperlink]

    def it_can_add_a_hyperlink_referencing_a_bookmark(self):
        p = cast(CT_P, element("w:p"))

        hyperlink = p.add_hyperlink("see bookmark", anchor="bmk1")

        assert hyperlink.anchor == "bmk1"
        assert hyperlink.rId is None
        assert hyperlink.r_lst[0].text == "see bookmark"
        assert hyperlink.r_lst[0].style == "Hyperlink"

    def it_knows_the_runs_it_contains_including_those_nested_in_a_hyperlink(self):
        p = cast(
            CT_P,
            element('w:p/(w:r/w:t"a",w:hyperlink/(w:r/w:t"b",w:r/w:t"c"),w:r/w:t"d")'),
        )

        all_runs = p.all_runs

        assert [r.text for r in all_runs] == ["a", "b", "c", "d"]
