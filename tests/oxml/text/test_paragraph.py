"""Test suite for the docx.oxml.text.paragraph module."""

from __future__ import annotations

from typing import cast

from docx.oxml.text.paragraph import CT_P

from ...unitutil.cxml import element, xml


class DescribeCT_P:
    """Unit-test suite for the CT_P (paragraph, <w:p>) element."""

    def it_strips_hidden_content_from_a_field_with_no_cached_result(self):
        """A `begin`/`end` pair with no `separate` has no visible result yet."""
        p = cast(
            CT_P,
            element(
                'w:p/(w:r/w:t"before ",w:r/w:fldChar{w:fldCharType=begin},'
                'w:r/w:instrText"DATE",w:r/w:fldChar{w:fldCharType=end},'
                'w:r/w:t" after")'
            ),
        )

        p.strip_hidden_fld_char_content()

        assert p.xml == xml('w:p/(w:r/w:t"before ",w:r/w:t" after")')

    def it_keeps_the_cached_result_of_a_field_that_has_one(self):
        """Content between `separate` and `end` is the field's visible result."""
        p = cast(
            CT_P,
            element(
                'w:p/(w:r/w:fldChar{w:fldCharType=begin},w:r/w:instrText"DATE",'
                'w:r/w:fldChar{w:fldCharType=separate},w:r/w:t"12/5/2023",'
                "w:r/w:fldChar{w:fldCharType=end})"
            ),
        )

        p.strip_hidden_fld_char_content()

        assert p.xml == xml('w:p/w:r/w:t"12/5/2023"')

    def it_strips_hidden_content_nested_inside_a_hyperlink(self):
        p = cast(
            CT_P,
            element(
                "w:p/w:hyperlink/(w:r/w:fldChar{w:fldCharType=begin},"
                'w:r/w:instrText"PAGEREF _Toc1",w:r/w:fldChar{w:fldCharType=separate},'
                'w:r/w:t"1",w:r/w:fldChar{w:fldCharType=end})'
            ),
        )

        p.strip_hidden_fld_char_content()

        assert p.xml == xml('w:p/w:hyperlink/w:r/w:t"1"')

    def it_is_idempotent(self):
        p = cast(
            CT_P,
            element(
                "w:p/(w:r/w:fldChar{w:fldCharType=begin},w:r/w:fldChar{w:fldCharType=separate},"
                'w:r/w:t"result",w:r/w:fldChar{w:fldCharType=end})'
            ),
        )

        p.strip_hidden_fld_char_content()
        p.strip_hidden_fld_char_content()

        assert p.xml == xml('w:p/w:r/w:t"result"')

    def it_leaves_a_paragraph_with_no_fields_unchanged(self):
        p = cast(CT_P, element('w:p/w:r/w:t"foobar"'))

        p.strip_hidden_fld_char_content()

        assert p.xml == xml('w:p/w:r/w:t"foobar"')

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
