"""Test suite for the docx.oxml.text.run module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.text.run import CT_R

from ...unitutil.cxml import element, xml


class DescribeCT_R:
    """Unit-test suite for the CT_R (run, <w:r>) element."""

    @pytest.mark.parametrize(
        ("initial_cxml", "text", "expected_cxml"),
        [
            ("w:r", "foobar", 'w:r/w:t"foobar"'),
            ("w:r", "foobar ", 'w:r/w:t{xml:space=preserve}"foobar "'),
            (
                "w:r/(w:rPr/w:rStyle{w:val=emphasis}, w:cr)",
                "foobar",
                'w:r/(w:rPr/w:rStyle{w:val=emphasis}, w:cr, w:t"foobar")',
            ),
        ],
    )
    def it_can_add_a_t_preserving_edge_whitespace(
        self, initial_cxml: str, text: str, expected_cxml: str
    ):
        r = cast(CT_R, element(initial_cxml))
        expected_xml = xml(expected_cxml)

        r.add_t(text)

        assert r.xml == expected_xml

    def it_can_assemble_the_text_in_the_run(self):
        cxml = 'w:r/(w:br,w:cr,w:noBreakHyphen,w:ptab,w:t"foobar",w:tab)'
        r = cast(CT_R, element(cxml))

        assert r.text == "\n\n-\tfoobar\t"

    def it_can_add_a_footnote_reference(self):
        r = cast(CT_R, element("w:r"))

        footnoteReference = r.add_footnoteReference(3)

        assert footnoteReference.id == 3
        assert r.xml == xml(
            "w:r/(w:rPr/w:rStyle{w:val=FootnoteReference},w:footnoteReference{w:id=3})"
        )

    def it_can_add_an_endnote_reference(self):
        r = cast(CT_R, element("w:r"))

        endnoteReference = r.add_endnoteReference(3)

        assert endnoteReference.id == 3
        assert r.xml == xml(
            "w:r/(w:rPr/w:rStyle{w:val=EndnoteReference},w:endnoteReference{w:id=3})"
        )

    def it_can_add_a_footnote_ref_mark(self):
        r = cast(CT_R, element("w:r"))

        r.add_footnoteRef()

        assert r.xml == xml("w:r/w:footnoteRef")

    def it_can_add_an_endnote_ref_mark(self):
        r = cast(CT_R, element("w:r"))

        r.add_endnoteRef()

        assert r.xml == xml("w:r/w:endnoteRef")

    @pytest.mark.parametrize(
        ("cxml", "expected_ids"),
        [
            ("w:r", []),
            ("w:r/w:footnoteReference{w:id=1}", [1]),
            ("w:r/(w:footnoteReference{w:id=1},w:footnoteReference{w:id=2})", [1, 2]),
        ],
    )
    def it_knows_its_footnote_reference_ids(self, cxml: str, expected_ids: list[int]):
        r = cast(CT_R, element(cxml))
        assert list(r.footnote_reference_ids) == expected_ids

    @pytest.mark.parametrize(
        ("cxml", "expected_ids"),
        [
            ("w:r", []),
            ("w:r/w:endnoteReference{w:id=1}", [1]),
            ("w:r/(w:endnoteReference{w:id=1},w:endnoteReference{w:id=2})", [1, 2]),
        ],
    )
    def it_knows_its_endnote_reference_ids(self, cxml: str, expected_ids: list[int]):
        r = cast(CT_R, element(cxml))
        assert list(r.endnote_reference_ids) == expected_ids

    def it_can_increment_its_footnote_reference_ids(self):
        r = cast(
            CT_R,
            element("w:r/(w:footnoteReference{w:id=1},w:footnoteReference{w:id=2})"),
        )

        r.increment_containing_footnote_reference_ids()

        assert list(r.footnote_reference_ids) == [2, 3]

    def it_can_increment_its_endnote_reference_ids(self):
        r = cast(
            CT_R,
            element("w:r/(w:endnoteReference{w:id=1},w:endnoteReference{w:id=2})"),
        )

        r.increment_containing_endnote_reference_ids()

        assert list(r.endnote_reference_ids) == [2, 3]

    @pytest.mark.parametrize(
        ("cxml", "expected_cxml"),
        [
            (
                'w:r/(w:rPr,w:t"foo",w:footnoteReference{w:id=1})',
                "w:r/(w:rPr,w:footnoteReference{w:id=1})",
            ),
            (
                'w:r/(w:rPr,w:t"foo",w:endnoteReference{w:id=1})',
                "w:r/(w:rPr,w:endnoteReference{w:id=1})",
            ),
            ('w:r/w:t"foo"', "w:r"),
        ],
    )
    def it_preserves_footnote_and_endnote_references_when_clearing_content(
        self, cxml: str, expected_cxml: str
    ):
        r = cast(CT_R, element(cxml))

        r.clear_content()

        assert r.xml == xml(expected_cxml)
