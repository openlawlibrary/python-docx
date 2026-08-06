"""Unit test suite for the numbering-resolution additions to docx.oxml.numbering."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.numbering import CT_AbstractNum, CT_Numbering
from docx.oxml.text.paragraph import CT_P

from ..unitutil.cxml import element


def _numbering_with_one_decimal_list() -> CT_Numbering:
    """A `w:numbering` element with a single decimal-numbered abstract list
    (abstractNumId=0) referenced by numId=1."""
    return cast(
        CT_Numbering,
        element(
            "w:numbering/("
            "w:abstractNum{w:abstractNumId=0}/w:lvl{w:ilvl=0}/(w:start{w:val=1},"
            "w:numFmt{w:val=decimal},w:lvlText{w:val=%1.}),"
            "w:num{w:numId=1}/w:abstractNumId{w:val=0}"
            ")"
        ),
    )


def _numbered_body(count: int) -> list[CT_P]:
    """A `w:body` containing `count` paragraphs, each numbered at ilvl=0/numId=1, and
    return its `w:p` children."""
    p_cxml = "w:p/w:pPr/w:numPr/(w:ilvl{w:val=0},w:numId{w:val=1})"
    body = element("w:body/(%s)" % ",".join([p_cxml] * count))
    return body.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")


class DescribeCT_Numbering:
    """Unit-test suite for the numbering-resolution additions to `CT_Numbering`."""

    @pytest.mark.parametrize(
        ("fmt", "num", "expected_value"),
        [
            ("decimal", 1, 1),
            ("decimal", 12, 12),
            ("lowerLetter", 1, "a"),
            ("lowerLetter", 26, "z"),
            ("lowerLetter", 27, "aa"),
            ("lowerLetter", 52, "zz"),
            ("upperLetter", 1, "A"),
            ("upperLetter", 27, "AA"),
            ("lowerRoman", 4, "iv"),
            ("lowerRoman", 12, "xii"),
            ("upperRoman", 4, "IV"),
            ("upperRoman", 2024, "MMXXIV"),
            ("none", 5, ""),
        ],
    )
    def it_formats_a_count_per_numFmt(self, fmt: str, num: int, expected_value):
        numbering = cast(CT_Numbering, element("w:numbering"))
        assert numbering.fmt_map[fmt](num) == expected_value

    def it_resolves_the_abstractNum_referenced_by_a_numId(self):
        numbering = _numbering_with_one_decimal_list()
        abstractNum = numbering.get_abstractNum(1)
        assert isinstance(abstractNum, CT_AbstractNum)
        assert abstractNum.abstractNumId == 0

    def it_returns_None_for_an_abstractNum_lookup_on_an_unknown_numId(self):
        numbering = _numbering_with_one_decimal_list()
        assert numbering.get_abstractNum(42) is None

    def it_computes_sequential_labels_for_a_simple_decimal_list(self):
        numbering = _numbering_with_one_decimal_list()
        p0, p1, p2 = _numbered_body(3)

        assert numbering.get_num_for_p(p0, {}) == "1.\t"
        assert numbering.get_num_for_p(p1, {}) == "2.\t"
        assert numbering.get_num_for_p(p2, {}) == "3.\t"

    def it_omits_the_suffix_when_append_suffix_is_False(self):
        numbering = _numbering_with_one_decimal_list()
        (p0,) = _numbered_body(1)

        assert numbering.get_num_for_p(p0, {}, append_suffix=False) == "1."

    def it_raises_for_a_paragraph_with_no_numbering_properties(self):
        """Matches `Paragraph.number`'s contract: this AttributeError is caught at the
        proxy layer and surfaced there as `None`, not swallowed here."""
        numbering = _numbering_with_one_decimal_list()
        p = cast(CT_P, element("w:p"))

        with pytest.raises(AttributeError):
            numbering.get_num_for_p(p, {})

    def it_can_start_a_new_list_at_a_given_level(self):
        numbering = _numbering_with_one_decimal_list()
        (p,) = _numbered_body(1)
        # -- clear the numId/ilvl this paragraph was born with, simulating a
        # -- paragraph that references a list definition but hasn't joined a
        # -- specific list instance yet.
        assert p.pPr is not None
        assert p.pPr.numPr is not None

        numbering.set_li_lvl(p, {}, prev_p=None, ilvl=0)

        assert p.pPr.numPr.numId is not None
        assert p.pPr.numPr.ilvl is not None
        assert p.pPr.numPr.ilvl.val == 0
        # -- starting a new list allocates a fresh numId, distinct from the one the
        # -- paragraph started with, with its own startOverride --
        new_numId = p.pPr.numPr.numId.val
        assert numbering.get_startOverride(new_numId, 0) == 1

    def it_can_join_an_existing_list_from_a_previous_paragraph(self):
        numbering = _numbering_with_one_decimal_list()
        p0, p1 = _numbered_body(2)

        numbering.set_li_lvl(p1, {}, prev_p=p0, ilvl=None)

        assert p1.pPr is not None
        assert p1.pPr.numPr is not None
        assert p0.pPr is not None
        assert p0.pPr.numPr is not None
        assert p1.pPr.numPr.numId is not None
        assert p0.pPr.numPr.numId is not None
        assert p1.pPr.numPr.numId.val == p0.pPr.numPr.numId.val
        assert p1.pPr.numPr.ilvl is not None
        assert p0.pPr.numPr.ilvl is not None
        assert p1.pPr.numPr.ilvl.val == p0.pPr.numPr.ilvl.val
