"""Unit test suite for the docx.text.paragraph module."""

from __future__ import annotations

from types import SimpleNamespace
from typing import List, cast

import pytest

import docx
from docx import types as t
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.run import CT_R
from docx.parts.document import DocumentPart
from docx.text.hyperlink import Hyperlink
from docx.text.paragraph import Paragraph
from docx.text.parfmt import ParagraphFormat
from docx.text.run import Run

from ..unitutil.cxml import element, xml
from ..unitutil.mock import call, class_mock, instance_mock, method_mock, property_mock


class DescribeParagraph:
    """Unit-test suite for `docx.text.run.Paragraph`."""

    @pytest.mark.parametrize(
        ("p_cxml", "expected_value"),
        [
            ("w:p/w:r", False),
            ('w:p/w:r/w:t"foobar"', False),
            ('w:p/w:hyperlink/w:r/(w:t"abc",w:lastRenderedPageBreak,w:t"def")', True),
            ("w:p/w:r/(w:lastRenderedPageBreak, w:lastRenderedPageBreak)", True),
        ],
    )
    def it_knows_whether_it_contains_a_page_break(
        self, p_cxml: str, expected_value: bool, fake_parent: t.ProvidesStoryPart
    ):
        p = cast(CT_P, element(p_cxml))
        paragraph = Paragraph(p, fake_parent)

        assert paragraph.contains_page_break == expected_value

    @pytest.mark.parametrize(
        ("p_cxml", "count"),
        [
            ("w:p", 0),
            ("w:p/w:r", 0),
            ("w:p/w:hyperlink", 1),
            ("w:p/(w:r,w:hyperlink,w:r)", 1),
            ("w:p/(w:r,w:hyperlink,w:r,w:hyperlink)", 2),
            ("w:p/(w:hyperlink,w:r,w:hyperlink,w:r)", 2),
        ],
    )
    def it_provides_access_to_the_hyperlinks_it_contains(
        self, p_cxml: str, count: int, fake_parent: t.ProvidesStoryPart
    ):
        p = cast(CT_P, element(p_cxml))
        paragraph = Paragraph(p, fake_parent)

        hyperlinks = paragraph.hyperlinks

        actual = [type(item).__name__ for item in hyperlinks]
        expected = ["Hyperlink" for _ in range(count)]
        assert actual == expected, f"expected: {expected}, got: {actual}"

    @pytest.mark.parametrize(
        ("p_cxml", "expected"),
        [
            ("w:p", []),
            ("w:p/w:r", ["Run"]),
            ("w:p/w:hyperlink", ["Hyperlink"]),
            ("w:p/(w:r,w:hyperlink,w:r)", ["Run", "Hyperlink", "Run"]),
            ("w:p/(w:hyperlink,w:r,w:hyperlink)", ["Hyperlink", "Run", "Hyperlink"]),
        ],
    )
    def it_can_iterate_its_inner_content_items(
        self, p_cxml: str, expected: List[str], fake_parent: t.ProvidesStoryPart
    ):
        p = cast(CT_P, element(p_cxml))
        paragraph = Paragraph(p, fake_parent)

        inner_content = paragraph.iter_inner_content()

        actual = [type(item).__name__ for item in inner_content]
        assert actual == expected, f"expected: {expected}, got: {actual}"

    def it_knows_its_paragraph_style(self, style_get_fixture):
        paragraph, style_id_, style_ = style_get_fixture
        style = paragraph.style
        paragraph.part.get_style.assert_called_once_with(style_id_, WD_STYLE_TYPE.PARAGRAPH)
        assert style is style_

    def it_can_change_its_paragraph_style(self, style_set_fixture):
        paragraph, value, expected_xml = style_set_fixture

        paragraph.style = value

        paragraph.part.get_style_id.assert_called_once_with(value, WD_STYLE_TYPE.PARAGRAPH)
        assert paragraph._p.xml == expected_xml

    @pytest.mark.parametrize(
        ("p_cxml", "count"),
        [
            ("w:p", 0),
            ("w:p/w:r", 0),
            ("w:p/w:r/w:lastRenderedPageBreak", 1),
            ("w:p/w:hyperlink/w:r/w:lastRenderedPageBreak", 1),
            (
                "w:p/(w:r/w:lastRenderedPageBreak,w:hyperlink/w:r/w:lastRenderedPageBreak)",
                2,
            ),
            (
                "w:p/(w:hyperlink/w:r/w:lastRenderedPageBreak,w:r,"
                "w:r/w:lastRenderedPageBreak,w:r,w:hyperlink)",
                2,
            ),
        ],
    )
    def it_provides_access_to_the_rendered_page_breaks_it_contains(
        self, p_cxml: str, count: int, fake_parent: t.ProvidesStoryPart
    ):
        p = cast(CT_P, element(p_cxml))
        paragraph = Paragraph(p, fake_parent)

        rendered_page_breaks = paragraph.rendered_page_breaks

        actual = [type(item).__name__ for item in rendered_page_breaks]
        expected = ["RenderedPageBreak" for _ in range(count)]
        assert actual == expected, f"expected: {expected}, got: {actual}"

    @pytest.mark.parametrize(
        ("p_cxml", "expected_value"),
        [
            ("w:p", ""),
            ("w:p/w:r", ""),
            ("w:p/w:r/w:t", ""),
            ('w:p/w:r/w:t"foo"', "foo"),
            ('w:p/w:r/(w:t"foo", w:t"bar")', "foobar"),
            ('w:p/w:r/(w:t"fo ", w:t"bar")', "fo bar"),
            ('w:p/w:r/(w:t"foo", w:tab, w:t"bar")', "foo\tbar"),
            ('w:p/w:r/(w:t"foo", w:br,  w:t"bar")', "foo\nbar"),
            ('w:p/w:r/(w:t"foo", w:cr,  w:t"bar")', "foo\nbar"),
            (
                'w:p/(w:r/w:t"click ",w:hyperlink{r:id=rId6}/w:r/w:t"here",w:r/w:t" for more")',
                "click here for more",
            ),
            (
                'w:p/(w:r/w:t"before ",w:r/w:fldChar{w:fldCharType=begin},'
                'w:r/w:instrText"DATE",w:r/w:fldChar{w:fldCharType=end},'
                'w:r/w:t" after")',
                "before  after",
            ),
            (
                'w:p/(w:r/w:fldChar{w:fldCharType=begin},w:r/w:instrText"DATE",'
                'w:r/w:fldChar{w:fldCharType=separate},w:r/w:t"12/5/2023",'
                "w:r/w:fldChar{w:fldCharType=end})",
                "12/5/2023",
            ),
        ],
    )
    def it_knows_the_text_it_contains(self, p_cxml: str, expected_value: str):
        """Including the text of embedded hyperlinks."""
        paragraph = Paragraph(element(p_cxml), None)
        assert paragraph.text == expected_value

    def it_can_replace_the_text_it_contains(self, text_set_fixture):
        paragraph, text, expected_text = text_set_fixture
        paragraph.text = text
        assert paragraph.text == expected_text

    def it_knows_its_alignment_value(self, alignment_get_fixture):
        paragraph, expected_value = alignment_get_fixture
        assert paragraph.alignment == expected_value

    def it_can_change_its_alignment_value(self, alignment_set_fixture):
        paragraph, value, expected_xml = alignment_set_fixture
        paragraph.alignment = value
        assert paragraph._p.xml == expected_xml

    def it_provides_access_to_its_paragraph_format(self, parfmt_fixture):
        paragraph, ParagraphFormat_, paragraph_format_ = parfmt_fixture
        paragraph_format = paragraph.paragraph_format
        ParagraphFormat_.assert_called_once_with(paragraph._element)
        assert paragraph_format is paragraph_format_

    def it_provides_access_to_the_runs_it_contains(self, runs_fixture):
        paragraph, Run_, r_, r_2_, run_, run_2_ = runs_fixture
        runs = paragraph.runs
        assert Run_.mock_calls == [call(r_, paragraph), call(r_2_, paragraph)]
        assert runs == [run_, run_2_]

    @pytest.mark.parametrize(
        ("p_cxml", "expected"),
        [
            ("w:p/(w:r,w:hyperlink/(w:r,w:r),w:r)", ["Run", "Run", "Run", "Run"]),
        ],
    )
    def it_flattens_hyperlink_runs_into_the_runs_it_contains(
        self, p_cxml: str, expected: List[str], fake_parent: t.ProvidesStoryPart
    ):
        paragraph = Paragraph(cast(CT_P, element(p_cxml)), fake_parent)

        actual = [type(item).__name__ for item in paragraph.runs]

        assert actual == expected

    def it_provides_access_to_its_runs_and_hyperlinks_without_flattening(
        self, fake_parent: t.ProvidesStoryPart
    ):
        paragraph = Paragraph(
            cast(CT_P, element("w:p/(w:r,w:hyperlink/(w:r,w:r),w:r)")), fake_parent
        )

        actual = [type(item).__name__ for item in paragraph.runs_and_hyperlinks]

        assert actual == ["Run", "Hyperlink", "Run"]

    @pytest.mark.parametrize(
        ("instrText", "expected_cxml"),
        [
            (
                None,
                "w:p/(w:r/w:fldChar{w:fldCharType=begin},w:r/w:fldChar{w:fldCharType=end})",
            ),
            (
                "DATE",
                'w:p/(w:r/w:fldChar{w:fldCharType=begin},w:r/w:instrText"DATE",'
                "w:r/w:fldChar{w:fldCharType=end})",
            ),
        ],
    )
    def it_can_add_a_field(self, instrText: str | None, expected_cxml: str):
        paragraph = Paragraph(cast(CT_P, element("w:p")), None)

        paragraph.add_field(instrText)

        assert paragraph._p.xml == xml(expected_cxml)

    def it_can_add_a_run_to_itself(self, add_run_fixture):
        paragraph, text, style, style_prop_, expected_xml = add_run_fixture
        run = paragraph.add_run(text, style)
        assert paragraph._p.xml == expected_xml
        assert isinstance(run, Run)
        assert run._r is paragraph._p.r_lst[0]
        if style:
            style_prop_.assert_called_once_with(style)

    def it_can_add_a_hyperlink_from_a_url(self, part_prop_, document_part_):
        document_part_.relate_to.return_value = "rId9"
        paragraph = Paragraph(cast(CT_P, element("w:p")), None)

        hyperlink = paragraph.add_hyperlink("open oll", "https://openlawlib.org/")

        document_part_.relate_to.assert_called_once_with(
            "https://openlawlib.org/", RT.HYPERLINK, is_external=True
        )
        assert isinstance(hyperlink, Hyperlink)
        assert paragraph._p.hyperlink_lst[0].rId == "rId9"
        assert paragraph._p.hyperlink_lst[0].r_lst[0].text == "open oll"

    def it_can_add_a_hyperlink_from_a_bookmark_name(self, part_prop_, document_part_):
        paragraph = Paragraph(cast(CT_P, element("w:p")), None)

        paragraph.add_hyperlink("see bookmark", "bmk1")

        document_part_.relate_to.assert_not_called()
        assert paragraph._p.hyperlink_lst[0].anchor == "bmk1"
        assert paragraph._p.hyperlink_lst[0].r_lst[0].text == "see bookmark"

    def it_can_insert_a_paragraph_before_itself(self, insert_before_fixture):
        text, style, paragraph_, add_run_calls = insert_before_fixture
        paragraph = Paragraph(None, None)

        new_paragraph = paragraph.insert_paragraph_before(text, style)

        paragraph._insert_paragraph_before.assert_called_once_with(paragraph)
        assert new_paragraph.add_run.call_args_list == add_run_calls
        assert new_paragraph.style == style
        assert new_paragraph is paragraph_

    def it_can_remove_its_content_while_preserving_formatting(self, clear_fixture):
        paragraph, expected_xml = clear_fixture
        _paragraph = paragraph.clear()
        assert paragraph._p.xml == expected_xml
        assert _paragraph is paragraph

    def it_can_be_cloned(self):
        paragraph = Paragraph(element('w:p/w:r/w:t"foobar"'), "parent")

        clone = paragraph.clone()

        assert clone is not paragraph
        assert clone._p is not paragraph._p
        assert clone.text == "foobar"
        assert clone._parent == "parent"

    def it_can_be_pickled_and_unpickled(self):
        paragraph = Paragraph(element('w:p/w:r/w:t"foobar"'), "parent")

        state = paragraph.__getstate__()

        assert "_parent" not in state

        restored = Paragraph.__new__(Paragraph)
        restored.__setstate__(state)

        assert restored.text == "foobar"

    @pytest.mark.parametrize(
        ("p_cxml", "expected_repr"),
        [
            ('w:p/w:r/w:t"foobar"', '<p:"foobar">'),
            (
                'w:p/w:r/w:t"01234567890123456789tail"',
                '<p:"01234567890123456789...">',
            ),
            ("w:p", '<p:"EMPTY PARAGRAPH">'),
        ],
    )
    def it_has_a_repr_useful_for_debugging(self, p_cxml: str, expected_repr: str):
        paragraph = Paragraph(element(p_cxml), None)
        assert repr(paragraph) == expected_repr

    def it_can_remove_itself_from_its_container(self):
        body = element('w:body/(w:p/w:r/w:t"foo", w:p/w:r/w:t"bar")')
        paragraph = Paragraph(body[0], None)

        paragraph.remove()

        assert body.xml == xml('w:body/w:p/w:r/w:t"bar"')

    @pytest.mark.parametrize(
        ("p_cxml", "chars", "expected_text"),
        [
            ('w:p/w:r/w:t"  foobar  "', None, "foobar  "),
            ('w:p/w:r/w:t"--foobar--"', "-", "foobar--"),
        ],
    )
    def it_can_lstrip_its_text(self, p_cxml, chars, expected_text):
        paragraph = Paragraph(element(p_cxml), None)
        result = paragraph.lstrip(chars)
        assert paragraph.text == expected_text
        assert result is paragraph

    @pytest.mark.parametrize(
        ("p_cxml", "chars", "expected_text"),
        [
            ('w:p/w:r/w:t"  foobar  "', None, "  foobar"),
            ('w:p/w:r/w:t"--foobar--"', "-", "--foobar"),
        ],
    )
    def it_can_rstrip_its_text(self, p_cxml, chars, expected_text):
        paragraph = Paragraph(element(p_cxml), None)
        result = paragraph.rstrip(chars)
        assert paragraph.text == expected_text
        assert result is paragraph

    def it_can_strip_its_text(self):
        paragraph = Paragraph(element('w:p/w:r/w:t"  foobar  "'), None)
        result = paragraph.strip()
        assert paragraph.text == "foobar"
        assert result is paragraph

    def it_removes_a_run_left_empty_by_stripping(self):
        paragraph = Paragraph(element('w:p/(w:r/w:t"   ", w:r/w:t"foobar")'), None)
        paragraph.lstrip()
        assert paragraph.text == "foobar"
        assert len(paragraph.runs) == 1

    def it_can_replace_a_character_throughout_its_text(self):
        paragraph = Paragraph(element('w:p/w:r/w:t"a-b-c"'), None)
        result = paragraph.replace_char("-", "_")
        assert paragraph.text == "a_b_c"
        assert result is paragraph

    def it_can_replace_multiple_characters_throughout_its_text(self):
        paragraph = Paragraph(element('w:p/w:r/w:t"a-b_c"'), None)
        result = paragraph.replace_chars(("-", "+"), ("_", "+"))
        assert paragraph.text == "a+b+c"
        assert result is paragraph

    def it_can_insert_text_at_a_position(self):
        paragraph = Paragraph(element('w:p/w:r/w:t"foobar"'), None)
        result = paragraph.insert_text(3, "-X-")
        assert paragraph.text == "foo-X-bar"
        assert result is paragraph

    @pytest.mark.parametrize(
        ("p_cxml", "old_text", "new_text", "expected_text"),
        [
            ('w:p/w:r/w:t"foobar"', "oob", "XYZ", "fXYZar"),
            (
                'w:p/(w:r/w:t"foo", w:r/w:t"bar")',
                "oob",
                "XYZ",
                "fXYZar",
            ),
        ],
    )
    def it_can_replace_text_spanning_runs(
        self, p_cxml: str, old_text: str, new_text: str, expected_text: str
    ):
        paragraph = Paragraph(element(p_cxml), None)
        result = paragraph.replace_text(old_text, new_text)
        assert paragraph.text == expected_text
        assert result is paragraph

    def it_can_remove_a_range_of_text_within_a_single_run(self):
        paragraph = Paragraph(element('w:p/w:r/w:t"foobar"'), None)
        result = paragraph.remove_text(1, 4)
        assert paragraph.text == "far"
        assert result is paragraph

    def it_can_remove_a_range_of_text_spanning_runs(self):
        paragraph = Paragraph(element('w:p/(w:r/w:t"foo", w:r/w:t"bar")'), None)
        result = paragraph.remove_text(1, 5)
        assert paragraph.text == "fr"
        assert result is paragraph

    def it_removes_a_run_left_empty_by_removing_its_text(self):
        paragraph = Paragraph(element('w:p/(w:r/w:t"foo", w:r/w:t"bar")'), None)
        paragraph.remove_text(0, 3)
        assert paragraph.text == "bar"
        assert len(paragraph.runs) == 1

    def it_can_replace_line_breaks_with_spaces(self):
        paragraph = Paragraph(element('w:p/w:r/(w:t"foo", w:br, w:t"bar")'), None)
        assert paragraph.text == "foo\nbar"
        paragraph.remove_new_line_breaks
        assert paragraph.text == "foo bar"

    def it_can_split_itself_at_a_position(self):
        body = element('w:body/w:p/w:r/w:t"foobar"')
        paragraph = Paragraph(body[0], None)

        paras = paragraph.split(3)

        assert len(paras) == 2
        assert paras[0].text == "foo"
        assert paras[1].text == "bar"
        assert body.xml == xml('w:body/(w:p/w:r/w:t"foo", w:p/w:r/w:t"bar")')

    def it_can_split_itself_at_multiple_positions(self):
        body = element('w:body/w:p/w:r/w:t"foobarbaz"')
        paragraph = Paragraph(body[0], None)

        paras = paragraph.split(3, 6)

        assert [p.text for p in paras] == ["foo", "bar", "baz"]

    def it_inserts_a_paragraph_before_to_help(self, _insert_before_fixture):
        paragraph, body, expected_xml = _insert_before_fixture
        new_paragraph = paragraph._insert_paragraph_before()
        assert isinstance(new_paragraph, Paragraph)
        assert body.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("w:p", None, None, "w:p/w:r"),
            ("w:p", "foobar", None, 'w:p/w:r/w:t"foobar"'),
            ("w:p", None, "Strong", "w:p/w:r"),
            ("w:p", "foobar", "Strong", 'w:p/w:r/w:t"foobar"'),
        ]
    )
    def add_run_fixture(self, request, run_style_prop_):
        before_cxml, text, style, after_cxml = request.param
        paragraph = Paragraph(element(before_cxml), None)
        expected_xml = xml(after_cxml)
        return paragraph, text, style, run_style_prop_, expected_xml

    @pytest.fixture(
        params=[
            ("w:p/w:pPr/w:jc{w:val=center}", WD_ALIGN_PARAGRAPH.CENTER),
            ("w:p", None),
        ]
    )
    def alignment_get_fixture(self, request):
        cxml, expected_alignment_value = request.param
        paragraph = Paragraph(element(cxml), None)
        return paragraph, expected_alignment_value

    @pytest.fixture(
        params=[
            ("w:p", WD_ALIGN_PARAGRAPH.LEFT, "w:p/w:pPr/w:jc{w:val=left}"),
            (
                "w:p/w:pPr/w:jc{w:val=left}",
                WD_ALIGN_PARAGRAPH.CENTER,
                "w:p/w:pPr/w:jc{w:val=center}",
            ),
            ("w:p/w:pPr/w:jc{w:val=left}", None, "w:p/w:pPr"),
            ("w:p", None, "w:p/w:pPr"),
        ]
    )
    def alignment_set_fixture(self, request):
        initial_cxml, new_alignment_value, expected_cxml = request.param
        paragraph = Paragraph(element(initial_cxml), None)
        expected_xml = xml(expected_cxml)
        return paragraph, new_alignment_value, expected_xml

    @pytest.fixture(
        params=[
            ("w:p", "w:p"),
            ("w:p/w:pPr", "w:p/w:pPr"),
            ('w:p/w:r/w:t"foobar"', "w:p"),
            ('w:p/(w:pPr, w:r/w:t"foobar")', "w:p/w:pPr"),
        ]
    )
    def clear_fixture(self, request):
        initial_cxml, expected_cxml = request.param
        paragraph = Paragraph(element(initial_cxml), None)
        expected_xml = xml(expected_cxml)
        return paragraph, expected_xml

    @pytest.fixture(
        params=[
            (None, None),
            ("Foo", None),
            (None, "Bar"),
            ("Foo", "Bar"),
        ]
    )
    def insert_before_fixture(self, request, _insert_paragraph_before_, add_run_):
        text, style = request.param
        paragraph_ = _insert_paragraph_before_.return_value
        add_run_calls = [] if text is None else [call(text)]
        paragraph_.style = None
        return text, style, paragraph_, add_run_calls

    @pytest.fixture(params=[("w:body/w:p{id=42}", "w:body/(w:p,w:p{id=42})")])
    def _insert_before_fixture(self, request):
        body_cxml, expected_cxml = request.param
        body = element(body_cxml)
        paragraph = Paragraph(body[0], None)
        expected_xml = xml(expected_cxml)
        return paragraph, body, expected_xml

    @pytest.fixture
    def parfmt_fixture(self, ParagraphFormat_, paragraph_format_):
        paragraph = Paragraph(element("w:p"), None)
        return paragraph, ParagraphFormat_, paragraph_format_

    @pytest.fixture
    def runs_fixture(self, p_, Run_, r_, r_2_, runs_):
        paragraph = Paragraph(p_, None)
        run_, run_2_ = runs_
        return paragraph, Run_, r_, r_2_, run_, run_2_

    @pytest.fixture
    def style_get_fixture(self, part_prop_):
        style_id = "Foobar"
        p_cxml = "w:p/w:pPr/w:pStyle{w:val=%s}" % style_id
        paragraph = Paragraph(element(p_cxml), None)
        style_ = part_prop_.return_value.get_style.return_value
        return paragraph, style_id, style_

    @pytest.fixture(
        params=[
            ("w:p", "Heading 1", "Heading1", "w:p/w:pPr/w:pStyle{w:val=Heading1}"),
            (
                "w:p/w:pPr",
                "Heading 1",
                "Heading1",
                "w:p/w:pPr/w:pStyle{w:val=Heading1}",
            ),
            (
                "w:p/w:pPr/w:pStyle{w:val=Heading1}",
                "Heading 2",
                "Heading2",
                "w:p/w:pPr/w:pStyle{w:val=Heading2}",
            ),
            ("w:p/w:pPr/w:pStyle{w:val=Heading1}", "Normal", None, "w:p/w:pPr"),
            ("w:p", None, None, "w:p/w:pPr"),
        ]
    )
    def style_set_fixture(self, request, part_prop_):
        p_cxml, value, style_id, expected_cxml = request.param
        paragraph = Paragraph(element(p_cxml), None)
        part_prop_.return_value.get_style_id.return_value = style_id
        expected_xml = xml(expected_cxml)
        return paragraph, value, expected_xml

    @pytest.fixture
    def text_set_fixture(self):
        paragraph = Paragraph(element("w:p"), None)
        paragraph.add_run("must not appear in result")
        new_text_value = "foo\tbar\rbaz\n"
        expected_text_value = "foo\tbar\nbaz\n"
        return paragraph, new_text_value, expected_text_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def add_run_(self, request):
        return method_mock(request, Paragraph, "add_run")

    @pytest.fixture
    def document_part_(self, request):
        return instance_mock(request, DocumentPart)

    @pytest.fixture
    def _insert_paragraph_before_(self, request):
        return method_mock(request, Paragraph, "_insert_paragraph_before")

    @pytest.fixture
    def p_(self, request, r_, r_2_):
        return instance_mock(request, CT_P, all_runs=(r_, r_2_))

    @pytest.fixture
    def ParagraphFormat_(self, request, paragraph_format_):
        return class_mock(
            request,
            "docx.text.paragraph.ParagraphFormat",
            return_value=paragraph_format_,
        )

    @pytest.fixture
    def paragraph_format_(self, request):
        return instance_mock(request, ParagraphFormat)

    @pytest.fixture
    def part_prop_(self, request, document_part_):
        return property_mock(request, Paragraph, "part", return_value=document_part_)

    @pytest.fixture
    def Run_(self, request, runs_):
        run_, run_2_ = runs_
        return class_mock(request, "docx.text.paragraph.Run", side_effect=[run_, run_2_])

    @pytest.fixture
    def r_(self, request):
        return instance_mock(request, CT_R)

    @pytest.fixture
    def r_2_(self, request):
        return instance_mock(request, CT_R)

    @pytest.fixture
    def run_style_prop_(self, request):
        return property_mock(request, Run, "style")

    @pytest.fixture
    def runs_(self, request):
        run_ = instance_mock(request, Run, name="run_")
        run_2_ = instance_mock(request, Run, name="run_2_")
        return run_, run_2_


class DescribeParagraphFootnotesAndEndnotes:
    """Integration-test suite for `Paragraph.add_footnote()`/`.add_endnote()`.

    Uses a real |Document| rather than mocks because the behavior under test spans
    the paragraph, the containing document, and the footnotes/endnotes part.
    """

    def it_can_add_a_footnote(self):
        document = docx.Document()
        paragraph = document.add_paragraph("Some guinea pig ")

        footnote = paragraph.add_footnote()

        assert len(footnote.paragraphs) == 1
        assert paragraph.footnotes[0] is not None
        assert paragraph.footnotes[0].id == footnote.id == 1

    def it_assigns_sequential_ids_to_multiple_footnotes_in_one_paragraph(self):
        document = docx.Document()
        paragraph = document.add_paragraph("Some text")

        footnote_1 = paragraph.add_footnote()
        footnote_2 = paragraph.add_footnote()
        footnote_3 = paragraph.add_footnote()

        assert [f.id for f in paragraph.footnotes] == [1, 2, 3]
        assert (footnote_1.id, footnote_2.id, footnote_3.id) == (1, 2, 3)

    def it_renumbers_earlier_footnotes_when_one_is_inserted_ahead_of_them(self):
        document = docx.Document()
        p1 = document.add_paragraph("First paragraph")
        p1.add_footnote()
        p2 = document.add_paragraph("Second paragraph")
        p2.add_footnote()

        p0 = p1.insert_paragraph_before("Inserted first")
        p0.add_footnote()

        assert p0.footnotes[0].id == 1
        assert p1.footnotes[0].id == 2
        assert p2.footnotes[0].id == 3

    def it_can_override_the_footnote_number_format_via_the_containing_section(self):
        document = docx.Document()
        paragraph = document.add_paragraph("Some text")

        paragraph.add_footnote(num_format="lowerRoman")

        assert document.sections[0].footnote_number_format == "lowerRoman"

    def it_can_add_an_endnote(self):
        document = docx.Document()
        paragraph = document.add_paragraph("Some guinea pig ")

        endnote = paragraph.add_endnote()

        assert len(endnote.paragraphs) == 1
        assert paragraph.endnotes[0] is not None
        assert paragraph.endnotes[0].id == endnote.id == 1

    def it_can_confine_an_endnote_to_the_end_of_its_section(self):
        document = docx.Document()
        paragraph = document.add_paragraph("Some text")

        paragraph.add_endnote(section_endnote=True)

        assert document.sections[0].endnote_position == "sectEnd"
        assert document.settings.endnote_position == "sectEnd"


class DescribeParagraphNumbering:
    """Unit-test suite for `Paragraph.number` and the write-path numbering methods.

    Uses a real numbering-part element (built from cxml) behind a mocked `part`, since
    numbering resolution needs a genuine `w:numbering` tree to walk but doesn't need a
    full `Document`/package.
    """

    def it_knows_its_list_item_label(self, numbering_part_):
        body = element(
            "w:body/("
            "w:p/w:pPr/w:numPr/(w:ilvl{w:val=0},w:numId{w:val=1}),"
            "w:p/w:pPr/w:numPr/(w:ilvl{w:val=0},w:numId{w:val=1})"
            ")"
        )
        p0, p1 = body[0], body[1]
        paragraph0, paragraph1 = Paragraph(p0, None), Paragraph(p1, None)

        assert paragraph0.number == "1.\t"
        assert paragraph1.number == "2.\t"

    def it_is_None_for_a_paragraph_that_is_not_in_a_numbered_list(self):
        paragraph = Paragraph(cast(CT_P, element('w:p/w:r/w:t"foobar"')), None)
        assert paragraph.number is None

    def it_can_start_a_new_numbered_list_via_add_paragraph(self, numbering_part_):
        body = element("w:body/w:p/w:pPr/w:numPr/(w:ilvl{w:val=0},w:numId{w:val=1})")
        seed_paragraph = Paragraph(body[0], None)

        seed_paragraph.set_li_lvl(prev=None, ilvl=0)

        assert seed_paragraph.number == "1.\t"

    def it_can_join_an_existing_list_from_a_previous_paragraph(self, numbering_part_):
        body = element("w:body/(w:p/w:pPr/w:numPr/(w:ilvl{w:val=0},w:numId{w:val=1}),w:p)")
        p0, p1 = Paragraph(body[0], None), Paragraph(body[1], None)

        p1.set_li_lvl(prev=p0, ilvl=None)

        assert p0.number == "1.\t"
        assert p1.number == "2.\t"

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture(autouse=True)
    def numbering_part_(self, request):
        """Wires `Paragraph.part` to a mock whose `.numbering_part._element` is a real,
        single-decimal-list `w:numbering` tree, and whose `.cached_styles` is empty."""
        numbering_el = element(
            "w:numbering/("
            "w:abstractNum{w:abstractNumId=0}/w:lvl{w:ilvl=0}/(w:start{w:val=1},"
            "w:numFmt{w:val=decimal},w:lvlText{w:val=%1.}),"
            "w:num{w:numId=1}/w:abstractNumId{w:val=0}"
            ")"
        )
        document_part_ = instance_mock(request, DocumentPart)
        document_part_.numbering_part = SimpleNamespace(_element=numbering_el)
        document_part_.cached_styles = {}
        return property_mock(request, Paragraph, "part", return_value=document_part_)
