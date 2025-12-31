# encoding: utf-8

"""Unit test suite for the docx.text.paragraph module"""

from __future__ import absolute_import, division, print_function, unicode_literals

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.run import CT_R
from docx.oxml.text.font import CT_RPr
from docx.parts.document import DocumentPart
from docx.text.paragraph import Paragraph
from docx.text.parfmt import ParagraphFormat
from docx.text.run import Run

import pytest

from ..unitutil.cxml import element, xml
from ..unitutil.mock import (
    call, class_mock, instance_mock, method_mock, property_mock
)


class DescribeParagraph(object):

    def it_knows_its_paragraph_style(self, style_get_fixture):
        paragraph, style_id_, style_ = style_get_fixture
        style = paragraph.style
        paragraph.part.get_style.assert_called_once_with(
            style_id_, WD_STYLE_TYPE.PARAGRAPH
        )
        assert style is style_

    def it_can_change_its_paragraph_style(self, style_set_fixture):
        paragraph, value, expected_xml = style_set_fixture

        paragraph.style = value

        paragraph.part.get_style_id.assert_called_once_with(
            value, WD_STYLE_TYPE.PARAGRAPH
        )
        assert paragraph._p.xml == expected_xml

    def it_knows_the_text_it_contains(self, text_get_fixture):
        paragraph, expected_text = text_get_fixture
        assert paragraph.text == expected_text

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
        assert Run_.mock_calls == [
            call(r_, paragraph), call(r_2_, paragraph)
        ]
        assert runs == [run_, run_2_]

    def it_can_add_a_run_to_itself(self, add_run_fixture):
        paragraph, text, style, style_prop_, expected_xml = add_run_fixture
        run = paragraph.add_run(text, style)
        assert paragraph._p.xml == expected_xml
        assert isinstance(run, Run)
        assert run._r is paragraph._p.r_lst[0]
        if style:
            style_prop_.assert_called_once_with(style)

    def it_can_add_a_footnote(self, add_footnote_fixture):
        paragraph, document_, footnote_, num_format = add_footnote_fixture
        footnote = paragraph.add_footnote(num_format=num_format)
        document_._calculate_next_footnote_reference_id.assert_called_once_with(
            paragraph._p
        )
        assert footnote is footnote_

    def it_can_add_a_footnote_with_auto_paragraph(self, add_footnote_auto_para_fixture):
        paragraph, document_, footnote_, p_ = add_footnote_auto_para_fixture
        footnote = paragraph.add_footnote(auto_paragraph=True)
        # Verify that a paragraph was added to the footnote
        footnote_.add_paragraph.assert_called_once_with()
        # Verify that the paragraph style was set
        assert p_.style == 'FootnoteText'

    def it_can_add_an_endnote(self, add_endnote_fixture):
        paragraph, document_, endnote_, section_endnote, num_format = (
            add_endnote_fixture
        )
        endnote = paragraph.add_endnote(
            section_endnote=section_endnote, num_format=num_format
        )
        document_._calculate_next_endnote_reference_id.assert_called_once_with(
            paragraph._p
        )
        assert endnote is endnote_

    def it_can_add_an_endnote_with_auto_paragraph(self, add_endnote_auto_para_fixture):
        paragraph, document_, endnote_, p_ = add_endnote_auto_para_fixture
        endnote = paragraph.add_endnote(auto_paragraph=True)
        # Verify that a paragraph was added to the endnote
        endnote_.add_paragraph.assert_called_once_with()
        # Verify that the paragraph style was set
        assert p_.style == 'EndnoteText'

    def it_can_insert_a_paragraph_before_itself(self, insert_before_fixture):
        text, style, paragraph_, add_run_calls = insert_before_fixture
        paragraph = Paragraph(None, None)

        new_paragraph = paragraph.insert_paragraph_before(text, style)

        paragraph._insert_paragraph_before.assert_called_once_with(paragraph)
        assert new_paragraph.add_run.call_args_list == add_run_calls
        assert new_paragraph.style == style
        assert new_paragraph is paragraph_

    def it_can_remove_its_content_while_preserving_formatting(
            self, clear_fixture):
        paragraph, expected_xml = clear_fixture
        _paragraph = paragraph.clear()
        assert paragraph._p.xml == expected_xml
        assert _paragraph is paragraph

    def it_inserts_a_paragraph_before_to_help(self, _insert_before_fixture):
        paragraph, body, expected_xml = _insert_before_fixture
        new_paragraph = paragraph._insert_paragraph_before()
        assert isinstance(new_paragraph, Paragraph)
        assert body.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('w:p', None,     None,     'w:p/w:r'),
        ('w:p', 'foobar', None,     'w:p/w:r/w:t"foobar"'),
        ('w:p', None,     'Strong', 'w:p/w:r'),
        ('w:p', 'foobar', 'Strong', 'w:p/w:r/w:t"foobar"'),
    ])
    def add_run_fixture(self, request, run_style_prop_):
        before_cxml, text, style, after_cxml = request.param
        paragraph = Paragraph(element(before_cxml), None)
        expected_xml = xml(after_cxml)
        return paragraph, text, style, run_style_prop_, expected_xml

    @pytest.fixture(params=[
        ('w:p/w:pPr/w:jc{w:val=center}', WD_ALIGN_PARAGRAPH.CENTER),
        ('w:p', None),
    ])
    def alignment_get_fixture(self, request):
        cxml, expected_alignment_value = request.param
        paragraph = Paragraph(element(cxml), None)
        return paragraph, expected_alignment_value

    @pytest.fixture(params=[
        ('w:p', WD_ALIGN_PARAGRAPH.LEFT,
         'w:p/w:pPr/w:jc{w:val=left}'),
        ('w:p/w:pPr/w:jc{w:val=left}', WD_ALIGN_PARAGRAPH.CENTER,
         'w:p/w:pPr/w:jc{w:val=center}'),
        ('w:p/w:pPr/w:jc{w:val=left}', None,
         'w:p/w:pPr'),
        ('w:p', None, 'w:p/w:pPr'),
    ])
    def alignment_set_fixture(self, request):
        initial_cxml, new_alignment_value, expected_cxml = request.param
        paragraph = Paragraph(element(initial_cxml), None)
        expected_xml = xml(expected_cxml)
        return paragraph, new_alignment_value, expected_xml

    @pytest.fixture(params=[
        ('w:p', 'w:p'),
        ('w:p/w:pPr', 'w:p/w:pPr'),
        ('w:p/w:r/w:t"foobar"', 'w:p'),
        ('w:p/(w:pPr, w:r/w:t"foobar")', 'w:p/w:pPr'),
    ])
    def clear_fixture(self, request):
        initial_cxml, expected_cxml = request.param
        paragraph = Paragraph(element(initial_cxml), None)
        expected_xml = xml(expected_cxml)
        return paragraph, expected_xml

    @pytest.fixture(params=[
        (None,  None),
        ('Foo', None),
        (None,  'Bar'),
        ('Foo', 'Bar'),
    ])
    def insert_before_fixture(self, request, _insert_paragraph_before_, add_run_):
        text, style = request.param
        paragraph_ = _insert_paragraph_before_.return_value
        add_run_calls = [] if text is None else [call(text)]
        paragraph_.style = None
        return text, style, paragraph_, add_run_calls

    @pytest.fixture(params=[
        ('w:body/w:p{id=42}', 'w:body/(w:p,w:p{id=42})')
    ])
    def _insert_before_fixture(self, request):
        body_cxml, expected_cxml = request.param
        body = element(body_cxml)
        paragraph = Paragraph(body[0], None)
        expected_xml = xml(expected_cxml)
        return paragraph, body, expected_xml

    @pytest.fixture
    def parfmt_fixture(self, ParagraphFormat_, paragraph_format_):
        paragraph = Paragraph(element('w:p'), None)
        return paragraph, ParagraphFormat_, paragraph_format_

    @pytest.fixture
    def runs_fixture(self, p_, Run_, r_, r_2_, runs_):
        paragraph = Paragraph(p_, None)
        run_, run_2_ = runs_
        return paragraph, Run_, r_, r_2_, run_, run_2_

    @pytest.fixture
    def style_get_fixture(self, part_prop_):
        style_id = 'Foobar'
        p_cxml = 'w:p/w:pPr/w:pStyle{w:val=%s}' % style_id
        paragraph = Paragraph(element(p_cxml), None)
        style_ = part_prop_.return_value.get_style.return_value
        return paragraph, style_id, style_

    @pytest.fixture(params=[
        ('w:p',                                 'Heading 1', 'Heading1',
         'w:p/w:pPr/w:pStyle{w:val=Heading1}'),
        ('w:p/w:pPr',                           'Heading 1', 'Heading1',
         'w:p/w:pPr/w:pStyle{w:val=Heading1}'),
        ('w:p/w:pPr/w:pStyle{w:val=Heading1}',  'Heading 2', 'Heading2',
         'w:p/w:pPr/w:pStyle{w:val=Heading2}'),
        ('w:p/w:pPr/w:pStyle{w:val=Heading1}',  'Normal',    None,
         'w:p/w:pPr'),
        ('w:p',                                 None,        None,
         'w:p/w:pPr'),
    ])
    def style_set_fixture(self, request, part_prop_):
        p_cxml, value, style_id, expected_cxml = request.param
        paragraph = Paragraph(element(p_cxml), None)
        part_prop_.return_value.get_style_id.return_value = style_id
        expected_xml = xml(expected_cxml)
        return paragraph, value, expected_xml

    @pytest.fixture(params=[
        ('w:p', ''),
        ('w:p/w:r', ''),
        ('w:p/w:r/w:t', ''),
        ('w:p/w:r/w:t"foo"', 'foo'),
        ('w:p/w:r/(w:t"foo", w:t"bar")', 'foobar'),
        ('w:p/w:r/(w:t"fo ", w:t"bar")', 'fo bar'),
        ('w:p/w:r/(w:t"foo", w:tab, w:t"bar")', 'foo\tbar'),
        ('w:p/w:r/(w:t"foo", w:br,  w:t"bar")', 'foo\nbar'),
        ('w:p/w:r/(w:t"foo", w:cr,  w:t"bar")', 'foo\nbar'),
    ])
    def text_get_fixture(self, request):
        p_cxml, expected_text_value = request.param
        paragraph = Paragraph(element(p_cxml), None)
        return paragraph, expected_text_value

    @pytest.fixture
    def text_set_fixture(self):
        paragraph = Paragraph(element('w:p'), None)
        paragraph.add_run('must not appear in result')
        new_text_value = 'foo\tbar\rbaz\n'
        expected_text_value = 'foo\tbar\nbaz\n'
        return paragraph, new_text_value, expected_text_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def add_run_(self, request):
        return method_mock(request, Paragraph, 'add_run')

    @pytest.fixture
    def document_part_(self, request):
        return instance_mock(request, DocumentPart)

    @pytest.fixture
    def _insert_paragraph_before_(self, request):
        return method_mock(request, Paragraph, '_insert_paragraph_before')

    @pytest.fixture
    def p_(self, request, r_, r_2_):
        return instance_mock(request, CT_P, r_lst=(r_, r_2_))

    @pytest.fixture
    def ParagraphFormat_(self, request, paragraph_format_):
        return class_mock(
            request, 'docx.text.paragraph.ParagraphFormat',
            return_value=paragraph_format_
        )

    @pytest.fixture
    def paragraph_format_(self, request):
        return instance_mock(request, ParagraphFormat)

    @pytest.fixture
    def part_prop_(self, request, document_part_):
        return property_mock(
            request, Paragraph, 'part', return_value=document_part_
        )

    @pytest.fixture
    def Run_(self, request, runs_):
        run_, run_2_ = runs_
        return class_mock(
            request, 'docx.text.paragraph.Run', side_effect=[run_, run_2_]
        )

    @pytest.fixture
    def r_(self, request):
        return instance_mock(request, CT_R)

    @pytest.fixture
    def r_2_(self, request):
        return instance_mock(request, CT_R)

    @pytest.fixture
    def run_style_prop_(self, request):
        return property_mock(request, Run, 'style')

    @pytest.fixture
    def runs_(self, request):
        run_ = instance_mock(request, Run, name='run_')
        run_2_ = instance_mock(request, Run, name='run_2_')
        return run_, run_2_

    # footnote and endnote fixtures ---------------------------------

    @pytest.fixture
    def document_(self, request):
        from docx.document import Document
        return instance_mock(request, Document)

    @pytest.fixture(params=[None, 'upperRoman'])
    def add_footnote_fixture(self, request, document_, footnote_, p_):
        num_format = request.param
        paragraph = Paragraph(p_, None)
        paragraph._parent = document_
        document_._calculate_next_footnote_reference_id.return_value = 1
        document_._add_footnote.return_value = footnote_
        p_.add_r.return_value = instance_mock(request, CT_R)
        return paragraph, document_, footnote_, num_format

    @pytest.fixture
    def add_footnote_auto_para_fixture(self, request, document_, footnote_, p_):
        from docx.text.paragraph import Paragraph as ParaProxy
        paragraph = Paragraph(p_, None)
        paragraph._parent = document_
        document_._calculate_next_footnote_reference_id.return_value = 1
        document_._add_footnote.return_value = footnote_
        p_.add_r.return_value = instance_mock(request, CT_R)

        # Mock the paragraph added to footnote
        footnote_para_ = instance_mock(request, ParaProxy)
        footnote_para_p_ = instance_mock(request, CT_P)
        footnote_para_r_ = instance_mock(request, CT_R)
        footnote_para_rPr_ = instance_mock(request, CT_RPr)

        footnote_.add_paragraph.return_value = footnote_para_
        object.__setattr__(footnote_para_, '_p', footnote_para_p_)
        footnote_para_p_.add_r.return_value = footnote_para_r_
        footnote_para_r_.get_or_add_rPr.return_value = footnote_para_rPr_

        return paragraph, document_, footnote_, footnote_para_

    @pytest.fixture(params=[
        (False, None),
        (True, 'lowerLetter'),
        (False, 'upperRoman')
    ])
    def add_endnote_fixture(self, request, document_, endnote_, p_, section_):
        section_endnote, num_format = request.param
        paragraph = Paragraph(p_, None)
        paragraph._parent = document_
        document_._calculate_next_endnote_reference_id.return_value = 1
        document_._add_endnote.return_value = endnote_
        document_.sections = [section_]
        p_.add_r.return_value = instance_mock(request, CT_R)
        return paragraph, document_, endnote_, section_endnote, num_format

    @pytest.fixture
    def add_endnote_auto_para_fixture(self, request, document_, endnote_, p_):
        from docx.text.paragraph import Paragraph as ParaProxy
        paragraph = Paragraph(p_, None)
        paragraph._parent = document_
        document_._calculate_next_endnote_reference_id.return_value = 1
        document_._add_endnote.return_value = endnote_
        p_.add_r.return_value = instance_mock(request, CT_R)

        # Mock the paragraph added to endnote
        endnote_para_ = instance_mock(request, ParaProxy)
        endnote_para_p_ = instance_mock(request, CT_P)
        endnote_para_r_ = instance_mock(request, CT_R)
        endnote_para_rPr_ = instance_mock(request, CT_RPr)

        endnote_.add_paragraph.return_value = endnote_para_
        object.__setattr__(endnote_para_, '_p', endnote_para_p_)
        endnote_para_p_.add_r.return_value = endnote_para_r_
        endnote_para_r_.get_or_add_rPr.return_value = endnote_para_rPr_

        return paragraph, document_, endnote_, endnote_para_

    @pytest.fixture
    def footnote_(self, request):
        from docx.footnotes import Footnote
        return instance_mock(request, Footnote)

    @pytest.fixture
    def endnote_(self, request):
        from docx.endnotes import Endnote
        return instance_mock(request, Endnote)

    @pytest.fixture
    def section_(self, request):
        from docx.section import Section
        from docx.oxml.section import CT_SectPr
        sectPr = instance_mock(request, CT_SectPr)
        section = instance_mock(request, Section)
        object.__setattr__(section, '_sectPr', sectPr)
        return section
