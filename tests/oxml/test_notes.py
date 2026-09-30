# encoding: utf-8

"""
Test suite for the reference marks of notes in ``docx.oxml.notes``.
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.notes import format_note_number, note_reference_marks
from docx.oxml.ns import nsdecls

from ..unitutil.cxml import element

FN = 'w:p/w:r/w:footnoteReference{w:id=%d}'
EN = 'w:p/w:r/w:endnoteReference{w:id=%d}'


def body(*paragraphs):
    return element('w:body/(%s)' % ','.join(paragraphs))


class Describe_note_reference_marks(object):

    def it_numbers_references_in_order_not_by_id(self):
        cut = note_reference_marks(body(FN % 3, FN % 2), None)

        assert cut == ({3: '1', 2: '2'}, {})

    def it_numbers_footnotes_and_endnotes_apart(self):
        cut = note_reference_marks(body(FN % 2, EN % 2, FN % 3, EN % 3), None)

        assert cut == ({2: '1', 3: '2'}, {2: 'i', 3: 'ii'})

    def it_uses_the_format_and_start_of_the_document(self):
        settings = element(
            'w:settings/w:endnotePr/(w:numFmt{w:val=lowerRoman},w:numStart{w:val=19})')

        cut = note_reference_marks(body(EN % 2, EN % 3), settings)

        assert cut == ({}, {2: 'xix', 3: 'xx'})

    def it_uses_the_format_of_the_section_that_holds_the_reference(self):
        # A w:sectPr in a paragraph ends its section: the first reference is
        # in the upper roman section, the second in the last section, which
        # has the letters of the document. The counter continues.
        settings = element('w:settings/w:footnotePr/w:numFmt{w:val=upperLetter}')
        sectPr = 'w:sectPr/w:footnotePr/w:numFmt{w:val=upperRoman}'

        cut = note_reference_marks(body(FN % 2, 'w:p/w:pPr/' + sectPr, FN % 3), settings)

        assert cut == ({2: 'I', 3: 'B'}, {})

    def it_restarts_the_numbers_for_a_section_that_restarts(self):
        restart = 'w:p/w:pPr/w:sectPr/w:footnotePr/w:numRestart{w:val=eachSect}'
        last = 'w:sectPr/w:footnotePr/w:numRestart{w:val=eachSect}'

        cut = note_reference_marks(body(FN % 2, FN % 3, restart, FN % 4, last), None)

        assert cut == ({2: '1', 3: '2', 4: '1'}, {})

    def it_shows_a_custom_mark_without_a_number(self):
        custom = 'w:p/w:r/(w:footnoteReference{w:id=2,w:customMarkFollows=1},w:t"*")'

        cut = note_reference_marks(body(custom, FN % 3), None)

        assert cut == ({2: '*', 3: '1'}, {})

    def it_skips_a_deleted_reference(self):
        deleted = 'w:p/w:del{w:id=9,w:author=A}/w:r/w:footnoteReference{w:id=2}'

        cut = note_reference_marks(body(deleted, FN % 3), None)

        assert cut == ({3: '1'}, {})


class Describe_format_note_number(object):

    @pytest.mark.parametrize('number, number_format, expected', [
        (1, 'decimal', '1'),
        (19, 'lowerRoman', 'xix'),
        (4, 'upperRoman', 'IV'),
        (28, 'lowerLetter', 'bb'),
        (3, 'upperLetter', 'C'),
        (5, 'chicago', '**'),
        (4, 'chicago', '§'),
        (7, 'ordinalText', '7'),
    ])
    def it_formats_the_number_as_word_does(self, number, number_format, expected):
        assert format_note_number(number, number_format) == expected


class DescribeDocument_reference_marks(object):

    def it_knows_the_marks_of_its_notes(self):
        document = Document()
        paragraph = document.add_paragraph('Text')
        for kind, note_id in (('footnote', 2), ('endnote', 2), ('footnote', 3)):
            paragraph._p.append(parse_xml(
                '<w:r %s><w:%sReference w:id="%d"/></w:r>' % (nsdecls('w'), kind, note_id)))

        assert document.footnote_reference_marks == {2: '1', 3: '2'}
        assert document.endnote_reference_marks == {2: 'i'}
