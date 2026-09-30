# encoding: utf-8

"""
Test suite for the list labels computed in ``docx.oxml.numbering``.
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls


class DescribeCT_Numbering(object):

    def it_computes_multi_level_labels(self, labels_fixture):
        document, expected_labels = labels_fixture
        assert [p.number for p in document.paragraphs] == expected_labels

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        # Word shows %1 as nothing when level 1 has numFmt="none"
        ('none', '%12030.%2', ['2030.1\t', '2030.2\t']),
        ('decimal', '%1.%2', ['2030.1\t', '2030.2\t']),
    ])
    def labels_fixture(self, request):
        lvl0_numFmt, lvl1_text, expected_labels = request.param
        document = Document()
        numbering = document.part.numbering_part.element
        numbering.append(parse_xml(
            '<w:abstractNum %s w:abstractNumId="90">'
            '  <w:lvl w:ilvl="0"><w:start w:val="2030"/>'
            '    <w:numFmt w:val="%s"/><w:lvlText w:val="2030"/></w:lvl>'
            '  <w:lvl w:ilvl="1"><w:start w:val="1"/>'
            '    <w:numFmt w:val="decimal"/><w:lvlText w:val="%s"/></w:lvl>'
            '</w:abstractNum>' % (nsdecls('w'), lvl0_numFmt, lvl1_text)
        ))
        numbering.append(parse_xml(
            '<w:num %s w:numId="90"><w:abstractNumId w:val="90"/></w:num>'
            % nsdecls('w')
        ))
        for text in ('First.', 'Second.'):
            paragraph = document.add_paragraph(text)
            paragraph._p.get_or_add_pPr().append(parse_xml(
                '<w:numPr %s><w:ilvl w:val="1"/><w:numId w:val="90"/></w:numPr>'
                % nsdecls('w')
            ))
        return document, expected_labels
