# encoding: utf-8

"""Unit test suite for the docx.footnotes module"""

from __future__ import absolute_import, division, print_function, unicode_literals

import pytest

from docx.footnotes import Footnote, Footnotes
from docx.oxml.footnote import CT_FtnEnd, CT_Footnotes

from .unitutil.cxml import element, xml
from .unitutil.mock import class_mock, instance_mock


class DescribeFootnotes(object):

    def it_provides_indexed_access_to_a_footnote(self, getitem_fixture):
        footnotes, reference_id, footnote_ = getitem_fixture
        footnote = footnotes[reference_id]
        assert footnote is footnote_

    def it_raises_on_indexed_access_with_bad_key(self):
        footnotes_elm = element('w:footnotes')
        footnotes = Footnotes(footnotes_elm, None)
        with pytest.raises(IndexError):
            footnotes[42]

    def it_knows_how_many_footnotes_it_contains(self, len_fixture):
        footnotes, expected_count = len_fixture
        assert len(footnotes) == expected_count

    def it_can_add_a_footnote(self, add_footnote_fixture):
        footnotes, reference_id, expected_xml = add_footnote_fixture
        footnote = footnotes.add_footnote(reference_id)
        assert footnotes._element.xml == expected_xml
        assert isinstance(footnote, Footnote)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_footnote_fixture(self):
        footnotes_cxml = 'w:footnotes/(w:footnote{w:id=-1},w:footnote{w:id=0})'
        footnotes = Footnotes(element(footnotes_cxml), None)
        reference_id = 1
        expected_xml = xml(
            'w:footnotes/(w:footnote{w:id=-1},w:footnote{w:id=0},'
            'w:footnote{w:id=1})'
        )
        return footnotes, reference_id, expected_xml

    @pytest.fixture
    def getitem_fixture(self, Footnote_, footnote_):
        footnotes_cxml = 'w:footnotes/w:footnote{w:id=1}'
        footnotes_elm = element(footnotes_cxml)
        footnotes = Footnotes(footnotes_elm, None)
        reference_id = 1
        return footnotes, reference_id, footnote_

    @pytest.fixture
    def len_fixture(self):
        footnotes_cxml = (
            'w:footnotes/(w:footnote{w:id=-1},w:footnote{w:id=0},'
            'w:footnote{w:id=1},w:footnote{w:id=2})'
        )
        footnotes = Footnotes(element(footnotes_cxml), None)
        expected_count = 4
        return footnotes, expected_count

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Footnote_(self, request, footnote_):
        return class_mock(
            request, 'docx.footnotes.Footnote',
            return_value=footnote_
        )

    @pytest.fixture
    def footnote_(self, request):
        return instance_mock(request, Footnote)


class DescribeFootnote(object):

    def it_knows_its_id(self, id_fixture):
        footnote, expected_id = id_fixture
        assert footnote.id == expected_id

    def it_can_add_a_paragraph_to_itself(self):
        footnote_cxml = 'w:footnote{w:id=1}'
        footnote = Footnote(element(footnote_cxml), None)
        p = footnote.add_paragraph('foobar')
        assert p.text == 'foobar'

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def id_fixture(self):
        footnote_cxml = 'w:footnote{w:id=42}'
        footnote = Footnote(element(footnote_cxml), None)
        expected_id = 42
        return footnote, expected_id
