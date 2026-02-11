# encoding: utf-8

"""Unit test suite for the docx.endnotes module"""

from __future__ import absolute_import, division, print_function, unicode_literals

import pytest

from docx.endnotes import Endnote, Endnotes
from docx.oxml.endnote import CT_FtnEdn, CT_Endnotes

from .unitutil.cxml import element, xml
from .unitutil.mock import class_mock, instance_mock


class DescribeEndnotes(object):

    def it_provides_indexed_access_to_an_endnote(self, getitem_fixture):
        endnotes, reference_id, endnote_ = getitem_fixture
        endnote = endnotes[reference_id]
        assert endnote is endnote_

    def it_raises_on_indexed_access_with_bad_key(self):
        endnotes_elm = element('w:endnotes')
        endnotes = Endnotes(endnotes_elm, None)
        with pytest.raises(IndexError):
            endnotes[42]

    def it_knows_how_many_endnotes_it_contains(self, len_fixture):
        endnotes, expected_count = len_fixture
        assert len(endnotes) == expected_count

    def it_can_add_an_endnote(self, add_endnote_fixture):
        endnotes, reference_id, expected_xml = add_endnote_fixture
        endnote = endnotes.add_endnote(reference_id)
        assert endnotes._element.xml == expected_xml
        assert isinstance(endnote, Endnote)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_endnote_fixture(self):
        endnotes_cxml = 'w:endnotes/(w:endnote{w:id=-1},w:endnote{w:id=0})'
        endnotes = Endnotes(element(endnotes_cxml), None)
        reference_id = 1
        expected_xml = xml(
            'w:endnotes/(w:endnote{w:id=-1},w:endnote{w:id=0},'
            'w:endnote{w:id=1})'
        )
        return endnotes, reference_id, expected_xml

    @pytest.fixture
    def getitem_fixture(self, Endnote_, endnote_):
        endnotes_cxml = 'w:endnotes/w:endnote{w:id=1}'
        endnotes_elm = element(endnotes_cxml)
        endnotes = Endnotes(endnotes_elm, None)
        reference_id = 1
        return endnotes, reference_id, endnote_

    @pytest.fixture
    def len_fixture(self):
        endnotes_cxml = (
            'w:endnotes/(w:endnote{w:id=-1},w:endnote{w:id=0},'
            'w:endnote{w:id=1},w:endnote{w:id=2})'
        )
        endnotes = Endnotes(element(endnotes_cxml), None)
        expected_count = 4
        return endnotes, expected_count

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Endnote_(self, request, endnote_):
        return class_mock(
            request, 'docx.endnotes.Endnote',
            return_value=endnote_
        )

    @pytest.fixture
    def endnote_(self, request):
        return instance_mock(request, Endnote)


class DescribeEndnote(object):

    def it_knows_its_id(self, id_fixture):
        endnote, expected_id = id_fixture
        assert endnote.id == expected_id

    def it_can_add_a_paragraph_to_itself(self):
        endnote_cxml = 'w:endnote{w:id=1}'
        endnote = Endnote(element(endnote_cxml), None)
        p = endnote.add_paragraph('foobar')
        assert p.text == 'foobar'

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def id_fixture(self):
        endnote_cxml = 'w:endnote{w:id=42}'
        endnote = Endnote(element(endnote_cxml), None)
        expected_id = 42
        return endnote, expected_id
