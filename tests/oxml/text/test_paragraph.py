# encoding: utf-8

"""
Test suite for the docx.oxml.text.paragraph module.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ...unitutil.cxml import element


class DescribeCT_P(object):

    def it_knows_the_ids_of_its_note_references(self):
        p = element(
            'w:p/(w:r/w:footnoteReference{w:id=2}'
            ',w:ins{w:id=8,w:author=A}/w:r/w:footnoteReference{w:id=3}'
            ',w:del{w:id=9,w:author=A}/w:r/w:footnoteReference{w:id=4}'
            ',w:r/w:endnoteReference{w:id=5})')

        assert p.footnote_reference_ids == [2, 3]
        assert p.endnote_reference_ids == [5]

    def it_leaves_the_note_references_of_a_text_box_to_its_paragraph(self):
        p = element(
            'w:p/(w:r/w:footnoteReference{w:id=2}'
            ',w:r/w:pict/w:txbxContent/w:p/w:r/w:footnoteReference{w:id=3})')
        p_in_text_box = p.xpath('.//w:txbxContent/w:p')[0]

        assert p.footnote_reference_ids == [2]
        assert p_in_text_box.footnote_reference_ids == [3]
