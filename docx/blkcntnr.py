# encoding: utf-8

"""Block item container, used by body, cell, header, etc.

Block level items are things like paragraph and table, although there are a few other
specialized ones like structured document tags.
"""

from __future__ import absolute_import, print_function
from collections import OrderedDict

from .oxml.table import CT_Tbl
from .shared import Parented
from docx.bookmark import BookmarkParent
from .text.paragraph import Paragraph

class BlockItemContainer(Parented, BookmarkParent):
    """
    Base class for proxy objects that can contain block items, such as _Body,
    _Cell, header, footer, footnote, endnote, comment, and text box objects.
    Provides the shared functionality to add a block item like a paragraph or
    table.
    """
    def __init__(self, element, parent):
        super(BlockItemContainer, self).__init__(parent)
        self._element = element

    def add_paragraph(self, text='', style=None, prev=None, ilvl=None):
        """
        Return a paragraph newly added to the end of the content in this
        container, having *text* in a single run if present, and having
        paragraph style *style*. If *style* is |None|, no paragraph style is
        applied, which has the same effect as applying the 'Normal' style.
        If paragraph is part of numbered list then ``prev_p`` (previous para)
        and ``ilvl``(indentation level) should be specified.
        """
        paragraph = self._add_paragraph()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        if prev is not None or ilvl is not None:
            paragraph.set_li_lvl(self.part.styles, prev, ilvl)
        return paragraph

    def add_table(self, rows, cols, width):
        """
        Return a table of *width* having *rows* rows and *cols* columns,
        newly appended to the content in this container. *width* is evenly
        distributed between the table columns.
        """
        from .table import Table
        tbl = CT_Tbl.new_tbl(rows, cols, width)
        self._element._insert_tbl(tbl)
        return Table(tbl, self)

    def add_sdt(self, tag_name, alias_name=''):
        """
        Returns Rich Text Content Control with given *tag_name*.
        Appends created content control to the content in this container.
        """
        from .sdt import SdtBase
        sdt = self._element._new_sdt()

        sdtPr = sdt._add_sdtPr()
        sdtPr.name = tag_name
        alias_name = alias_name or tag_name
        sdtPr.alias_val = alias_name

        sdt._add_sdtContent()
        self._element.append(sdt)
        return SdtBase(sdt, self)

    @property
    def paragraphs(self):
        """
        A list containing the paragraphs in this container, in document
        order. Read-only.
        """
        return [Paragraph(p, self) for p in self._element.p_lst]

    @property
    def sdts(self):
        """
        A list of children sdts (content controls) in this container, in
        document order. Read-only.
        """
        from .sdt import SdtBase
        return OrderedDict({k:SdtBase(s, self) for (s,k) in self._iter_sdts()})

    @property
    def sdts_all(self):
        """
        A list of descendants sdts (content controls) in this container, in
        document order. Read-only.
        """
        from .sdt import SdtBase
        return OrderedDict({k:SdtBase(s, self) for (s,k) in self._iter_sdts_all()})

    @property
    def tables(self):
        """
        A list containing the tables in this container, in document order.
        Read-only.
        """
        from .table import Table
        return [Table(tbl, self) for tbl in self._element.tbl_lst]

    def _add_paragraph(self):
        """
        Return a paragraph newly added to the end of the content in this
        container.
        """
        return Paragraph(self._element.add_p(), self)

    def _iter_sdts(self):
        for sdt in self._element.sdt_lst:
            yield sdt, sdt.name

    def _iter_sdts_all(self):
        nsmap = self._element.nsmap
        sections = self._parent.sections
        for s in sections:
            hdr_ftrs = (s.header, s.first_page_header, s.even_page_header,
                        s.footer, s.first_page_footer, s.even_page_footer)
            for hdr_ftr in hdr_ftrs:
                for sdt in hdr_ftr._element.iterdescendants('{%s}sdt' % nsmap['w']):
                    yield sdt, sdt.name
        for sdt in self._element.iterdescendants('{%s}sdt' % nsmap['w']):
            yield sdt, sdt.name
