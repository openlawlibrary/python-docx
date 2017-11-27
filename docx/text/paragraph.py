# encoding: utf-8

"""
Paragraph-related proxy types.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import copy
from ..enum.style import WD_STYLE_TYPE
from .parfmt import ParagraphFormat
from .run import Run
from ..shared import Parented


class Paragraph(Parented):
    """
    Proxy object wrapping ``<w:p>`` element.
    """
    def __init__(self, p, parent):
        super(Paragraph, self).__init__(parent)
        self._p = self._element = p

    def add_run(self, text=None, style=None):
        """
        Append a run to this paragraph containing *text* and having character
        style identified by style ID *style*. *text* can contain tab
        (``\\t``) characters, which are converted to the appropriate XML form
        for a tab. *text* can also include newline (``\\n``) or carriage
        return (``\\r``) characters, each of which is converted to a line
        break.
        """
        r = self._p.add_r()
        run = Run(r, self)
        if text:
            run.text = text
        if style:
            run.style = style
        return run

    @property
    def alignment(self):
        """
        A member of the :ref:`WdParagraphAlignment` enumeration specifying
        the justification setting for this paragraph. A value of |None|
        indicates the paragraph has no directly-applied alignment value and
        will inherit its alignment value from its style hierarchy. Assigning
        |None| to this property removes any directly-applied alignment value.
        """
        return self._p.alignment

    @alignment.setter
    def alignment(self, value):
        self._p.alignment = value

    def clear(self):
        """
        Return this same paragraph after removing all its content.
        Paragraph-level formatting, such as style, is preserved.
        """
        self._p.clear_content()
        return self

    def insert_paragraph_before(self, text=None, style=None):
        """
        Return a newly created paragraph, inserted directly before this
        paragraph. If *text* is supplied, the new paragraph contains that
        text in a single run. If *style* is provided, that style is assigned
        to the new paragraph.
        """
        paragraph = self._insert_paragraph_before()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    def split(self, *positions):
        """Splits paragraph at given positions keeping formatting.

        Original unsplitted runs are retained. Original paragraph is kept but
        the next runs are deleted. New paragraphs are created to follow with
        the rest of the runs. Split is done in-place, i.e. this paragraph will
        be replaced by splitted ones.

        Returns: new splitted paragraphs.

        """
        positions = list(positions)
        paras = [self]
        splitpos = positions.pop(0)
        curpos = 0
        runidx = 0
        curpara = self
        while runidx < len(curpara.runs):
            run = curpara.runs[runidx]
            endpos = curpos + len(run.text)
            while curpos <= splitpos < endpos:
                run_split_pos = splitpos - curpos
                run.split(run_split_pos)
                next_para = copy.deepcopy(curpara)
                for crunidx, crun in enumerate(curpara.runs):
                    if crunidx > runidx:
                        crun._r.getparent().remove(crun._r)
                for crunidx, crun in enumerate(next_para.runs):
                    if crunidx <= runidx:
                        crun._r.getparent().remove(crun._r)
                curpara._p.addnext(next_para._p)
                paras.append(next_para)
                if not positions:
                    break
                splitpos = positions.pop(0)
                curpara = next_para
                runidx = -1
                run = curpara.runs[0]

            runidx += 1
            curpos = endpos

        return paras

    def remove(self):
        """Removes this paragraph from its container."""
        self._p.getparent().remove(self._p)

    def remove_text(self, start=0, end=-1):
        """Removes part of text retaining runs and styling."""

        if end == -1:
            end = len(self.text)
        assert end > start and end <= len(self.text)
        runstart = 0
        runidx = 0
        while runidx < len(self.runs) and end > start:
            run = self.runs[runidx]
            runend = runstart + len(run.text)
            to_del = None
            if start <= runstart and runend <= end:
                to_del = run
            else:
                if runstart <= start < runend:
                    _, to_del = run.split(start - runstart)
                if runstart < end <= runend:
                    if to_del:
                        run = to_del
                        split_pos = end - start
                    else:
                        split_pos = end - runstart
                    to_del, _ = run.split(split_pos)
            if to_del:
                runstart = runend - len(to_del.text)
                end -= len(to_del.text)
                to_del._r.getparent().remove(to_del._r)
            else:
                runstart = runend
            runidx += 1

    @property
    def paragraph_format(self):
        """
        The |ParagraphFormat| object providing access to the formatting
        properties for this paragraph, such as line spacing and indentation.
        """
        return ParagraphFormat(self._element)

    @property
    def runs(self):
        """
        Sequence of |Run| instances corresponding to the <w:r> elements in
        this paragraph.
        """
        return [Run(r, self) for r in self._p.r_lst]

    @property
    def style(self):
        """
        Read/Write. |_ParagraphStyle| object representing the style assigned
        to this paragraph. If no explicit style is assigned to this
        paragraph, its value is the default paragraph style for the document.
        A paragraph style name can be assigned in lieu of a paragraph style
        object. Assigning |None| removes any applied style, making its
        effective value the default paragraph style for the document.
        """
        style_id = self._p.style
        return self.part.get_style(style_id, WD_STYLE_TYPE.PARAGRAPH)

    @style.setter
    def style(self, style_or_name):
        style_id = self.part.get_style_id(
            style_or_name, WD_STYLE_TYPE.PARAGRAPH
        )
        self._p.style = style_id

    @property
    def text(self):
        """
        String formed by concatenating the text of each run in the paragraph.
        Tabs and line breaks in the XML are mapped to ``\\t`` and ``\\n``
        characters respectively.

        Assigning text to this property causes all existing paragraph content
        to be replaced with a single run containing the assigned text.
        A ``\\t`` character in the text is mapped to a ``<w:tab/>`` element
        and each ``\\n`` or ``\\r`` character is mapped to a line break.
        Paragraph-level formatting, such as style, is preserved. All
        run-level formatting, such as bold or italic, is removed.
        """
        text = ''
        for run in self.runs:
            text += run.text
        return text

    @text.setter
    def text(self, text):
        self.clear()
        self.add_run(text)

    def _insert_paragraph_before(self):
        """
        Return a newly created paragraph, inserted directly before this
        paragraph.
        """
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)
