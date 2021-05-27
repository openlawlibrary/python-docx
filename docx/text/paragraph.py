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
from ..shared import Parented, lazyproperty
from ..oxml.ns import nsmap

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

    def add_sdt(self, tag_name, text='', alias_name='', temporary='false',
                placeholder_txt=None, style='Normal', bold=False, italic=False):
        """
        Adds new Structured Document Type ``w:sdt`` (Plain Text Content Control) field to the Paragraph element.
        """
        # TODO: Add support for SdtType and locked property
        from docx.sdt import SdtBase

        def apply_run_formatting(rPr, style='Normal', bold=False, italic=False, underline=False):
            if style != 'Normal':
                rStyle = rPr._add_rStyle()
                rStyle.set('{%s}val' % nsmap['w'], style)
            if bold:
                rPr._add_b()
                rPr._add_bCs()
            if italic:
                rPr._add_i()
            # TODO: impl underline

        def set_std_placeholder_text(r, text=None):
            rPr = r._add_rPr()
            rStyle = rPr._add_rStyle()
            rStyle.set('{%s}val' % nsmap['w'], 'PlaceholderText')
            rPr._add_b()
            rPr._add_bCs()
            t = r._add_t()
            placeholder_txt = text or 'Click or tap here to enter text'
            t.text = placeholder_txt
            active_placeholder = sdtPr._add_active_placeholder()
            active_placeholder.set('{%s}val' % nsmap['w'], 'true')


        sdt = self._p._new_sdt()

        sdtPr = sdt._add_sdtPr()
        alias_name = alias_name or tag_name

        # set styling on sdt lvl
        rPr = sdtPr.get_or_add_rPr()
        apply_run_formatting(rPr, style, bold, italic)

        sdtPr.name = tag_name
        sdtPr.alias_val = alias_name
        sdtPr.temp_val = temporary

        sdtContent = sdt._add_sdtContent()

        r = sdtContent._add_r()
        if not text:
            set_std_placeholder_text(r, placeholder_txt)
        else:
            # set styling on content lvl
            rPr = r.get_or_add_rPr()
            apply_run_formatting(rPr, style, bold, italic)

            t = r._add_t()
            t.text = text

        self._p.append(sdt)
        return SdtBase(sdt, self)

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
        for p in positions:
            assert 0 < p < len(self.text)
        paras = []
        splitpos = positions.pop(0)
        curpos = 0
        runidx = 0
        curpara = self
        prevtextlen = 0
        while runidx < len(curpara.runs):
            run = curpara.runs[runidx]
            endpos = curpos + len(run.text)
            if curpos <= splitpos < endpos:
                run_split_pos = splitpos - curpos
                lrun, _ = run.split(run_split_pos)
                idx_cor = 0 if lrun is None else 1
                next_para = curpara.clone()
                for crunidx, crun in enumerate(curpara.runs):
                    if crunidx >= runidx + idx_cor:
                        crun._r.getparent().remove(crun._r)
                for crunidx, crun in enumerate(next_para.runs):
                    if crunidx < runidx + idx_cor:
                        crun._r.getparent().remove(crun._r)
                curpara._p.addnext(next_para._p)
                paras.append(curpara)
                if not positions:
                    break
                curpos = splitpos
                splitpos = positions.pop(0)
                prevtextlen += len(curpara.text)
                curpara = next_para
                runidx = 0
            else:
                runidx += 1
                curpos = endpos

        paras.append(next_para)
        return paras

    def remove(self):
        """Removes this paragraph from its container."""
        self._p.getparent().remove(self._p)

    def remove_text(self, start=0, end=-1):
        """Removes part of text retaining runs and styling."""

        if end == -1:
            end = len(self.text)
        assert end > start and end <= len(self.text)

        # Check a special case
        # where both start and end fall in a single run.
        runstart = 0
        for run in self.runs:
            runend = runstart + len(run.text)
            if runstart <= start and end <= runend:
                run.text = run.text[:(start-runstart)] \
                           + run.text[(end-runstart):]
                if not run.text:
                    run._r.getparent().remove(run._r)
                return self
            runstart = runend

        # We are removing text spanning multiple runs.
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
                        runidx += 1
                    else:
                        split_pos = end - runstart
                    to_del, _ = run.split(split_pos)
                else:
                    runidx += 1
            if to_del:
                runstart = runend - len(to_del.text)
                end -= len(to_del.text)
                to_del._r.getparent().remove(to_del._r)
            else:
                runstart = runend
        return self

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
        return [Run(r, self) for r in self._p.iter_r_lst_recursive()]

    @property
    def bookmark_starts(self):
        return self._element.bookmarkStart_lst

    @property
    def bookmark_ends(self):
        return self._element.bookmarkEnd_lst

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

    def replace_char(self, oldch, newch):
        """
        Replaces all occurences of oldch character with newch.
        """
        for run in self.runs:
            run.text = run.text.replace(oldch, newch)
        return self

    def insert_text(self, position, new_text):
        """
        Inserts text at a given position.
        """
        runend = 0
        runstart = 0
        for run in self.runs:
            runstart = runend
            runend += len(run.text)
            if runend >= position:
                run.text = run.text[:(position-runstart)] \
                           + new_text + run.text[(position-runstart):]
                break
        return self

    def replace_text(self, old_text, new_text):
        """
        Replace all occurences of old_text with new_text. Keep runs formatting.
        old_text can span multiple runs.
        new_text is added to the run where old_text starts.
        """
        assert new_text
        assert old_text
        startpos = 0
        while startpos < len(self.text):
            try:
                old_start = startpos + self.text[startpos:].index(old_text)
                startpos = old_start + len(old_text)
            except ValueError:
                break

            self.remove_text(start=old_start, end=startpos)\
                .insert_text(old_start, new_text)
        return self

    def lstrip(self, chars=None):
        """
        Left strip paragraph text.
        """
        while self.runs:
            run = self.runs[0]
            run.text = run.text.lstrip(chars)
            if not run.text:
                run._r.getparent().remove(run._r)
            else:
                break
        return self

    def rstrip(self, chars=None):
        """
        Right strip paragraph text.
        """
        while self.runs:
            run = self.runs[len(self.runs) - 1]
            run.text = run.text.rstrip(chars)
            if not run.text:
                run._r.getparent().remove(run._r)
            else:
                break
        return self

    def strip(self, chars=None):
        """
        Strips paragraph text.
        """
        return self.lstrip(chars).rstrip(chars)

    @property
    def sdts(self):
        """
        Returns list of inline content controls for this paragraph.
        """
        from ..sdt import SdtBase
        return [SdtBase(sdt, self) for sdt in self._element.sdt_lst]

    def _insert_paragraph_before(self):
        """
        Return a newly created paragraph, inserted directly before this
        paragraph.
        """
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)

    def __repr__(self):
        text_stripped = self.text.strip()
        text = text_stripped[:20]
        if len(text_stripped) > len(text):
            text += '...'
        if not text:
            text = "EMPTY PARAGRAPH"
        text = '<p:"{}">'.format(text)
        return text

    def clone(self):
        """
        Cloning by selective deep copying.
        """
        c = copy.deepcopy(self)
        c._parent = self._parent
        return c

    def __getstate__(self):
        state = dict(self.__dict__)
        state.pop('_parent', None)
        return state

    def __setstate__(self, state):
        self.__dict__ = state

    @lazyproperty
    def image_parts(self):
        """
        Return all image parts related to this paragraph.
        """
        drawings = []
        for r in self.runs:
            if r._element.drawing_lst:
                drawings.extend(r._element.drawing_lst)
        blips = [drawing.xpath(".//*[local-name() = 'blip']")[0]
                 for drawing in drawings]
        rIds = [b.embed for b in blips]
        doc = self.part.document
        parts = [doc.part.related_parts[rId] for rId in rIds]
        return parts
