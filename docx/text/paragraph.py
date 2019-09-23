# encoding: utf-8

"""
Paragraph-related proxy types.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import copy
import math
from ..enum.style import WD_STYLE_TYPE
from .parfmt import ParagraphFormat
from .run import Run
from ..shared import Parented, Length, lazyproperty, Inches, cache, bust_cache
from ..oxml.ns import nsmap
from docx.bookmark import BookmarkParent


# Decorator for all text changing functions used to invalidate text cache.
text_changing = bust_cache(('text', 'run_text'))


class Paragraph(Parented, BookmarkParent):
    """
    Proxy object wrapping ``<w:p>`` element.
    """

    def __init__(self, p, parent):
        super(Paragraph, self).__init__(parent)
        self._p = self._element = p
        self._number = None
        self._lvl = None
        self._cache = {}

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

    def insert_paragraph_before(self, text=None, style=None, ilvl=None):
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
        if ilvl is not None:
            paragraph.set_li_lvl(self.part.styles, self, ilvl)
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
        self._cache = {}
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

    @text_changing
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
    def number(self):
        """
        Gets the list item number with trailing space, if paragraph is part of the numbered
        list, otherwise returns None.
        """
        if self._number is None:
            try:
                self._number = self._p.number(self.part.numbering_part._element,
                                              self.part.cached_styles)
                return self._number
            except (AttributeError, NotImplementedError):
                return None
        else:
            return self._number

    @number.setter
    def number(self, new_number):
        self._number = new_number

    @property
    def lvl(self):
        """
        Gets the `lvl` element based on the indentation index.
        """
        if self._lvl is None:
            try:
                self._lvl = self._p.lvl(self.part.numbering_part._element, self.part.cached_styles)
                return self._lvl
            except (AttributeError, NotImplementedError):
                return None
        else:
            return self._lvl

    @property
    def numbering_format(self):
        """
        Returns |ParagraphFormat| object based on the formatting for the given
        level of the numbered list.
        """
        return ParagraphFormat(self.lvl) if self.lvl is not None else None

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
    @cache
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
        self._cache = {}

    def set_li_lvl(self, styles, prev, ilvl):
        """
        Sets list indentation level for this paragraph. If ``prev`` is not specified
        it starts a new list. ``ilvl`` specifies indentation level. Default
        indentation level is 0.
        """
        prev_el = prev._element if prev else None
        _ilvl = 0 if ilvl is None else ilvl
        self._p.set_li_lvl(self.part.numbering_part._element,
                           self.part.cached_styles, prev_el, _ilvl)

    @property
    @cache
    def run_text(self):
        if self.runs:
            return ''.join(r.text for r in self.runs)
        else:
            return ''

    @property
    @cache
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
        para_num = self.number
        return (para_num if para_num is not None else '') + self.run_text

    @text.setter
    @text_changing
    def text(self, text):
        self.clear()
        self.add_run(text)

    @text_changing
    def replace_char(self, oldch, newch):
        """
        Replaces all occurences of oldch character with newch.
        """
        for run in self.runs:
            run.text = run.text.replace(oldch, newch)
        return self

    @text_changing
    def replace_chars(self, *replacement_pairs):
        """
        *replacement_pairs is tuples of (oldch, newch)

        Replaces all occurances of each replacement pair's oldch with the newch.
        """
        for run in self.runs:
            new_text = run.text
            for oldch, newch in replacement_pairs:
                new_text = new_text.replace(oldch, newch)
            run.text = new_text
        return self

    @text_changing
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

    @text_changing
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

    @text_changing
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

    @text_changing
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

    @text_changing
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

    @property
    def norm_left_indent(self):
        """
        Returns left indentation ``i`` by unifying different approaches for paragraph
        indentation like: tab characters, tab stops ``ts``, and first line indentation ``fli``.
        It takes into account user parameters ``u_*``, and inherited style parameters ``s_*``,
        where ``*`` is param name. Default tab stop ``DEFAULT_TAB_STOP`` has (.5 Inches) val.
        """

        def get_attr_with_style(obj, attr_name):
            """
            Return the first non-None attribute following the style hierarchy.
            """
            attr = getattr(obj.paragraph_format, attr_name, None)
            if attr is not None:
                return attr

            new_obj = getattr(obj, 'style', None)
            if new_obj is not None:
                return get_attr_with_style(new_obj, attr_name)

            new_obj = getattr(obj, 'base_style', None)
            if new_obj is not None:
                return get_attr_with_style(new_obj, attr_name)

            # If the attribute is not defined we end-up here and return
            # zero
            return Length(0)

        def get_tabstops(para):
            """
            Build tab-stop list from elements found following the style hierarchy.
            `clear` tab-stops lower in the hierarchy remove tab-stops from upper in
            the hierarchy.
            """
            tabstops = []

            def _inner_get_tabstops(obj):
                nonlocal tabstops
                if obj is None:
                    return
                if hasattr(obj, 'base_style'):
                    _inner_get_tabstops(obj.base_style)
                elif hasattr(obj, 'style'):
                    _inner_get_tabstops(obj.style)

                obj = obj.paragraph_format

                tabstops.extend([round(ts.position.inches, 2) for ts in obj.tab_stops])
                clear_t_stops = [round(ts.position.inches, 2)
                                 for ts in obj.tab_stops
                                 if ts._element.attrib['{%s}val' % nsmap['w']] == 'clear']
                tabstops = [ts for ts in tabstops if ts not in clear_t_stops]

            _inner_get_tabstops(para)
            return tabstops

        # Get explicitly set indentation
        first_line_indent = self.paragraph_format.first_line_indent
        if first_line_indent is not None:
            first_line_indent = round(first_line_indent.inches, 2)
        left_indent = self.paragraph_format.left_indent
        if left_indent is not None:
            left_indent = round(left_indent.inches, 2)

        if self.numbering_format:
            indent = first_line_indent if first_line_indent is not None \
                 else self.numbering_format.first_line_indent.inches
            indent += left_indent if left_indent is not None \
                else self.numbering_format.left_indent.inches
        else:
            # If para is not numbered we shall calculate using tabs and tab stops
            DEFAULT_TAB_STOP = 0.5
            tab_count = 0

            # Calculate the base first line indent and para left indent.
            left_indent = left_indent or round(get_attr_with_style(self, 'left_indent').inches, 2)
            indent = first_line_indent = \
                (first_line_indent
                 or round(get_attr_with_style(self, 'first_line_indent').inches, 2)) + left_indent

            # Find out the number of tabs at the beginning of the paragraph.
            tab_count = len(self.text) - len(self.text.lstrip('\t'))

            if tab_count:

                # Get tab stops but only those to the right of first line indent as the previous
                # don't affect the indentation.
                tab_stops = [ts for ts in get_tabstops(self) if ts > first_line_indent]

                # If the first line indent is left of the paragraph indent, first tab will tab to
                # the paragraph indent.
                if first_line_indent < left_indent:
                    tab_stops.append(left_indent)

                # Eliminate duplicates and sort.
                tab_stops = list(sorted(set(tab_stops)))

                if len(tab_stops) >= tab_count:
                    # We have enough tab stops to cover all tab chars.
                    if tab_stops:
                        indent = tab_stops[tab_count - 1]

                else:
                    if tab_stops:
                        indent = tab_stops[-1]
                        tab_count -= len(tab_stops)

                    # It's easier to calculate in whole tab stop indents instead of inches
                    indent *= (1 / DEFAULT_TAB_STOP)

                    # Let's round up to the first tab char indent. If already rounded add one.
                    tab_count -= 1
                    indent = math.ceil(indent) if not indent.is_integer() else indent + 1

                    # The remaining tab chars just adds whole indents.
                    indent += tab_count

                    # Scale back to inches
                    indent /= (1 / DEFAULT_TAB_STOP)

        return Inches(indent)

    def __repr__(self):
        text_stripped = self.text.strip()
        text = text_stripped[:20]
        if len(text_stripped) > len(text):
            text += '...'
        if not text:
            text = "EMPTY PARAGRAPH"
        text = '<p:"{}{}">'.format(
            "{} ".format(str(self.num))
            if hasattr(self, 'num') and self.num else '', text)
        return text

    def clone(self):
        """
        Cloning by selective deep copying.
        """
        c = copy.deepcopy(self)
        c._parent = self._parent
        c._cache = {}
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
