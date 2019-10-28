# encoding: utf-8

"""
Paragraph-related proxy types.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

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
        self._number = None
        self._lvl = None

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
        text = para_num if para_num is not None else ''
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
