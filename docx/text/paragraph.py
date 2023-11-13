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

    def add_footnote(self):
        """
        Append a run that contains a ``<w:footnoteReferenceId>`` element.
        The footnotes are kept in order by `footnote_reference_id`, so
        the appropriate id is calculated based on the current state.
        """
        # When adding a footnote it can be inserted 
        # in front of some other footnotes, so
        # we need to sort footnotes by `footnote_reference_id`
        # in |Footnotes| and in |Paragraph|
        #
        # resolve reference ids in |Paragraph|
        new_fr_id = 1
        # If paragraph already contains footnotes
        # and it's the last paragraph with footnotes, then
        # append the new footnote and the end with the next reference id.
        if self._p.footnote_reference_ids is not None:
            new_fr_id = self._p.footnote_reference_ids[-1] + 1

        # If the document has footnotes after this paragraph,
        # the increment all footnotes pass this paragraph, 
        # and insert a new footnote at the proper position.
        document = self._parent._parent
        paragraphs = document.paragraphs
        has_passed_self = False # break the loop when we get to the footnote before the one we are inserting.
        for p_i in reversed(range(len(paragraphs))):
            if self._p is not paragraphs[p_i]._p:
                if paragraphs[p_i]._p.footnote_reference_ids is not None:
                    if not has_passed_self:
                        for r in paragraphs[p_i].runs:
                            r._r.increment_footnote_reference_id()
                    else:
                        new_fr_id = max(paragraphs[p_i]._p.footnote_reference_ids)+1
                        break
            else:
                has_passed_self = True
        r = self._p.add_r()
        r.add_footnoteReference(new_fr_id)
        footnote = document._add_footnote(new_fr_id)
        return footnote

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

    @property
    def footnotes(self):
        """
        Returns a list of |Footnote| instances that refers to the footnotes in this paragraph,
        or |None| if none footnote is defined.
        """
        reference_ids = self._p.footnote_reference_ids
        if reference_ids == None:
            return None
        footnotes = self._parent._parent.footnotes
        footnote_list = []
        for ref_id in reference_ids:
            footnote_list.append(footnotes[ref_id])
        return footnote_list

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
