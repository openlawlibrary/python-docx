"""Paragraph-related proxy types."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, List, cast

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.text.run import CT_R
from docx.shared import StoryChild, find_containing_document
from docx.styles.style import ParagraphStyle
from docx.text.hyperlink import Hyperlink
from docx.text.pagebreak import RenderedPageBreak
from docx.text.parfmt import ParagraphFormat
from docx.text.run import Run

if TYPE_CHECKING:
    import docx.types as t
    from docx.document import Document
    from docx.endnotes import Endnote
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.footnotes import Footnote
    from docx.oxml.section import CT_SectPr
    from docx.oxml.text.paragraph import CT_P
    from docx.section import Section
    from docx.styles.style import CharacterStyle


class Paragraph(StoryChild):
    """Proxy object wrapping a `<w:p>` element."""

    def __init__(self, p: CT_P, parent: t.ProvidesStoryPart):
        super(Paragraph, self).__init__(parent)
        self._p = self._element = p

    def add_endnote(
        self,
        section_endnote: bool = False,
        num_format: str | None = None,
        auto_paragraph: bool = True,
    ) -> Endnote:
        """Append an endnote reference to this paragraph and return the newly created
        |Endnote|.

        The endnote is inserted into |Endnotes| at the position that keeps endnotes
        ordered by reference id, renumbering any endnotes that come after it in the
        document as needed.

        When `auto_paragraph` is |True| (the default), the endnote is given an initial
        paragraph containing the endnote reference mark, matching how Word creates a
        new endnote. Use `endnote.add_paragraph()` to add the endnote's actual content.

        When `section_endnote` is |True|, or `num_format` is provided, the section
        containing this paragraph has its endnote-position and/or endnote-numbering-
        format overridden accordingly; see :attr:`.Section.endnote_position` and
        :attr:`.Section.endnote_number_format`.
        """
        document = find_containing_document(self)
        new_er_id = document._calculate_next_endnote_reference_id(  # pyright: ignore[reportPrivateUsage]
            self._p
        )
        r = self._p.add_r()
        r.add_endnoteReference(new_er_id)
        endnote = document._add_endnote(new_er_id)  # pyright: ignore[reportPrivateUsage]

        if auto_paragraph:
            p = endnote.add_paragraph()
            p._p.style = "EndnoteText"
            ref_r = p._p.add_r()
            ref_rPr = ref_r.get_or_add_rPr()
            ref_rPr.style = "EndnoteReference"
            ref_r.add_endnoteRef()

        if section_endnote or num_format is not None:
            section = self._find_containing_section(document)
            if section_endnote:
                section._sectPr.endnote_position = "sectEnd"  # pyright: ignore[reportPrivateUsage]
                document.settings.endnote_position = "sectEnd"
            if num_format is not None:
                section._sectPr.endnote_number_format = num_format  # pyright: ignore[reportPrivateUsage]

        return endnote

    def add_footnote(self, num_format: str | None = None, auto_paragraph: bool = True) -> Footnote:
        """Append a footnote reference to this paragraph and return the newly created
        |Footnote|.

        The footnote is inserted into |Footnotes| at the position that keeps footnotes
        ordered by reference id, renumbering any footnotes that come after it in the
        document as needed.

        When `auto_paragraph` is |True| (the default), the footnote is given an initial
        paragraph containing the footnote reference mark, matching how Word creates a
        new footnote. Use `footnote.add_paragraph()` to add the footnote's actual
        content.

        When `num_format` is provided, the section containing this paragraph has its
        footnote-numbering-format overridden; see
        :attr:`.Section.footnote_number_format`.
        """
        document = find_containing_document(self)
        new_fr_id = document._calculate_next_footnote_reference_id(  # pyright: ignore[reportPrivateUsage]
            self._p
        )
        r = self._p.add_r()
        r.add_footnoteReference(new_fr_id)
        footnote = document._add_footnote(new_fr_id)  # pyright: ignore[reportPrivateUsage]

        if auto_paragraph:
            p = footnote.add_paragraph()
            p._p.style = "FootnoteText"
            ref_r = p._p.add_r()
            ref_rPr = ref_r.get_or_add_rPr()
            ref_rPr.style = "FootnoteReference"
            ref_r.add_footnoteRef()

        if num_format is not None:
            section = self._find_containing_section(document)
            section._sectPr.footnote_number_format = num_format  # pyright: ignore[reportPrivateUsage]

        return footnote

    def add_run(self, text: str | None = None, style: str | CharacterStyle | None = None) -> Run:
        """Append run containing `text` and having character-style `style`.

        `text` can contain tab (``\\t``) characters, which are converted to the
        appropriate XML form for a tab. `text` can also include newline (``\\n``) or
        carriage return (``\\r``) characters, each of which is converted to a line
        break. When `text` is `None`, the new run is empty.
        """
        r = self._p.add_r()
        run = Run(r, self)
        if text:
            run.text = text
        if style:
            run.style = style
        return run

    @property
    def alignment(self) -> WD_PARAGRAPH_ALIGNMENT | None:
        """A member of the :ref:`WdParagraphAlignment` enumeration specifying the
        justification setting for this paragraph.

        A value of |None| indicates the paragraph has no directly-applied alignment
        value and will inherit its alignment value from its style hierarchy. Assigning
        |None| to this property removes any directly-applied alignment value.
        """
        return self._p.alignment

    @alignment.setter
    def alignment(self, value: WD_PARAGRAPH_ALIGNMENT):
        self._p.alignment = value

    def clear(self):
        """Return this same paragraph after removing all its content.

        Paragraph-level formatting, such as style, is preserved.
        """
        self._p.clear_content()
        return self

    @property
    def contains_page_break(self) -> bool:
        """`True` when one or more rendered page-breaks occur in this paragraph."""
        return bool(self._p.lastRenderedPageBreaks)

    @property
    def endnotes(self) -> List[Endnote]:
        """The |Endnote| referenced by each `w:endnoteReference` in this paragraph."""
        document = find_containing_document(self)
        return [document.endnotes[id] for id in self._p.endnote_reference_ids]

    @property
    def footnotes(self) -> List[Footnote]:
        """The |Footnote| referenced by each `w:footnoteReference` in this
        paragraph."""
        document = find_containing_document(self)
        return [document.footnotes[id] for id in self._p.footnote_reference_ids]

    @property
    def hyperlinks(self) -> List[Hyperlink]:
        """A |Hyperlink| instance for each hyperlink in this paragraph."""
        return [Hyperlink(hyperlink, self) for hyperlink in self._p.hyperlink_lst]

    def increment_containing_endnote_reference_ids(self) -> None:
        """Increment the `@w:id` of each `w:endnoteReference` in this paragraph's runs
        by one.

        Used to renumber later endnotes when an earlier one is inserted ahead of them.
        """
        for run in self.runs:
            run._r.increment_containing_endnote_reference_ids()  # pyright: ignore[reportPrivateUsage]

    def increment_containing_footnote_reference_ids(self) -> None:
        """Increment the `@w:id` of each `w:footnoteReference` in this paragraph's runs
        by one.

        Used to renumber later footnotes when an earlier one is inserted ahead of them.
        """
        for run in self.runs:
            run._r.increment_containing_footnote_reference_ids()  # pyright: ignore[reportPrivateUsage]

    def insert_paragraph_before(
        self, text: str | None = None, style: str | ParagraphStyle | None = None
    ) -> Paragraph:
        """Return a newly created paragraph, inserted directly before this paragraph.

        If `text` is supplied, the new paragraph contains that text in a single run. If
        `style` is provided, that style is assigned to the new paragraph.
        """
        paragraph = self._insert_paragraph_before()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    def iter_inner_content(self) -> Iterator[Run | Hyperlink]:
        """Generate the runs and hyperlinks in this paragraph, in the order they appear.

        The content in a paragraph consists of both runs and hyperlinks. This method
        allows accessing each of those separately, in document order, for when the
        precise position of the hyperlink within the paragraph text is important. Note
        that a hyperlink itself contains runs.
        """
        for r_or_hlink in self._p.inner_content_elements:
            yield (
                Run(r_or_hlink, self)
                if isinstance(r_or_hlink, CT_R)
                else Hyperlink(r_or_hlink, self)
            )

    @property
    def paragraph_format(self):
        """The |ParagraphFormat| object providing access to the formatting properties
        for this paragraph, such as line spacing and indentation."""
        return ParagraphFormat(self._element)

    @property
    def rendered_page_breaks(self) -> List[RenderedPageBreak]:
        """All rendered page-breaks in this paragraph.

        Most often an empty list, sometimes contains one page-break, but can contain
        more than one is rare or contrived cases.
        """
        return [RenderedPageBreak(lrpb, self) for lrpb in self._p.lastRenderedPageBreaks]

    @property
    def runs(self) -> List[Run]:
        """Sequence of |Run| instances corresponding to the <w:r> elements in this
        paragraph."""
        return [Run(r, self) for r in self._p.r_lst]

    @property
    def style(self) -> ParagraphStyle | None:
        """Read/Write.

        |_ParagraphStyle| object representing the style assigned to this paragraph. If
        no explicit style is assigned to this paragraph, its value is the default
        paragraph style for the document. A paragraph style name can be assigned in lieu
        of a paragraph style object. Assigning |None| removes any applied style, making
        its effective value the default paragraph style for the document.
        """
        style_id = self._p.style
        style = self.part.get_style(style_id, WD_STYLE_TYPE.PARAGRAPH)
        return cast(ParagraphStyle, style)

    @style.setter
    def style(self, style_or_name: str | ParagraphStyle | None):
        style_id = self.part.get_style_id(style_or_name, WD_STYLE_TYPE.PARAGRAPH)
        self._p.style = style_id

    @property
    def text(self) -> str:
        """The textual content of this paragraph.

        The text includes the visible-text portion of any hyperlinks in the paragraph.
        Tabs and line breaks in the XML are mapped to ``\\t`` and ``\\n`` characters
        respectively.

        Assigning text to this property causes all existing paragraph content to be
        replaced with a single run containing the assigned text. A ``\\t`` character in
        the text is mapped to a ``<w:tab/>`` element and each ``\\n`` or ``\\r``
        character is mapped to a line break. Paragraph-level formatting, such as style,
        is preserved. All run-level formatting, such as bold or italic, is removed.
        """
        return self._p.text

    @text.setter
    def text(self, text: str | None):
        self.clear()
        self.add_run(text)

    def _find_containing_section(self, document: Document) -> Section:
        """The |Section| that contains this paragraph.

        In WordprocessingML, a paragraph belongs to the section whose `w:sectPr`
        element comes after it in document order -- the last section's `w:sectPr` is
        stored on `w:body` itself, while every other section's `w:sectPr` is the last
        child of the paragraph that ends that section.
        """
        all_paragraphs = document.paragraphs

        def ends_a_section(p: Paragraph) -> bool:
            pPr = p._p.pPr
            if pPr is None:
                return False
            return cast("CT_SectPr | None", pPr.sectPr) is not None

        try:
            para_index = next(i for i, p in enumerate(all_paragraphs) if p._p is self._p)
        except StopIteration:
            return document.sections[-1]

        for i in range(para_index, len(all_paragraphs)):
            if ends_a_section(all_paragraphs[i]):
                section_index = sum(1 for p in all_paragraphs[: i + 1] if ends_a_section(p))
                return document.sections[section_index]

        return document.sections[-1]

    def _insert_paragraph_before(self):
        """Return a newly created paragraph, inserted directly before this paragraph."""
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)
