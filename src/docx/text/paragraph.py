"""Paragraph-related proxy types."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, List, cast

from docx.bookmark import BookmarkParent
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.text.run import CT_R
from docx.shared import StoryChild
from docx.styles.style import ParagraphStyle
from docx.text.hyperlink import Hyperlink
from docx.text.pagebreak import RenderedPageBreak
from docx.text.parfmt import ParagraphFormat
from docx.text.run import Run

if TYPE_CHECKING:
    import docx.types as t
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml.bookmark import CT_BookmarkEnd, CT_BookmarkStart
    from docx.oxml.text.font import CT_RPr
    from docx.oxml.text.paragraph import CT_P
    from docx.sdt import SdtBase
    from docx.styles.style import CharacterStyle


class Paragraph(StoryChild, BookmarkParent):
    """Proxy object wrapping a `<w:p>` element."""

    def __init__(self, p: CT_P, parent: t.ProvidesStoryPart):
        super(Paragraph, self).__init__(parent)
        self._p = self._element = p

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

    def add_sdt(
        self,
        tag_name: str,
        text: str = "",
        alias_name: str = "",
        placeholder_txt: str | None = None,
        style: str = "Normal",
        bold: bool = False,
        italic: bool = False,
    ) -> SdtBase:
        """Return a new inline content control (structured document tag) appended to
        the end of this paragraph, tagged `tag_name`.

        When `text` is not provided, the content control displays `placeholder_txt`
        (or a default prompt) as its placeholder text. `alias_name` defaults to
        `tag_name` when not provided. `style`, `bold`, and `italic` apply character
        formatting to the content control's default and, when `text` is provided, its
        actual content.
        """
        from docx.sdt import SdtBase

        sdt = self._p._new_sdt()  # pyright: ignore[reportPrivateUsage]
        sdtPr = sdt.get_or_add_sdtPr()
        sdtPr.name = tag_name
        sdtPr.alias_val = alias_name or tag_name
        _apply_run_formatting(sdtPr.get_or_add_rPr(), style, bold, italic)

        sdtContent = sdt.get_or_add_sdtContent()
        r = sdtContent._add_r()  # pyright: ignore[reportPrivateUsage]
        if text:
            _apply_run_formatting(r.get_or_add_rPr(), style, bold, italic)
            r.text = text
        else:
            rPr = r.get_or_add_rPr()
            rPr.get_or_add_rStyle().val = "PlaceholderText"
            rPr.get_or_add_b()
            rPr.get_or_add_bCs()
            r.text = placeholder_txt or "Click or tap here to enter text"
            sdtPr.active_placeholder = True

        self._p.append(sdt)
        return SdtBase(sdt, self)

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

    @property
    def bookmark_ends(self) -> List[CT_BookmarkEnd]:
        """The `w:bookmarkEnd` elements appearing directly in this paragraph."""
        return self._p.bookmarkEnd_lst

    @property
    def bookmark_starts(self) -> List[CT_BookmarkStart]:
        """The `w:bookmarkStart` elements appearing directly in this paragraph."""
        return self._p.bookmarkStart_lst

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
    def hyperlinks(self) -> List[Hyperlink]:
        """A |Hyperlink| instance for each hyperlink in this paragraph."""
        return [Hyperlink(hyperlink, self) for hyperlink in self._p.hyperlink_lst]

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
    def sdts(self) -> List[SdtBase]:
        """The inline content controls (structured document tags) directly contained
        in this paragraph, in document order."""
        from docx.sdt import SdtBase

        return [SdtBase(sdt, self) for sdt in self._p.sdt_lst]

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

    def _insert_paragraph_before(self):
        """Return a newly created paragraph, inserted directly before this paragraph."""
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)


def _apply_run_formatting(rPr: CT_RPr, style: str, bold: bool, italic: bool) -> None:
    """Apply `style`, `bold`, and `italic` to `rPr`, a run-properties element not
    otherwise attached to a `Run` proxy object (such as a content control's default
    `w:sdtPr/w:rPr`)."""
    if style != "Normal":
        rPr.get_or_add_rStyle().val = style
    if bold:
        rPr.get_or_add_b()
        rPr.get_or_add_bCs()
    if italic:
        rPr.get_or_add_i()
