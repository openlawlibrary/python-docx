"""Paragraph-related proxy types."""

from __future__ import annotations

import math
from typing import TYPE_CHECKING, Any, Iterator, List, cast

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.text.run import CT_R
from docx.shared import Inches, Length, StoryChild
from docx.styles.style import ParagraphStyle
from docx.text.hyperlink import Hyperlink
from docx.text.pagebreak import RenderedPageBreak
from docx.text.parfmt import ParagraphFormat
from docx.text.run import Run

if TYPE_CHECKING:
    import docx.types as t
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml.numbering import CT_Lvl, CT_Numbering
    from docx.oxml.text.paragraph import CT_P
    from docx.parts.document import DocumentPart
    from docx.styles.style import CharacterStyle


class Paragraph(StoryChild):
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
    def hyperlinks(self) -> List[Hyperlink]:
        """A |Hyperlink| instance for each hyperlink in this paragraph."""
        return [Hyperlink(hyperlink, self) for hyperlink in self._p.hyperlink_lst]

    def insert_paragraph_before(
        self,
        text: str | None = None,
        style: str | ParagraphStyle | None = None,
        ilvl: int | None = None,
    ) -> Paragraph:
        """Return a newly created paragraph, inserted directly before this paragraph.

        If `text` is supplied, the new paragraph contains that text in a single run. If
        `style` is provided, that style is assigned to the new paragraph. If `ilvl` is
        provided, the new paragraph joins this paragraph's numbered list (if any) at
        indentation level `ilvl`; see :meth:`.set_li_lvl`.
        """
        paragraph = self._insert_paragraph_before()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        if ilvl is not None:
            paragraph.set_li_lvl(self, ilvl)
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
    def lvl_from_para_props(self) -> CT_Lvl | None:
        """The `w:lvl` numbering-level definition governing this paragraph, resolved
        via its own direct numbering properties (not those inherited from its style).

        |None| if this paragraph is not part of a numbered list.
        """
        try:
            return self._p.lvl_from_para_props(self._numbering_element)
        except (AttributeError, NotImplementedError):
            return None

    @property
    def lvl_from_style_props(self) -> CT_Lvl | None:
        """The `w:lvl` numbering-level definition governing this paragraph, resolved
        via its paragraph style's numbering properties.

        |None| if this paragraph's style is not part of a numbered list.
        """
        try:
            return self._p.lvl_from_style_props(
                self._numbering_element, self._document_part.cached_styles
            )
        except (AttributeError, NotImplementedError):
            return None

    @property
    def norm_left_indent(self) -> Length:
        """The effective left indentation of this paragraph, in inches, unifying tab
        characters, tab stops, and first-line/hanging indentation.

        Resolves indentation set anywhere in the style hierarchy, in priority order
        from lowest to highest: the paragraph style's base style, the paragraph
        style's own numbering format, the paragraph style's direct paragraph format,
        this paragraph's own numbering format, and finally this paragraph's own direct
        paragraph format. When no explicit indentation is set anywhere in that chain,
        falls back to counting leading tab characters against the applicable tab
        stops (or, absent those, the OOXML default tab stop of half an inch).
        """

        def get_base_style_attr(obj: Any, attr_name: str) -> Length:
            base_style = getattr(obj, "base_style", None)
            if base_style is None:
                return Length(0)
            base_paragraph_format = getattr(base_style, "paragraph_format", None)
            if base_paragraph_format is not None:
                value = getattr(base_paragraph_format, attr_name, None)
                if value is not None:
                    return value
            return get_base_style_attr(base_style, attr_name)

        def get_tabstops(para: Any) -> list[float]:
            """Tab-stop positions (in inches) from `para`'s style hierarchy.

            A `clear` tab-stop lower in the hierarchy removes a tab-stop from higher in
            the hierarchy.
            """
            from docx.enum.text import WD_TAB_ALIGNMENT

            tabstops: list[float] = []

            def _inner_get_tabstops(obj: Any) -> None:
                nonlocal tabstops
                if obj is None:
                    return
                if hasattr(obj, "base_style"):
                    _inner_get_tabstops(obj.base_style)
                elif hasattr(obj, "style"):
                    _inner_get_tabstops(obj.style)

                obj = obj.paragraph_format

                tabstops.extend(round(ts.position.inches, 2) for ts in obj.tab_stops)
                clear_t_stops = [
                    round(ts.position.inches, 2)
                    for ts in obj.tab_stops
                    if ts.alignment == WD_TAB_ALIGNMENT.CLEAR
                ]
                tabstops[:] = [ts for ts in tabstops if ts not in clear_t_stops]

            _inner_get_tabstops(para)
            return tabstops

        def apply_formatting(
            source: Any, first_line_indent: Length | None, left_indent: Length | None
        ) -> tuple[Length | None, Length | None]:
            if source:
                if getattr(source, "first_line_indent", None) is not None:
                    first_line_indent = source.first_line_indent
                if getattr(source, "left_indent", None) is not None:
                    left_indent = source.left_indent
            return first_line_indent, left_indent

        # -- Apply paragraph styles by priority (from lowest to highest). Formatting
        # -- from the base style has the lowest priority.
        first_line_indent = get_base_style_attr(self.style, "first_line_indent")
        left_indent = get_base_style_attr(self.style, "left_indent")
        # -- Next, formatting from numbering properties defined in the paragraph style.
        first_line_indent, left_indent = apply_formatting(
            self.style_numbering_format, first_line_indent, left_indent
        )
        # -- Then formatting from the paragraph style itself.
        first_line_indent, left_indent = apply_formatting(
            self.style.paragraph_format if self.style is not None else None,
            first_line_indent,
            left_indent,
        )
        # -- Next, formatting from numbering properties on the paragraph directly.
        first_line_indent, left_indent = apply_formatting(
            self.para_numbering_format, first_line_indent, left_indent
        )
        # -- Finally, direct paragraph formatting.
        first_line_indent, left_indent = apply_formatting(
            self.paragraph_format, first_line_indent, left_indent
        )

        left_indent_in = round(left_indent.inches, 2) if left_indent else 0
        first_line_indent_in = round(first_line_indent.inches, 2) if first_line_indent else 0
        # -- `first_line_indent_in` becomes the paragraph's absolute first-line
        # -- position (not just its offset from `left_indent_in`) for the remainder
        # -- of this calculation, matching `indent`'s starting value.
        indent = first_line_indent_in = first_line_indent_in + left_indent_in

        # -- Count leading tab characters (ignoring regular spaces). --
        text = self.text
        tab_count = text[: len(text) - len(text.lstrip())].count("\t")

        if tab_count:
            DEFAULT_TAB_STOP = 0.5

            # -- Only tab stops to the right of the first-line indent affect this
            # -- paragraph's indentation.
            tab_stops = [ts for ts in get_tabstops(self) if ts > first_line_indent_in]

            # -- If the first-line indent is left of the paragraph indent, the first
            # -- tab lands on the paragraph indent.
            if first_line_indent_in < left_indent_in:
                tab_stops.append(left_indent_in)

            tab_stops = sorted(set(tab_stops))

            if tab_stops and len(tab_stops) >= tab_count:
                indent = tab_stops[tab_count - 1]
            else:
                if tab_stops:
                    indent = tab_stops[-1]
                    tab_count -= len(tab_stops)

                # -- Calculate in whole tab-stop units instead of inches. --
                indent *= 1 / DEFAULT_TAB_STOP

                # -- Round up to the next tab-stop indent (or add one if already a
                # -- whole number).
                tab_count -= 1
                indent = math.ceil(indent) if not indent.is_integer() else indent + 1

                # -- Each remaining tab char just adds one whole indent. --
                indent += tab_count

                # -- Scale back to inches. --
                indent /= 1 / DEFAULT_TAB_STOP

        return Inches(indent)

    @property
    def number(self) -> str | None:
        """This paragraph's list-item label, e.g. `"1)\\t"` or `"(a) "`.

        |None| if this paragraph is not part of a numbered list, or its numbering
        format is not one of the supported ordinal formats (an unordered/bullet list,
        for example, always returns |None|).
        """
        # -- a paragraph with no `w:pPr` at all can have neither direct numbering
        # -- properties nor a paragraph style reference, so it can never be part of a
        # -- numbered list; skip the numbering/styles-part lookup entirely in that
        # -- (common) case.
        if self._p.pPr is None:
            return None
        try:
            return self._p.number(self._numbering_element, self._document_part.cached_styles)
        except (AttributeError, NotImplementedError):
            return None

    @property
    def para_numbering_format(self) -> ParagraphFormat | None:
        """|ParagraphFormat| for this paragraph's numbering level, resolved via its own
        direct numbering properties.

        Takes priority over :attr:`.style_numbering_format`. |None| if this paragraph
        is not part of a numbered list.
        """
        lvl = self.lvl_from_para_props
        return None if lvl is None else ParagraphFormat(lvl)

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

    def set_li_lvl(self, prev: Paragraph | None, ilvl: int | None = None) -> None:
        """Set this paragraph's list-item indentation level.

        When `prev` is given, this paragraph joins `prev`'s numbered list (copying its
        `numId`) at level `ilvl` -- or `prev`'s own level, when `ilvl` is |None|. When
        `prev` is |None|, this paragraph starts a new numbering-list instance, reusing
        the numbering definition already referenced by this paragraph's (or its
        style's) numbering properties, at level `ilvl` -- or level 0, when `ilvl` is
        |None|.
        """
        prev_el = prev._element if prev is not None else None
        self._p.set_li_lvl(
            self._numbering_element, self._document_part.cached_styles, prev_el, ilvl
        )

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
    def style_numbering_format(self) -> ParagraphFormat | None:
        """|ParagraphFormat| for this paragraph's numbering level, resolved via its
        paragraph style's numbering properties.

        |None| if this paragraph's style is not part of a numbered list.
        """
        lvl = self.lvl_from_style_props
        return None if lvl is None else ParagraphFormat(lvl)

    @property
    def text(self) -> str:
        """The textual content of this paragraph.

        The text includes the visible-text portion of any hyperlinks in the paragraph,
        preceded by this paragraph's list-item label (e.g. `"1)\\t"`) when it is part
        of a numbered list. Tabs and line breaks in the XML are mapped to ``\\t`` and
        ``\\n`` characters respectively.

        Assigning text to this property causes all existing paragraph content to be
        replaced with a single run containing the assigned text. A ``\\t`` character in
        the text is mapped to a ``<w:tab/>`` element and each ``\\n`` or ``\\r``
        character is mapped to a line break. Paragraph-level formatting, such as style,
        is preserved. All run-level formatting, such as bold or italic, is removed.
        """
        number = self.number
        return (number if number is not None else "") + self._p.text

    @text.setter
    def text(self, text: str | None):
        self.clear()
        self.add_run(text)

    @property
    def _numbering_element(self) -> CT_Numbering:
        """The root `w:numbering` element of this document's numbering part."""
        return cast(
            "CT_Numbering",
            self._document_part.numbering_part._element,  # pyright: ignore[reportPrivateUsage]
        )

    @property
    def _document_part(self) -> DocumentPart:
        """This paragraph's containing part, narrowed from `StoryPart` to
        `DocumentPart`.

        Numbering resolution (and the `cached_styles`/`styles` lookups it needs) is
        currently only implemented on `DocumentPart`; a numbered paragraph nested in a
        header, footer, or comment would need that same support added there before
        this cast would be safe for it.
        """
        return cast("DocumentPart", self.part)

    def _insert_paragraph_before(self):
        """Return a newly created paragraph, inserted directly before this paragraph."""
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)
