"""Paragraph-related proxy types."""

from __future__ import annotations

import copy
import math
from typing import TYPE_CHECKING, Any, Iterator, List, cast

from docx.bookmark import BookmarkParent
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.text.run import CT_R
from docx.shared import Inches, Length, StoryChild, find_containing_document
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
    from docx.oxml.bookmark import CT_BookmarkEnd, CT_BookmarkStart
    from docx.oxml.numbering import CT_Lvl, CT_Numbering
    from docx.oxml.section import CT_SectPr
    from docx.oxml.text.font import CT_RPr
    from docx.oxml.text.paragraph import CT_P
    from docx.parts.document import DocumentPart
    from docx.sdt import SdtBase
    from docx.section import Section
    from docx.styles.style import CharacterStyle


class Paragraph(StoryChild, BookmarkParent):
    """Proxy object wrapping a `<w:p>` element."""

    def __init__(self, p: CT_P, parent: t.ProvidesStoryPart):
        super(Paragraph, self).__init__(parent)
        self._p = self._element = p

    def __getstate__(self):
        state = dict(self.__dict__)
        state.pop("_parent", None)
        return state

    def __repr__(self):
        text_stripped = self.text.strip()
        text = text_stripped[:20]
        if len(text_stripped) > len(text):
            text += "..."
        if not text:
            text = "EMPTY PARAGRAPH"
        return f'<p:"{text}">'

    def __setstate__(self, state: dict[str, Any]):
        self.__dict__ = state

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

    def add_field(self, instrText: str | None = None) -> None:
        """Append a complex field to this paragraph.

        A complex field is a `begin`/`end` pair of `<w:fldChar>` elements, each in
        their own run, optionally bracketing an `<w:instrText>`-bearing run holding the
        field-code instruction (e.g. `"DATE"`, `"TOC \\o '1-3'"`) when `instrText` is
        given.

        Word computes and caches the field's displayed *result* the next time it
        opens or updates fields in this document; until then the field displays
        nothing, matching Word's own behavior for a field with no cached result yet.
        """
        self.add_run().add_fldChar()
        if instrText:
            self.add_run().add_instrText(instrText)
        self.add_run().add_fldChar(fldCharType="end")

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

    def clone(self) -> Paragraph:
        """Return a copy of this paragraph, cloned by selective deep copying.

        The clone is not attached to a document body; the caller is responsible for
        inserting it into the document tree.
        """
        c = copy.deepcopy(self)
        c._parent = self._parent
        return c

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
        self._p.strip_hidden_fld_char_content()
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

    def insert_text(self, position: int, new_text: str) -> Paragraph:
        """Insert `new_text` at `position` in this paragraph's text, retaining runs and
        their formatting."""
        runend = 0
        runstart = 0
        for run in self.runs:
            runstart = runend
            runend += len(run.text)
            if runend >= position:
                run.text = (
                    run.text[: position - runstart] + new_text + run.text[position - runstart :]
                )
                break
        return self

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

    def lstrip(self, chars: str | None = None) -> Paragraph:
        """Left-strip whitespace (or `chars` if given) from this paragraph's text."""
        while self.runs:
            run = self.runs[0]
            run.text = run.text.lstrip(chars)
            if not run.text:
                run._r.getparent().remove(run._r)  # pyright: ignore[reportPrivateUsage, reportOptionalMemberAccess]
            else:
                break
        return self

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
    def remove_new_line_breaks(self) -> None:
        """Replace each line-break (`w:br`) in this paragraph's runs with a space."""
        for run in self.runs:
            run._r.remove_br_tag_childrens()  # pyright: ignore[reportPrivateUsage]

    def remove(self) -> None:
        """Remove this paragraph from its containing document body."""
        parent = self._p.getparent()
        assert parent is not None
        parent.remove(self._p)

    def remove_text(self, start: int = 0, end: int = -1) -> Paragraph:
        """Remove the text in `[start, end)`, retaining the surrounding runs and their
        formatting."""
        if end == -1:
            end = len(self.text)
        assert end > start
        assert end <= len(self.text)

        # -- special case: both start and end fall within a single run --
        runstart = 0
        for run in self.runs:
            runend = runstart + len(run.text)
            if runstart <= start and end <= runend:
                run.text = run.text[: start - runstart] + run.text[end - runstart :]
                if not run.text:
                    run._r.getparent().remove(run._r)  # pyright: ignore[reportPrivateUsage, reportOptionalMemberAccess]
                return self
            runstart = runend

        # -- general case: the removed range spans multiple runs --
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
                to_del._r.getparent().remove(to_del._r)  # pyright: ignore[reportPrivateUsage, reportOptionalMemberAccess]
            else:
                runstart = runend
        return self

    @property
    def rendered_page_breaks(self) -> List[RenderedPageBreak]:
        """All rendered page-breaks in this paragraph.

        Most often an empty list, sometimes contains one page-break, but can contain
        more than one is rare or contrived cases.
        """
        return [RenderedPageBreak(lrpb, self) for lrpb in self._p.lastRenderedPageBreaks]

    def replace_char(self, oldch: str, newch: str) -> Paragraph:
        """Replace all occurrences of `oldch` with `newch` in this paragraph's text."""
        for run in self.runs:
            run.text = run.text.replace(oldch, newch)
        return self

    def replace_chars(self, *replacement_pairs: tuple[str, str]) -> Paragraph:
        """Replace, for each `(oldch, newch)` pair in `replacement_pairs`, all
        occurrences of `oldch` with `newch` in this paragraph's text."""
        for run in self.runs:
            new_text = run.text
            for oldch, newch in replacement_pairs:
                new_text = new_text.replace(oldch, newch)
            run.text = new_text
        return self

    def replace_text(self, old_text: str, new_text: str) -> Paragraph:
        """Replace all occurrences of `old_text` with `new_text`, retaining run
        formatting. `old_text` can span multiple runs; `new_text` is added to the run
        where `old_text` starts."""
        assert new_text
        assert old_text
        startpos = 0
        while startpos < len(self.text):
            try:
                old_start = startpos + self.text[startpos:].index(old_text)
                startpos = old_start + len(old_text)
            except ValueError:
                break
            self.remove_text(start=old_start, end=startpos).insert_text(old_start, new_text)
        return self

    def rstrip(self, chars: str | None = None) -> Paragraph:
        """Right-strip whitespace (or `chars` if given) from this paragraph's text."""
        while self.runs:
            run = self.runs[-1]
            run.text = run.text.rstrip(chars)
            if not run.text:
                run._r.getparent().remove(run._r)  # pyright: ignore[reportPrivateUsage, reportOptionalMemberAccess]
            else:
                break
        return self

    @property
    def runs(self) -> List[Run]:
        """Sequence of |Run| instances corresponding to the <w:r> elements in this
        paragraph."""
        self._p.strip_hidden_fld_char_content()
        return [Run(r, self) for r in self._p.r_lst]

    def split(self, *positions: int) -> List[Paragraph]:
        """Split this paragraph at each offset in `positions`, keeping formatting.

        This paragraph is replaced in the document by the returned paragraphs: the
        original `w:p` element (retaining its runs up to the first split position) is
        kept and reused as the first returned paragraph, while a new `w:p` element is
        inserted immediately after it for each subsequent segment.
        """
        remaining = list(positions)
        for p in remaining:
            assert 0 < p < len(self.text)
        paras: List[Paragraph] = []
        splitpos = remaining.pop(0)
        curpos = 0
        runidx = 0
        curpara = self
        next_para = curpara
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
                        parent = crun._r.getparent()  # pyright: ignore[reportPrivateUsage]
                        assert parent is not None
                        parent.remove(crun._r)  # pyright: ignore[reportPrivateUsage]
                for crunidx, crun in enumerate(next_para.runs):
                    if crunidx < runidx + idx_cor:
                        parent = crun._r.getparent()  # pyright: ignore[reportPrivateUsage]
                        assert parent is not None
                        parent.remove(crun._r)  # pyright: ignore[reportPrivateUsage]
                curpara._p.addnext(next_para._p)
                paras.append(curpara)
                if not remaining:
                    break
                curpos = splitpos
                splitpos = remaining.pop(0)
                curpara = next_para
                runidx = 0
            else:
                runidx += 1
                curpos = endpos

        paras.append(next_para)
        return paras

    def strip(self, chars: str | None = None) -> Paragraph:
        """Strip whitespace (or `chars` if given) from both ends of this paragraph's
        text."""
        return self.lstrip(chars).rstrip(chars)

    @property
    def sdts(self) -> List[SdtBase]:
        """The inline content controls (structured document tags) directly contained
        in this paragraph, in document order."""
        from docx.sdt import SdtBase

        return [SdtBase(sdt, self) for sdt in self._p.sdt_lst]

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
