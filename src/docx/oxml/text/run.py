"""Custom element classes related to text runs (CT_R)."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator, List, cast

from docx.oxml.drawing import CT_Drawing
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.simpletypes import ST_BrClear, ST_BrType, ST_FldCharType
from docx.oxml.text.font import CT_RPr
from docx.oxml.text.symbol import CT_Sym
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)
from docx.shared import TextAccumulator

if TYPE_CHECKING:
    from docx.oxml.shape import CT_Anchor, CT_Inline
    from docx.oxml.text.footnote_reference import CT_FtnEdnRef
    from docx.oxml.text.pagebreak import CT_LastRenderedPageBreak
    from docx.oxml.text.parfmt import CT_TabStop

# ------------------------------------------------------------------------------------
# Run-level elements


class CT_R(BaseOxmlElement):
    """`<w:r>` element, containing the properties and text for a run."""

    add_br: Callable[[], CT_Br]
    add_fldChar: Callable[[], CT_FldChar]
    add_instrText: Callable[[], BaseOxmlElement]
    add_tab: Callable[[], CT_TabStop]
    get_or_add_rPr: Callable[[], CT_RPr]
    _add_drawing: Callable[[], CT_Drawing]
    _add_sym: Callable[[], CT_Sym]
    _add_t: Callable[..., CT_Text]
    _add_footnoteReference: Callable[[], CT_FtnEdnRef]
    _add_endnoteReference: Callable[[], CT_FtnEdnRef]
    _add_footnoteRef: Callable[[], CT_FtnEdnRef]
    _add_endnoteRef: Callable[[], CT_FtnEdnRef]
    footnoteReference_lst: List[CT_FtnEdnRef]
    endnoteReference_lst: List[CT_FtnEdnRef]

    bookmarkStart = ZeroOrMore(
        "w:bookmarkStart", successors=("w:t", "w:rPr", "w:br", "w:cr", "w:tab", "w:drawing")
    )
    rPr: CT_RPr | None = ZeroOrOne("w:rPr")  # pyright: ignore[reportAssignmentType]
    br = ZeroOrMore("w:br")
    cr = ZeroOrMore("w:cr")
    drawing = ZeroOrMore("w:drawing")
    sym = ZeroOrMore("w:sym")
    fldChar = ZeroOrMore("w:fldChar")
    instrText = ZeroOrMore("w:instrText")
    t = ZeroOrMore("w:t")
    tab = ZeroOrMore("w:tab")
    bookmarkEnd = ZeroOrMore("w:bookmarkEnd")
    footnoteReference = ZeroOrMore("w:footnoteReference")
    footnoteRef = ZeroOrMore("w:footnoteRef")
    endnoteReference = ZeroOrMore("w:endnoteReference")
    endnoteRef = ZeroOrMore("w:endnoteRef")

    def add_footnoteReference(self, id: int) -> CT_FtnEdnRef:
        """Return a newly added `w:footnoteReference` element containing `id`."""
        rPr = self.get_or_add_rPr()
        rPr.style = "FootnoteReference"
        new_fr = self._add_footnoteReference()
        new_fr.id = id
        return new_fr

    def add_endnoteReference(self, id: int) -> CT_FtnEdnRef:
        """Return a newly added `w:endnoteReference` element containing `id`."""
        rPr = self.get_or_add_rPr()
        rPr.style = "EndnoteReference"
        new_er = self._add_endnoteReference()
        new_er.id = id
        return new_er

    def add_footnoteRef(self) -> CT_FtnEdnRef:
        """Return a newly added `w:footnoteRef` element.

        This element displays the footnote reference mark within the footnote itself.
        """
        return self._add_footnoteRef()

    def add_endnoteRef(self) -> CT_FtnEdnRef:
        """Return a newly added `w:endnoteRef` element.

        This element displays the endnote reference mark within the endnote itself.
        """
        return self._add_endnoteRef()

    def add_t(self, text: str) -> CT_Text:
        """Return a newly added `<w:t>` element containing `text`."""
        t = self._add_t(text=text)
        if len(text.strip()) < len(text):
            t.set(qn("xml:space"), "preserve")
        return t

    def add_symbol(self, char: str | None, font: str | None) -> CT_Sym:
        """Return a newly added `<w:sym>` element with `char`/`font` attributes."""
        sym = self._add_sym()
        if char is not None:
            sym.char = char
        if font is not None:
            sym.font = font
        return sym

    def add_drawing(self, inline_or_anchor: CT_Inline | CT_Anchor) -> CT_Drawing:
        """Return newly appended `CT_Drawing` (`w:drawing`) child element.

        The `w:drawing` element has `inline_or_anchor` as its child.
        """
        drawing = self._add_drawing()
        drawing.append(inline_or_anchor)
        return drawing

    def clear_content(self) -> None:
        """Remove all child elements except a `w:rPr`, `w:footnoteReference`, or
        `w:endnoteReference` element if present."""
        # -- remove all run inner-content except a `w:rPr` when present, and keep any
        # -- footnote/endnote reference since those anchor content elsewhere in the
        # -- document that would otherwise become orphaned.
        for e in self.xpath(
            "./*[not(self::w:rPr)"
            " and not(self::w:footnoteReference)"
            " and not(self::w:endnoteReference)]"
        ):
            self.remove(e)

    @property
    def endnote_reference_ids(self) -> Iterator[int]:
        """Generate the `@w:id` of each `w:endnoteReference` child of this run."""
        for child in self:
            if child.tag == qn("w:endnoteReference"):
                yield cast("CT_FtnEdnRef", child).id

    @property
    def footnote_reference_ids(self) -> Iterator[int]:
        """Generate the `@w:id` of each `w:footnoteReference` child of this run."""
        for child in self:
            if child.tag == qn("w:footnoteReference"):
                yield cast("CT_FtnEdnRef", child).id

    def increment_containing_endnote_reference_ids(self) -> None:
        """Increment the `@w:id` of each `w:endnoteReference` child of this run by
        one."""
        for endnoteReference in self.endnoteReference_lst:
            endnoteReference.id += 1

    def increment_containing_footnote_reference_ids(self) -> None:
        """Increment the `@w:id` of each `w:footnoteReference` child of this run by
        one."""
        for footnoteReference in self.footnoteReference_lst:
            footnoteReference.id += 1

    @property
    def inner_content_items(self) -> List[str | CT_Drawing | CT_LastRenderedPageBreak]:
        """Text of run, possibly punctuated by `w:lastRenderedPageBreak` elements."""
        from docx.oxml.text.pagebreak import CT_LastRenderedPageBreak

        accum = TextAccumulator()

        def iter_items() -> Iterator[str | CT_Drawing | CT_LastRenderedPageBreak]:
            for e in self.xpath(
                "w:br"
                " | w:cr"
                " | w:drawing"
                " | w:lastRenderedPageBreak"
                " | w:noBreakHyphen"
                " | w:ptab"
                " | w:sym"
                " | w:t"
                " | w:tab"
            ):
                if isinstance(e, (CT_Drawing, CT_LastRenderedPageBreak)):
                    yield from accum.pop()
                    yield e
                else:
                    accum.push(str(e))

            # -- don't forget the "tail" string --
            yield from accum.pop()

        return list(iter_items())

    def insert_comment_range_end_and_reference_below(self, comment_id: int) -> None:
        """Insert a `w:commentRangeEnd` and `w:commentReference` element after this run.

        The `w:commentRangeEnd` element is the immediate sibling of this `w:r` and is followed by
        a `w:r` containing the `w:commentReference` element.
        """
        self.addnext(self._new_comment_reference_run(comment_id))
        self.addnext(OxmlElement("w:commentRangeEnd", attrs={qn("w:id"): str(comment_id)}))

    def insert_comment_range_start_above(self, comment_id: int) -> None:
        """Insert a `w:commentRangeStart` element with `comment_id` before this run."""
        self.addprevious(OxmlElement("w:commentRangeStart", attrs={qn("w:id"): str(comment_id)}))

    @property
    def lastRenderedPageBreaks(self) -> List[CT_LastRenderedPageBreak]:
        """All `w:lastRenderedPageBreaks` descendants of this run."""
        return self.xpath("./w:lastRenderedPageBreak")

    def remove_br_tag_childrens(self) -> None:
        """Convert each `w:br` child of this run into a `w:t` containing a single space.

        Used to collapse line breaks to whitespace when normalizing run text.
        """
        for child in self:
            if child.tag == qn("w:br"):
                child.tag = qn("w:t")
                child.text = " "

    @property
    def style(self) -> str | None:
        """String contained in `w:val` attribute of `w:rStyle` grandchild.

        |None| if that element is not present.
        """
        rPr = self.rPr
        if rPr is None:
            return None
        return rPr.style

    @style.setter
    def style(self, style: str | None):
        """Set character style of this `w:r` element to `style`.

        If `style` is None, remove the style element.
        """
        rPr = self.get_or_add_rPr()
        rPr.style = style

    @property
    def text(self) -> str:
        """The textual content of this run.

        Inner-content child elements like `w:tab` are translated to their text
        equivalent.
        """
        return "".join(
            str(e)
            for e in self.xpath("w:br | w:cr | w:noBreakHyphen | w:ptab | w:sym | w:t | w:tab")
        )

    @text.setter
    def text(self, text: str):  # pyright: ignore[reportIncompatibleMethodOverride]
        self.clear_content()
        _RunContentAppender.append_to_run_from_text(self, text)

    def _insert_rPr(self, rPr: CT_RPr) -> CT_RPr:
        self.insert(0, rPr)
        return rPr

    def _new_comment_reference_run(self, comment_id: int) -> CT_R:
        """Return a new `w:r` element with `w:commentReference` referencing `comment_id`.

        Should look like this:

            <w:r>
              <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
              <w:commentReference w:id="0"/>
            </w:r>

        """
        r = cast(CT_R, OxmlElement("w:r"))
        rPr = r.get_or_add_rPr()
        rPr.style = "CommentReference"
        r.append(OxmlElement("w:commentReference", attrs={qn("w:id"): str(comment_id)}))
        return r


# ------------------------------------------------------------------------------------
# Run inner-content elements


class CT_Br(BaseOxmlElement):
    """`<w:br>` element, indicating a line, page, or column break in a run."""

    type: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:type", ST_BrType, default="textWrapping"
    )
    clear: str | None = OptionalAttribute("w:clear", ST_BrClear)  # pyright: ignore

    def __str__(self) -> str:
        """Text equivalent of this element. Actual value depends on break type.

        A line break is translated as "\n". Column and page breaks produce the empty
        string ("").

        This allows the text of run inner-content to be accessed in a consistent way
        for all run inner-context text elements.
        """
        return "\n" if self.type == "textWrapping" else ""


class CT_Cr(BaseOxmlElement):
    """`<w:cr>` element, representing a carriage-return (0x0D) character within a run.

    In Word, this represents a "soft carriage-return" in the sense that it does not end
    the paragraph the way pressing Enter (aka. Return) on the keyboard does. Here the
    text equivalent is considered to be newline ("\n") since in plain-text that's the
    closest Python equivalent.

    NOTE: this complex-type name does not exist in the schema, where `w:tab` maps to
    `CT_Empty`. This name was added to give it distinguished behavior. CT_Empty is used
    for many elements.
    """

    def __str__(self) -> str:
        """Text equivalent of this element, a single newline ("\n")."""
        return "\n"


class CT_FldChar(BaseOxmlElement):
    """`<w:fldChar>` element, marking a boundary in a complex-field sequence.

    A complex field (table of contents, cross-reference, computed date, and the
    like) is represented as a `begin`/`separate`/`end` sequence of these elements
    spread across sibling runs, bracketing the field-code and cached-result content.
    """

    fldCharType: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:fldCharType", ST_FldCharType
    )


class CT_NoBreakHyphen(BaseOxmlElement):
    """`<w:noBreakHyphen>` element, a hyphen ineligible for a line-wrap position.

    This maps to a plain-text dash ("-").

    NOTE: this complex-type name does not exist in the schema, where `w:noBreakHyphen`
    maps to `CT_Empty`. This name was added to give it behavior distinguished from the
    many other elements represented in the schema by CT_Empty.
    """

    def __str__(self) -> str:
        """Text equivalent of this element, a single dash character ("-")."""
        return "-"


class CT_PTab(BaseOxmlElement):
    """`<w:ptab>` element, representing an absolute-position tab character within a run.

    This character advances the rendering position to the specified position regardless
    of any tab-stops, perhaps for layout of a table-of-contents (TOC) or similar.
    """

    def __str__(self) -> str:
        """Text equivalent of this element, a single tab ("\t") character.

        This allows the text of run inner-content to be accessed in a consistent way
        for all run inner-context text elements.
        """
        return "\t"


# -- CT_Tab functionality is provided by CT_TabStop which also uses `w:tab` tag. That
# -- element class provides the __str__() method for this empty element, unconditionally
# -- returning "\t".


class CT_Text(BaseOxmlElement):
    """`<w:t>` element, containing a sequence of characters within a run."""

    def __str__(self) -> str:
        """Text contained in this element, the empty string if it has no content.

        This property allows this run inner-content element to be queried for its text
        the same way as other run-content elements are. In particular, this never
        returns None, as etree._Element does when there is no content.
        """
        return self.text or ""


# ------------------------------------------------------------------------------------
# Utility


class _RunContentAppender:
    """Translates a Python string into run content elements appended in a `w:r` element.

    Contiguous sequences of regular characters are appended in a single `<w:t>` element.
    Each tab character ('\t') causes a `<w:tab/>` element to be appended. Likewise a
    newline or carriage return character ('\n', '\r') causes a `<w:cr>` element to be
    appended.
    """

    def __init__(self, r: CT_R):
        self._r = r
        self._bfr: List[str] = []

    @classmethod
    def append_to_run_from_text(cls, r: CT_R, text: str):
        """Append inner-content elements for `text` to `r` element."""
        appender = cls(r)
        appender.add_text(text)

    def add_text(self, text: str):
        """Append inner-content elements for `text` to the `w:r` element."""
        for char in text:
            self.add_char(char)
        self.flush()

    def add_char(self, char: str):
        """Process next character of input through finite state maching (FSM).

        There are two possible states, buffer pending and not pending, but those are
        hidden behind the `.flush()` method which must be called at the end of text to
        ensure any pending `<w:t>` element is written.
        """
        if char == "\t":
            self.flush()
            self._r.add_tab()
        elif char in "\r\n":
            self.flush()
            self._r.add_br()
        else:
            self._bfr.append(char)

    def flush(self):
        text = "".join(self._bfr)
        if text:
            self._r.add_t(text)
        self._bfr.clear()
