"""Custom element classes related to the numbering part."""

from __future__ import annotations

import contextlib
import math
import re
from typing import TYPE_CHECKING, Any, Callable, Iterator, List, cast

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.shared import CT_DecimalNumber
from docx.oxml.simpletypes import ST_DecimalNumber
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from docx.oxml.shared import CT_String
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.text.parfmt import CT_PPr

_w_p_tag = qn("w:p")

_ROMAN_NUMERALS = (
    (1000, "M"),
    (900, "CM"),
    (500, "D"),
    (400, "CD"),
    (100, "C"),
    (90, "XC"),
    (50, "L"),
    (40, "XL"),
    (10, "X"),
    (9, "IX"),
    (5, "V"),
    (4, "IV"),
    (1, "I"),
)


def _to_upper_roman(num: int) -> str:
    """`num` (a positive integer) as an uppercase Roman numeral, e.g. `12` -> `"XII"`."""
    parts: List[str] = []
    for value, symbol in _ROMAN_NUMERALS:
        count, num = divmod(num, value)
        parts.append(symbol * count)
    return "".join(parts)


def _to_letters(num: int, base_char: str) -> str:
    """`num` (a positive integer) as a repeating-letter label, e.g. with `base_char="a"`:
    `1`->`"a"`, `26`->`"z"`, `27`->`"aa"`, `52`->`"zz"`, `53`->`"aaa"`.

    This is the numbering scheme Word actually uses past the 26th item -- repeating the
    26th letter's position rather than incrementing like a base-26 number (so `27` is
    `"aa"`, not `"ab"`).
    """
    letter_idx = num % 26 if num % 26 != 0 else 26
    return chr(ord(base_char) + letter_idx - 1) * math.ceil(num / 26)


class CT_Num(BaseOxmlElement):
    """``<w:num>`` element, which represents a concrete list definition instance, having
    a required child <w:abstractNumId> that references an abstract numbering definition
    that defines most of the formatting details."""

    _add_lvlOverride: Callable[..., CT_NumLvl]
    lvlOverride_lst: List[CT_NumLvl]

    abstractNumId: CT_DecimalNumber = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "w:abstractNumId"
    )
    lvlOverride = ZeroOrMore("w:lvlOverride")
    numId: int = RequiredAttribute("w:numId", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]

    def add_lvlOverride(self, ilvl: int) -> CT_NumLvl:
        """Return a newly added CT_NumLvl (<w:lvlOverride>) element having its ``ilvl``
        attribute set to `ilvl`."""
        return self._add_lvlOverride(ilvl=ilvl)

    @classmethod
    def new(cls, num_id: int, abstractNum_id: int) -> CT_Num:
        """Return a new ``<w:num>`` element having numId of `num_id` and having a
        ``<w:abstractNumId>`` child with val attribute set to `abstractNum_id`."""
        num = cast("CT_Num", OxmlElement("w:num"))
        num.numId = num_id
        abstractNumId = CT_DecimalNumber.new("w:abstractNumId", abstractNum_id)
        num.append(abstractNumId)
        return num


class CT_NumLvl(BaseOxmlElement):
    """``<w:lvlOverride>`` element, which identifies a level in a list definition to
    override with settings it contains."""

    _add_startOverride: Callable[..., CT_DecimalNumber]

    startOverride: CT_DecimalNumber | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:startOverride", successors=("w:lvl",)
    )
    ilvl: int = RequiredAttribute("w:ilvl", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]

    def add_startOverride(self, val: int) -> CT_DecimalNumber:
        """Return a newly added CT_DecimalNumber element having tagname
        ``w:startOverride`` and ``val`` attribute set to `val`."""
        return self._add_startOverride(val=val)


class CT_NumPr(BaseOxmlElement):
    """A ``<w:numPr>`` element, a container for numbering properties applied to a
    paragraph."""

    get_or_add_ilvl: Callable[[], CT_DecimalNumber]
    get_or_add_numId: Callable[[], CT_DecimalNumber]

    ilvl: CT_DecimalNumber | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:ilvl", successors=("w:numId", "w:numberingChange", "w:ins")
    )
    numId: CT_DecimalNumber | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:numId", successors=("w:numberingChange", "w:ins")
    )


class CT_AbstractNum(BaseOxmlElement):
    """``<w:abstractNum>`` element, containing the level definitions for one numbering
    scheme."""

    lvl_lst: List[CT_Lvl]

    abstractNumId: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:abstractNumId", ST_DecimalNumber
    )
    lvl = ZeroOrMore("w:lvl")

    def get_lvl(self, ilvl: int) -> CT_Lvl | None:
        """The ``<w:lvl>`` child having ``@w:ilvl`` equal to `ilvl`, or |None| if not
        present."""
        for lvl in self.lvl_lst:
            if lvl.ilvl == ilvl:
                return lvl
        return None


class CT_Lvl(BaseOxmlElement):
    """``<w:lvl>`` element, located within ``<w:abstractNum>``, describing list-item
    formatting for one indentation level."""

    ilvl: int = RequiredAttribute("w:ilvl", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    start: CT_DecimalNumber | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:start"
    )
    pPr: CT_PPr | None = ZeroOrOne("w:pPr")  # pyright: ignore[reportAssignmentType]
    numFmt: BaseOxmlElement | None = ZeroOrOne("w:numFmt")  # pyright: ignore[reportAssignmentType]
    lvlText: BaseOxmlElement | None = ZeroOrOne("w:lvlText")  # pyright: ignore[reportAssignmentType]
    suff: BaseOxmlElement | None = ZeroOrOne("w:suff")  # pyright: ignore[reportAssignmentType]

    @property
    def suffix(self) -> str:
        """The whitespace (or absence of it) that follows this level's numbering label.

        `' '` for `@w:val="space"`, `''` for any other explicit `@w:val`, and the OOXML
        default `'\\t'` when `w:suff` is not present.
        """
        suff = self.suff
        if suff is None:
            return "\t"
        return " " if suff.get(qn("w:val")) == "space" else ""


class CT_Numbering(BaseOxmlElement):
    """``<w:numbering>`` element, the root element of a numbering part, i.e.
    numbering.xml."""

    abstractNum_lst: List[CT_AbstractNum]
    _insert_num: Callable[[CT_Num], CT_Num]

    abstractNum = ZeroOrMore("w:abstractNum", successors=("w:num", "w:numIdMacAtCleanup"))
    num = ZeroOrMore("w:num", successors=("w:numIdMacAtCleanup",))

    fmt_map: dict[str, Callable[[int], str | int]] = {
        "lowerLetter": lambda num: _to_letters(num, "a"),
        "decimal": lambda num: num,
        "upperLetter": lambda num: _to_letters(num, "A"),
        "lowerRoman": lambda num: _to_upper_roman(num).lower(),
        "upperRoman": lambda num: _to_upper_roman(num),
        "none": lambda num: "",
    }

    def add_num(self, abstractNum_id: int) -> CT_Num:
        """Return a newly added CT_Num (<w:num>) element referencing the abstract
        numbering definition identified by `abstractNum_id`."""
        next_num_id = self._next_numId
        num = CT_Num.new(next_num_id, abstractNum_id)
        self._invalidate_num_caches()
        return self._insert_num(num)

    @property
    def _num_caches(self) -> dict[str, dict[Any, Any]]:
        """Per-instance lookup caches: ``abstractNum`` maps numId to its
        |CT_AbstractNum| (or None), ``startOverride`` maps (numId, ilvl) to the
        startOverride value, ``para_props`` maps a paragraph element to its resolved
        (ilvl, numId) -- or None where resolution raises AttributeError -- and
        ``para_pStyle`` maps a paragraph element to its ``pPr/pStyle`` element.
        Numbering definitions, paragraph properties, and styles are treated as
        immutable during a document read; ``add_num`` and ``set_li_lvl`` invalidate.
        """
        try:
            return self._num_caches_dict
        except AttributeError:
            caches: dict[str, dict[Any, Any]] = {
                "abstractNum": {},
                "startOverride": {},
                "para_props": {},
                "para_pStyle": {},
            }
            self._num_caches_dict = caches
            return caches

    def _invalidate_num_caches(self) -> None:
        with contextlib.suppress(AttributeError):
            del self._num_caches_dict

    def get_abstractNum(self, numId: int) -> CT_AbstractNum | None:
        """The |CT_AbstractNum| referenced (via its `<w:num>`) by `numId`, or |None| if
        `numId` doesn't resolve to one."""
        cache = self._num_caches["abstractNum"]
        try:
            return cache[numId]
        except KeyError:
            pass

        abstractNum = None
        try:
            num_el = self.num_having_numId(numId)
        except KeyError:
            num_el = None

        if num_el is not None:
            abstractNum_id = num_el.abstractNumId.val
            for el in self.abstractNum_lst:
                if el.abstractNumId == abstractNum_id:
                    abstractNum = el
                    break

        cache[numId] = abstractNum
        return abstractNum

    def get_startOverride(self, numId: int, ilvl: int) -> int:
        """The ``w:startOverride`` value of the ``<w:num>`` having `numId` for level
        `ilvl`, or 0 if there is none.

        Raises |KeyError| if no ``<w:num>`` has `numId`.
        """
        cache = self._num_caches["startOverride"]
        key = (numId, ilvl)
        try:
            return cache[key]
        except KeyError:
            pass

        val = 0
        w_num = self.num_having_numId(numId)
        for lvlOverride in w_num.lvlOverride_lst:
            if lvlOverride.ilvl == ilvl and lvlOverride.startOverride is not None:
                val = lvlOverride.startOverride.val
                break

        cache[key] = val
        return val

    def get_lvl_from_props(
        self, p: CT_P, styles_cache: dict[str, Any] | None = None
    ) -> CT_Lvl | None:
        """The ``<w:lvl>`` formatting definition for paragraph `p`'s numbering level.

        Resolved from `p`'s own numbering properties when `styles_cache` is |None|, or
        from `p`'s paragraph style's numbering properties when `styles_cache` is
        provided. Returns |None| if `p` (or its style) specifies no numbering, or if
        that numbering doesn't resolve to a defined abstract numbering level.
        """
        try:
            pPr = p.pPr
            if pPr is None:
                return None
            numPr = (
                pPr.get_style_numPr(styles_cache)
                if styles_cache
                else cast("CT_NumPr | None", pPr.numPr)
            )
            if numPr is None or numPr.numId is None:
                return None
            ilvl_el, numId = numPr.ilvl, numPr.numId.val
            ilvl = ilvl_el.val if ilvl_el is not None else 0
            abstractNum_el = self.get_abstractNum(numId)
            if abstractNum_el is None:
                return None
            return abstractNum_el.get_lvl(ilvl)
        except AttributeError:
            return None

    def get_num_for_p(
        self, p: CT_P, styles_cache: dict[str, Any], append_suffix: bool = True
    ) -> str | None:
        """The list-item label (e.g. `"1)\\t"`, `"(a) "`) for paragraph `p`, or |None|
        if `p` is not a numbered-list paragraph, or its numbering format is not one of
        the six supported by `fmt_map` (notably `bullet` lists always return |None|).

        `append_suffix` controls whether the level's trailing whitespace (tab, space,
        or nothing) is appended after the label.
        """

        para_props_cache = self._num_caches["para_props"]
        para_pStyle_cache = self._num_caches["para_pStyle"]
        _no_pPr = para_pStyle_cache  # sentinel distinct from any pStyle element

        def get_ilvl_and_numId(paragraph: CT_P) -> tuple[int, int]:
            """(ilvl, numId) for `paragraph`.

            Raises AttributeError if `paragraph` has no resolvable numbering
            properties, matching the uncached lookup's behavior.
            """
            try:
                props = para_props_cache[paragraph]
            except KeyError:
                try:
                    pPr = paragraph.pPr
                    if pPr is None:
                        raise AttributeError
                    para_numPr = pPr.get_numPr(styles_cache)
                    if para_numPr is None or para_numPr.numId is None:
                        raise AttributeError
                    para_ilvl_el, para_numId = para_numPr.ilvl, para_numPr.numId.val
                    para_ilvl = para_ilvl_el.val if para_ilvl_el is not None else 0
                    props = (para_ilvl, para_numId)
                except AttributeError:
                    props = None
                para_props_cache[paragraph] = props
            if props is None:
                raise AttributeError("paragraph has no numbering properties")
            return props

        def get_pStyle(paragraph: CT_P) -> CT_String | None:
            """The ``pPr/pStyle`` element of `paragraph`, or |None| if `pPr` has no
            `pStyle`.

            Raises AttributeError if `paragraph` has no `pPr`.
            """
            try:
                pStyle = para_pStyle_cache[paragraph]
            except KeyError:
                pPr = paragraph.pPr
                pStyle = _no_pPr if pPr is None else pPr.pStyle
                para_pStyle_cache[paragraph] = pStyle
            if pStyle is _no_pPr:
                raise AttributeError("paragraph has no pPr")
            return cast("CT_String | None", pStyle)

        def iter_preceding_paragraphs(p: CT_P) -> Iterator[CT_P]:
            """Yield all paragraphs preceding `p` in document order (reversed),
            regardless of nesting (table cells, rows, etc.).

            Walks up the XML tree level by level. At each level, iterates preceding
            siblings and yields any ``<w:p>`` elements found -- either directly or
            nested inside the sibling's descendants. This correctly counts numbered
            paragraphs across table cells, rows, and between body-level and
            table-internal contexts.
            """
            current = p
            parent = current.getparent()
            while parent is not None:
                for sibling in current.itersiblings(preceding=True):
                    if sibling.tag == _w_p_tag:
                        yield cast("CT_P", sibling)
                    else:
                        paras = list(sibling.iterdescendants(_w_p_tag))
                        for sp in reversed(paras):
                            yield cast("CT_P", sp)
                current = parent
                parent = current.getparent()

        def get_start_override(for_numId: int) -> int:
            return self.get_startOverride(for_numId, ilvl)

        def same_abstract_num(numId_a: int, numId_b: int) -> bool:
            """Two numIds belong to the same numbering family if they resolve to the
            same abstractNumId AND neither carries a startOverride for the current
            level.

            A startOverride signals an intentional restart -- a new list instance from
            the same template -- so it must not be counted as a continuation.
            """
            if numId_a == numId_b:
                return True
            abs_a = self.get_abstractNum(numId_a)
            abs_b = self.get_abstractNum(numId_b)
            if abs_a is None or abs_a is not abs_b:
                return False
            return not (get_start_override(numId_a) or get_start_override(numId_b))

        def get_preceding_paragraphs_numIds(p: CT_P, p_ilvl: int, p_numId: int) -> Iterator[int]:
            """Preceding-sibling ``numId``s that are in the same numbered list as `p`.

            Paragraphs are in the same list if they are on the same level (`p_ilvl`),
            and either have the same `p_numId` or resolve to the same abstract
            numbering definition. Skips unnumbered paragraphs within the numbering
            list. Stops on a paragraph at a lower level.
            """
            assert p.pPr is not None
            pStyle = p.pPr.pStyle
            for prev_p in iter_preceding_paragraphs(p):
                try:
                    prev_p_ilvl, prev_p_numId = get_ilvl_and_numId(prev_p)
                    # -- skip unnumbered paragraphs within the numbering list --
                    if prev_p_numId == 0:
                        continue
                    prev_p_pStyle = get_pStyle(prev_p)
                    if prev_p_ilvl < p_ilvl and (
                        prev_p_numId == p_numId
                        or (prev_p_pStyle is not None and prev_p_pStyle.val in linked_styles)
                    ):
                        break
                    if prev_p_ilvl == p_ilvl and (
                        prev_p_numId == p_numId
                        or same_abstract_num(prev_p_numId, p_numId)
                        or (prev_p_pStyle is not None and prev_p_pStyle.val in linked_styles)
                    ):
                        yield prev_p_numId
                    # -- a paragraph `p` that has only a style defined, matching
                    # -- `prev_p`'s style, should be counted even with a different
                    # -- numId.
                    if (
                        prev_p_ilvl == p_ilvl
                        and prev_p_numId != p_numId
                        and cast("CT_NumPr | None", p.pPr.numPr) is None
                        and pStyle is not None
                        and prev_p_pStyle is not None
                        and prev_p_pStyle.val == pStyle.val
                    ):
                        startOverride = get_start_override(prev_p_numId)
                        if startOverride > 1:
                            yield prev_p_numId
                        else:
                            yield p_numId
                        break
                except AttributeError:
                    continue

        def count_same_numIds(
            preceding_paragraphs_numIds: Iterator[int], numId: int, num: int
        ) -> int:
            """Add 1 to `num` for each preceding paragraph that belongs to the same
            numbering family -- same `w:numId` or same abstract numbering definition.

            On a paragraph from a different list, behavior splits:

            * If our list carries its own `startOverride > 1`, that is an explicit
              "new list instance" -- stop counting and leave `num` at our own
              start/startOverride. An unrelated list's startOverride must not bleed
              into ours.
            * If only the preceding list carries `startOverride > 1`, the document is
              using that list's reset as the continuation point for ours (typical for
              style-driven section sequences that re-key the numId at each section).
              Add the startOverride to advance `num` to match.
            * Otherwise the unrelated paragraph is an interleaved short list with no
              real reset -- skip past it and inspect the next yield, so our own list
              continues across the gap.
            """
            for p_numId in preceding_paragraphs_numIds:
                try:
                    if numId == p_numId or same_abstract_num(p_numId, numId):
                        num += 1
                        continue
                    if get_start_override(numId) > 1:
                        break
                    prev_startOverride = get_start_override(p_numId)
                    if prev_startOverride > 1:
                        num += prev_startOverride
                        break
                    try:
                        next_p_numId = next(preceding_paragraphs_numIds)
                    except StopIteration:
                        break
                    if next_p_numId == numId or same_abstract_num(next_p_numId, numId):
                        num += 1
                    break
                except AttributeError:
                    continue
            return num

        ilvl, numId = get_ilvl_and_numId(p)

        abstractNum_el = self.get_abstractNum(numId)
        if abstractNum_el is None:
            return None
        lvl_el = abstractNum_el.get_lvl(ilvl)
        if lvl_el is None or lvl_el.numFmt is None or lvl_el.lvlText is None:
            return None
        linked_styles = {
            s.xpath("w:pStyle/@w:val")[0]
            for s in lvl_el.xpath("preceding-sibling::w:lvl[w:pStyle]")
        }

        startOverride = get_start_override(numId)
        start_el = lvl_el.start
        start = int(start_el.get(qn("w:val")) or "0") if start_el is not None else 0

        p_num = startOverride if startOverride else start

        preceding_paragraphs_numIds = get_preceding_paragraphs_numIds(p, ilvl, numId)
        p_num = count_same_numIds(preceding_paragraphs_numIds, numId, p_num)

        try:
            p_num_str = self.fmt_map[lvl_el.numFmt.get(qn("w:val")) or ""](p_num)
        except KeyError:
            return None

        suffix = lvl_el.suffix if append_suffix else ""
        lvlText = lvl_el.lvlText.get(qn("w:val")) or ""
        if lvlText.count("%") <= 1:
            return re.sub(r"%(\d)", str(p_num_str), lvlText, count=1) + suffix

        # -- Multi-component lvlText (e.g. '%1.%2.%3'). Each %N references the
        # -- counter at ilvl=N-1 in the same numbering family. Compute each component
        # -- independently from per-level counts and <w:start> values -- that is how
        # -- Word encodes parent-level context (e.g. a chapter number stored as the
        # -- start value at ilvl=1).

        def value_at_ilvl(target_ilvl: int) -> int:
            if target_ilvl == ilvl:
                return p_num
            target_lvl_el = abstractNum_el.get_lvl(target_ilvl)
            if target_lvl_el is None:
                return 1
            target_start_el = target_lvl_el.start
            target_start = (
                int(target_start_el.get(qn("w:val")) or "1") if target_start_el is not None else 1
            )
            target_override = self.get_startOverride(numId, target_ilvl)
            base = target_override if target_override else target_start
            count = 0
            for prev_p in iter_preceding_paragraphs(p):
                try:
                    prev_p_ilvl, prev_p_numId = get_ilvl_and_numId(prev_p)
                    if prev_p_numId == 0:
                        continue
                    prev_p_pStyle = get_pStyle(prev_p)
                    same_list = (
                        prev_p_numId == numId
                        or same_abstract_num(prev_p_numId, numId)
                        or (prev_p_pStyle is not None and prev_p_pStyle.val in linked_styles)
                    )
                    if not same_list:
                        continue
                    if prev_p_ilvl < target_ilvl:
                        # -- a paragraph at a lower level resets the target counter --
                        break
                    if prev_p_ilvl == target_ilvl:
                        count += 1
                except AttributeError:
                    continue
            # -- with no preceding paragraphs at target_ilvl, the counter is at its
            # -- initial value; with N preceding, the latest emitted value is
            # -- base + N - 1.
            if count == 0:
                return base
            return base + count - 1

        # -- All %N substitutions inside a single lvlText share the current level's
        # -- numFmt. Word does not apply the referenced level's own format -- e.g. an
        # -- ilvl=0 with upperRoman seen via %1 from an ilvl=1 lvlText (decimal) renders
        # -- as decimal "2", not "II".
        assert lvl_el.numFmt is not None
        cur_numFmt = lvl_el.numFmt.get(qn("w:val")) or ""
        cur_formatter = self.fmt_map[cur_numFmt]

        def replace_token(m: re.Match[str]) -> str:
            n = int(m.group(1))
            return str(cur_formatter(value_at_ilvl(n - 1)))

        return re.sub(r"%(\d)", replace_token, lvlText) + suffix

    def num_having_numId(self, numId: int) -> CT_Num:
        """The ``<w:num>`` child element having ``numId`` attribute matching `numId`."""
        xpath = './w:num[@w:numId="%d"]' % numId
        try:
            return self.xpath(xpath)[0]
        except IndexError:
            raise KeyError("no <w:num> element with numId %d" % numId)

    @property
    def _next_numId(self) -> int:
        """The first ``numId`` unused by a ``<w:num>`` element, starting at 1 and
        filling any gaps in numbering between existing ``<w:num>`` elements."""
        numId_strs = self.xpath("./w:num/@w:numId")
        num_ids = [int(numId_str) for numId_str in numId_strs]
        num = 1
        for num in range(1, len(num_ids) + 2):
            if num not in num_ids:
                break
        return num

    def set_li_lvl(
        self, para_el: CT_P, styles_cache: dict[str, Any], prev_p: CT_P | None, ilvl: int | None
    ) -> None:
        """Set the list-item indentation level of `para_el`.

        When `prev_p` is specified, looks up the existing numbering list `prev_p`
        belongs to and adds `para_el` to it as the next item. When `prev_p` is |None|,
        starts a new numbering list at indentation level `ilvl`.
        """
        prev_numPr = (
            cast("CT_NumPr | None", prev_p.pPr.numPr)
            if prev_p is not None and prev_p.pPr is not None
            else None
        )
        if prev_numPr is None or prev_numPr.numId is None:
            if ilvl is None:
                ilvl = 0
            pPr = para_el.pPr
            numPr = pPr.get_numPr(styles_cache) if pPr is not None else None
            if numPr is None or numPr.numId is None:
                return
            numId = numPr.numId.val
            num_el = self.num_having_numId(numId)
            anum = num_el.abstractNumId.val
            num = self.add_num(anum)
            num.add_lvlOverride(ilvl=ilvl).add_startOverride(1)
            new_numId = num.numId
        else:
            if ilvl is None:
                assert prev_numPr.ilvl is not None
                ilvl = prev_numPr.ilvl.val
            new_numId = prev_numPr.numId.val
        para_el.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = new_numId
        para_el.get_or_add_pPr().get_or_add_numPr().get_or_add_ilvl().val = ilvl
        self._invalidate_num_caches()
