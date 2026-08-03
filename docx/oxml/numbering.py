# encoding: utf-8

"""
Custom element classes related to the numbering part
"""
import re
import math
from roman import toRoman

from .text.parfmt import CT_PPr
from . import OxmlElement
from .shared import CT_DecimalNumber
from .simpletypes import ST_DecimalNumber
from .xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrMore, ZeroOrOne
)
from .ns import nsmap, qn
from .text.paragraph import CT_P

_w_p_tag = qn('w:p')


class CT_Num(BaseOxmlElement):
    """
    ``<w:num>`` element, which represents a concrete list definition
    instance, having a required child <w:abstractNumId> that references an
    abstract numbering definition that defines most of the formatting details.
    """
    abstractNumId = OneAndOnlyOne('w:abstractNumId')
    lvlOverride = ZeroOrMore('w:lvlOverride')
    numId = RequiredAttribute('w:numId', ST_DecimalNumber)

    def add_lvlOverride(self, ilvl):
        """
        Return a newly added CT_NumLvl (<w:lvlOverride>) element having its
        ``ilvl`` attribute set to *ilvl*.
        """
        return self._add_lvlOverride(ilvl=ilvl)

    @classmethod
    def new(cls, num_id, abstractNum_id):
        """
        Return a new ``<w:num>`` element having numId of *num_id* and having
        a ``<w:abstractNumId>`` child with val attribute set to
        *abstractNum_id*.
        """
        num = OxmlElement('w:num')
        num.numId = num_id
        abstractNumId = CT_DecimalNumber.new(
            'w:abstractNumId', abstractNum_id
        )
        num.append(abstractNumId)
        return num


class CT_NumLvl(BaseOxmlElement):
    """
    ``<w:lvlOverride>`` element, which identifies a level in a list
    definition to override with settings it contains.
    """
    startOverride = ZeroOrOne('w:startOverride', successors=('w:lvl',))
    ilvl = RequiredAttribute('w:ilvl', ST_DecimalNumber)

    def add_startOverride(self, val):
        """
        Return a newly added CT_DecimalNumber element having tagname
        ``w:startOverride`` and ``val`` attribute set to *val*.
        """
        return self._add_startOverride(val=val)


class CT_NumPr(BaseOxmlElement):
    """
    A ``<w:numPr>`` element, a container for numbering properties applied to
    a paragraph.
    """
    ilvl = ZeroOrOne('w:ilvl', successors=(
        'w:numId', 'w:numberingChange', 'w:ins'
    ))
    numId = ZeroOrOne('w:numId', successors=('w:numberingChange', 'w:ins'))

    # @ilvl.setter
    # def _set_ilvl(self, val):
    #     """
    #     Get or add a <w:ilvl> child and set its ``w:val`` attribute to *val*.
    #     """
    #     ilvl = self.get_or_add_ilvl()
    #     ilvl.val = val

    # @numId.setter
    # def numId(self, val):
    #     """
    #     Get or add a <w:numId> child and set its ``w:val`` attribute to
    #     *val*.
    #     """
    #     numId = self.get_or_add_numId()
    #     numId.val = val


class CT_Numbering(BaseOxmlElement):
    """
    ``<w:numbering>`` element, the root element of a numbering part, i.e.
    numbering.xml
    """
    abstractNum = ZeroOrMore('w:abstractNum', successors=('w:num', 'w:numIdMacAtCleanup'))
    num = ZeroOrMore('w:num', successors=('w:numIdMacAtCleanup',))

    fmt_map = {
        'lowerLetter': lambda num: (chr(96 + (num % 26 if num % 26 != 0 else 26))
                                    * math.ceil(num / 26)),
        'decimal': lambda num: num,
        'upperLetter': lambda num: (chr(64 + (num % 26 if num % 26 != 0 else 26))
                                    * math.ceil(num / 26)),
        'lowerRoman': lambda num: toRoman(num).lower(),
        'upperRoman': lambda num: toRoman(num),
        'none': lambda num: '',
    }

    # xpath_options = {
    #     True: {'single': 'count(w:lvl)=1 and ', 'level': 0},
    #     False: {'single': '', 'level': level},
    # }

    def add_num(self, abstractNum_id):
        """
        Return a newly added CT_Num (<w:num>) element referencing the
        abstract numbering definition identified by *abstractNum_id*.
        """
        next_num_id = self._next_numId
        num = CT_Num.new(next_num_id, abstractNum_id)
        self._invalidate_num_caches()
        return self._insert_num(num)

    @property
    def _num_caches(self):
        """
        Per-instance lookup caches: ``abstractNum`` maps numId to its
        |CT_AbstractNum| (or None), ``startOverride`` maps (numId, ilvl) to
        the startOverride value, ``para_props`` maps a paragraph element to
        its resolved (ilvl, numId) — or None where resolution raises
        AttributeError — and ``para_pStyle`` maps a paragraph element to its
        ``pPr/pStyle`` element. Numbering definitions, paragraph properties
        and styles are treated as immutable during a document read (the same
        assumption Paragraph._number and the text cache already make);
        ``add_num`` and ``set_li_lvl`` invalidate. Holding paragraph elements
        as keys keeps their lxml proxies alive, which keeps identity-based
        lookups stable. The cache itself lives on this element's proxy, so it
        survives only while the element is referenced from Python (e.g. via
        NumberingPart) and is silently rebuilt otherwise.
        """
        try:
            return self._num_caches_dict
        except AttributeError:
            caches = {
                'abstractNum': {}, 'startOverride': {},
                'para_props': {}, 'para_pStyle': {},
            }
            self._num_caches_dict = caches
            return caches

    def _invalidate_num_caches(self):
        try:
            del self._num_caches_dict
        except AttributeError:
            pass

    def get_abstractNum(self, numId):
        """
        Returns |CT_AbstractNum| instance with corresponding
        paragraph ``pPr.numPr.numId`` if any
        """
        cache = self._num_caches['abstractNum']
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

    def get_startOverride(self, numId, ilvl):
        """
        Returns the ``w:startOverride`` value of the ``<w:num>`` having
        *numId* for level *ilvl*, or 0 if there is none. Raises |KeyError|
        if no ``<w:num>`` has *numId*, |AttributeError| if the matching
        ``<w:lvlOverride>`` has no ``<w:startOverride>`` child.
        """
        cache = self._num_caches['startOverride']
        key = (numId, ilvl)
        try:
            return cache[key]
        except KeyError:
            pass

        val = 0
        w_num = self.num_having_numId(numId)
        for lvlOverride in w_num.lvlOverride_lst:
            if lvlOverride.ilvl == ilvl:
                val = lvlOverride.startOverride.val
                break

        cache[key] = val
        return val

    def get_lvl_from_props(self, p, styles_cache=None):
        """
        Gets the formatting based on current paragraph indentation level defined in paragraph styles.
        If ``styles_cache`` is not None then level from style formating is fetched otherwise
        level from direct paragraph formating is used.
        """
        try:
            numPr  = p.pPr.get_style_numPr(styles_cache) if styles_cache else p.pPr.numPr

            ilvl, numId = numPr.ilvl, numPr.numId.val
            ilvl = ilvl.val if ilvl is not None else 0
            abstractNum_el = self.get_abstractNum(numId)
            return abstractNum_el.get_lvl(ilvl)
        except AttributeError:
            return None

    def get_num_for_p(self, p, styles_cache, append_suffix=True):
        """
        Returns list item for the given paragraph.
        The ``append_suffix`` flag is used to indicate the appending of
        numbered level suffix after the number
        (usually blank spaces between the number and the text of numbered paragraph).
        """

        para_props_cache = self._num_caches['para_props']
        para_pStyle_cache = self._num_caches['para_pStyle']
        _no_pPr = para_pStyle_cache  # sentinel distinct from any pStyle element

        def get_ilvl_and_numId(paragraph):
            """
            Return ``ilvl`` and ``numId`` for the given ``paragraph``.
            Raises AttributeError if the paragraph has no resolvable
            numbering properties, as the un-cached lookup did.
            """
            try:
                props = para_props_cache[paragraph]
            except KeyError:
                try:
                    para_numPr = paragraph.pPr.get_numPr(styles_cache)
                    para_ilvl, para_numId = para_numPr.ilvl, para_numPr.numId.val
                    para_ilvl = para_ilvl.val if para_ilvl is not None else 0
                    props = (para_ilvl, para_numId)
                except AttributeError:
                    props = None
                para_props_cache[paragraph] = props
            if props is None:
                raise AttributeError('paragraph has no numbering properties')
            return props

        def get_pStyle(paragraph):
            """
            Return the ``pPr/pStyle`` element of *paragraph*, or None if pPr
            has no pStyle. Raises AttributeError if the paragraph has no pPr,
            as the direct ``paragraph.pPr.pStyle`` access did.
            """
            try:
                pStyle = para_pStyle_cache[paragraph]
            except KeyError:
                pPr = paragraph.pPr
                pStyle = _no_pPr if pPr is None else pPr.pStyle
                para_pStyle_cache[paragraph] = pStyle
            if pStyle is _no_pPr:
                raise AttributeError('paragraph has no pPr')
            return pStyle

        def iter_preceding_paragraphs(p):
            """
            Yield all paragraphs preceding *p* in document order (reversed),
            regardless of nesting (table cells, rows, etc.).

            Walks up the XML tree level by level. At each level, iterates
            preceding siblings and yields any ``<w:p>`` elements found —
            either directly or nested inside the sibling's descendants.
            This correctly counts numbered paragraphs across table cells,
            rows, and between body-level and table-internal contexts.
            """
            current = p
            parent = current.getparent()
            while parent is not None:
                for sibling in current.itersiblings(preceding=True):
                    if isinstance(sibling, CT_P):
                        yield sibling
                    else:
                        # Yield all w:p descendants in reverse document order.
                        # Use iterdescendants with Clark-notation tag to avoid
                        # xpath namespace issues on non-BaseOxmlElement nodes.
                        paras = list(sibling.iterdescendants(_w_p_tag))
                        for sp in reversed(paras):
                            yield sp
                current = parent
                parent = current.getparent()

        def same_abstract_num(numId_a, numId_b):
            """
            Two numIds belong to the same numbering family if they resolve to
            the same abstractNumId AND neither carries a startOverride for the
            current level.  A startOverride signals an intentional restart —
            a new list instance from the same template — so it must not be
            counted as a continuation.
            """
            if numId_a == numId_b:
                return True
            abs_a = self.get_abstractNum(numId_a)
            abs_b = self.get_abstractNum(numId_b)
            if abs_a is None or abs_a is not abs_b:
                return False
            # A startOverride on either num means "new list instance";
            # treat as separate even though they share an abstract definition.
            if get_start_override(numId_a) or get_start_override(numId_b):
                return False
            return True

        def get_preceding_paragraphs_numIds(p, p_ilvl, p_numId):
            """
            Return preceding siblings ``numId`` that are in the same numbered list as the paragraph ``p``.
            Paragraphs are in the same list if they are on the same level (``p_ilvl``), and either
            have the same ``p_numId`` or resolve to the same abstract numbering definition.
            Skips unnumbered paragraphs within the numbering list.
            Stops on the paragraph that is on the lower level.
            """
            pStyle = p.pPr.pStyle
            for prev_p in iter_preceding_paragraphs(p):
                try:
                    prev_p_ilvl, prev_p_numId = get_ilvl_and_numId(prev_p)
                    # skip unnumbered paragraphs within numbering list
                    if prev_p_numId == 0:
                        continue
                    prev_p_pStyle = get_pStyle(prev_p)
                    if prev_p_ilvl < p_ilvl and (prev_p_numId == p_numId or
                                         (prev_p_pStyle is not None and prev_p_pStyle.val in linked_styles)):
                        break
                    if prev_p_ilvl == p_ilvl and (prev_p_numId == p_numId or same_abstract_num(prev_p_numId, p_numId) or prev_p_pStyle.val in linked_styles):
                        yield prev_p_numId
                    # para `p` that has only style defined which is same as the `prev_p` style
                    # should be counted even though they have different `numId`s.
                    if prev_p_ilvl == p_ilvl and prev_p_numId != p_numId:
                        if p.pPr.numPr is None and get_pStyle(prev_p).val == pStyle.val:
                            startOverride = get_start_override(prev_p_numId)
                            if startOverride > 1:
                                yield prev_p_numId
                            else:
                                yield p_numId
                            break
                except AttributeError:
                    continue

        def count_same_numIds(preceding_paragraphs_numIds, numId, num):
            """
            Add 1 to ``num`` for each preceding paragraph that belongs to
            the same numbering family — same ``w:numId`` or same abstract
            numbering definition.

            On a paragraph from a different list, behavior splits:

            * If our list carries its own ``startOverride > 1``, that is an
              explicit "new list instance" — stop counting and leave
              ``num`` at our own start/startOverride. Word treats those
              as independent counters, so an unrelated list's
              ``startOverride`` must not bleed into ours.
            * If only the preceding list carries ``startOverride > 1``,
              the document is using that list's reset as the continuation
              point for ours (typical for style-driven section sequences
              that re-key the numId at each section). Add the
              ``startOverride`` to advance ``num`` to match.
            * Otherwise the unrelated paragraph is an interleaved short
              list with no real reset — skip past it and inspect the next
              yield, so we can continue our own list across the gap.
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

        def get_start_override(for_numId):
            return self.get_startOverride(for_numId, ilvl)

        ilvl, numId = get_ilvl_and_numId(p)

        abstractNum_el = self.get_abstractNum(numId)
        if abstractNum_el is None:
            return None
        lvl_el = abstractNum_el.get_lvl(ilvl)
        linked_styles = {s.xpath('w:pStyle/@w:val')[0]
            for s in lvl_el.xpath('preceding-sibling::w:lvl[w:pStyle]')}

        startOverride = get_start_override(numId)
        try:
            start = int(lvl_el.start.get('{%s}val' % nsmap['w']))
        except AttributeError:
            # default start value
            start = 0

        p_num = startOverride if startOverride else start

        preceding_paragraphs_numIds = get_preceding_paragraphs_numIds(p, ilvl, numId)
        p_num = count_same_numIds(preceding_paragraphs_numIds, numId, p_num)

        try:
            # apply numbering style
            p_num_str = self.fmt_map[lvl_el.numFmt.get('{%s}val' % nsmap['w'])](p_num)
        except KeyError:
            return None

        suffix = ''
        if append_suffix is True:
            suffix = lvl_el.suffix
        lvlText = lvl_el.lvlText.get('{%s}val' % nsmap['w'])
        if lvlText.count('%') <= 1:
            return re.sub(r'%(\d)', str(p_num_str), lvlText, 1) + suffix

        # Multi-component lvlText (e.g. '%1.%2.%3'). Each %N references the
        # counter at ilvl=N-1 in the same numbering family. Compute each
        # component independently from per-level counts and <w:start>
        # values — that is how Word encodes parent-level context (e.g. a
        # chapter number stored as the start value at ilvl=1).

        def value_at_ilvl(target_ilvl):
            if target_ilvl == ilvl:
                return p_num
            target_lvl_el = abstractNum_el.get_lvl(target_ilvl)
            if target_lvl_el is None:
                return 1
            try:
                target_start = int(target_lvl_el.start.get('{%s}val' % nsmap['w']))
            except AttributeError:
                target_start = 1
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
                        or (prev_p_pStyle is not None
                            and prev_p_pStyle.val in linked_styles)
                    )
                    if not same_list:
                        continue
                    if prev_p_ilvl < target_ilvl:
                        # paragraph at lower level resets the target counter
                        break
                    if prev_p_ilvl == target_ilvl:
                        count += 1
                except AttributeError:
                    continue
            # No preceding paragraphs at target_ilvl: the counter is at its
            # initial value. With N preceding, the latest emitted value is
            # base + N - 1.
            if count == 0:
                return base
            return base + count - 1

        # All %N substitutions inside a single lvlText share the current
        # level's numFmt. Word does not apply the referenced level's own
        # format — e.g. an ilvl=0 with upperRoman seen via %1 from an
        # ilvl=1 lvlText (decimal) renders as decimal "2", not "II".
        cur_numFmt = lvl_el.numFmt.get('{%s}val' % nsmap['w'])
        cur_formatter = self.fmt_map[cur_numFmt]

        def replace_token(m):
            n = int(m.group(1))
            return str(cur_formatter(value_at_ilvl(n - 1)))

        return re.sub(r'%(\d)', replace_token, lvlText) + suffix

    def num_having_numId(self, numId):
        """
        Return the ``<w:num>`` child element having ``numId`` attribute
        matching *numId*.
        """
        xpath = './w:num[@w:numId="%d"]' % numId
        try:
            return self.xpath(xpath)[0]
        except IndexError:
            raise KeyError('no <w:num> element with numId %d' % numId)

    @property
    def _next_numId(self):
        """
        The first ``numId`` unused by a ``<w:num>`` element, starting at
        1 and filling any gaps in numbering between existing ``<w:num>``
        elements.
        """
        numId_strs = self.xpath('./w:num/@w:numId')
        num_ids = [int(numId_str) for numId_str in numId_strs]
        for num in range(1, len(num_ids)+2):
            if num not in num_ids:
                break
        return num

    def set_li_lvl(self, para_el, styles, prev_p, ilvl):
        """
        Sets paragraph list item indentation level. When previous
        paragraph ``prev_p`` is specified, it will look up for existing numbering
        list of ``prev_p`` and add new list item. If no ``prev_p`` is specified,
        it will create a new numbering list with specified indentation level ``ilvl``.
        """
        if (prev_p is None or
                prev_p.pPr is None or
                prev_p.pPr.numPr is None or
                prev_p.pPr.numPr.numId is None):
            if ilvl is None:
                ilvl = 0
            numPr = para_el.pPr.get_numPr(styles)
            if numPr is None:
                return
            numId = numPr.numId.val
            num_el = self.num_having_numId(numId)
            anum = num_el.abstractNumId.val
            num = self.add_num(anum)
            num.add_lvlOverride(ilvl=ilvl).add_startOverride(1)
            num = num.numId
        else:
            if ilvl is None:
                ilvl = prev_p.pPr.numPr.ilvl.val
            num = prev_p.pPr.numPr.numId.val
        para_el.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = num
        para_el.get_or_add_pPr().get_or_add_numPr().get_or_add_ilvl().val = ilvl
        self._invalidate_num_caches()

class CT_AbstractNum(BaseOxmlElement):
    """
    ``<w:abstractNum>`` element, contains definitions for numbering part.
    """
    abstractNumId = RequiredAttribute('w:abstractNumId', ST_DecimalNumber)
    lvl = ZeroOrMore('w:lvl')

    def get_lvl(self, ilvl):
        """
        Returns |CT_Lvl| instance with corresponding ``ilvl`` if any
        """
        for el in self.lvl_lst:
            if el.ilvl == ilvl:
                return el


class CT_Lvl(BaseOxmlElement):
    """
    ``<w:lvl>`` element located within ``<w:abstractNum>`` describing
    list item formatting
    """
    ilvl = RequiredAttribute('w:ilvl', ST_DecimalNumber)
    start = ZeroOrOne('w:start', CT_DecimalNumber)
    pPr = ZeroOrOne('w:pPr', CT_PPr)
    numFmt = ZeroOrOne('w:numFmt')
    lvlText = ZeroOrOne('w:lvlText')
    suff = ZeroOrOne('w:suff')

    @property
    def suffix(self):
        if self.suff is not None:
            if self.suff.get('{%s}val' % nsmap['w']) == 'space':
                return ' '
            else:
                return ''
        else:
            return '\t'
