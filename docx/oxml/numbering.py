# encoding: utf-8

"""
Custom element classes related to the numbering part
"""
import re
import math
from roman import RomanError, toRoman

from .text.parfmt import CT_PPr
from . import OxmlElement
from .shared import CT_DecimalNumber
from .simpletypes import ST_DecimalNumber
from .xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrMore, ZeroOrOne
)
from .ns import nsmap, qn

_w_p_tag = qn('w:p')
_w_txbxContent_tag = qn('w:txbxContent')


def _story_of(p):
    """
    Return the element whose paragraphs Word numbers as one sequence with
    *p*: the text box that holds it, or else the root of its tree.
    """
    story = p
    for ancestor in p.iterancestors():
        story = ancestor
        if ancestor.tag == _w_txbxContent_tag:
            break
    return story


def _start_value(lvl_el):
    """The ``w:start`` value of *lvl_el*, or 0 (the default) if it has none."""
    try:
        return int(lvl_el.start.get('{%s}val' % nsmap['w']))
    except AttributeError:
        return 0


def _restart_level(lvl_el, ilvl):
    """
    An item at a level lower than the returned level restarts *lvl_el*. By
    default that is any level above it; ``w:lvlRestart`` 0 means never.
    """
    restart = lvl_el.find(qn('w:lvlRestart')) if lvl_el is not None else None
    if restart is None:
        return ilvl
    return int(restart.get('{%s}val' % nsmap['w']))


def _is_on(el):
    """
    True if the on/off element *el* is there and its ``w:val`` does not
    turn it off.
    """
    return el is not None and el.get(qn('w:val')) not in ('0', 'false', 'off')


def _style_numPrs(style_id, styles_cache):
    """
    Yield the ``w:numPr`` of the style having *style_id* and of the styles
    it is based on, nearest first.
    """
    seen = set()
    while style_id is not None and style_id not in seen:
        seen.add(style_id)
        style = styles_cache.get(style_id)
        if style is None:
            return
        if style.pPr is not None and style.pPr.numPr is not None:
            yield style.pPr.numPr
        style_id = style.basedOn_val


def _ilvl_and_numId(numPrs):
    """
    Return ``(ilvl, numId)`` from the first of *numPrs* that sets each, or
    None if none sets a numId other than 0.
    """
    ilvl = numId = None
    for numPr in numPrs:
        if numId is None and numPr.numId is not None:
            numId = numPr.numId.val
        if ilvl is None and numPr.ilvl is not None:
            ilvl = numPr.ilvl.val
    if not numId:
        return None
    return (ilvl or 0), numId


def _format_value(formatter, value):
    try:
        return str(formatter(value))
    except RomanError:
        # roman numerals have no 0 or negative values
        return str(value)


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
        'decimalZero': lambda num: '%02d' % num,
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
        |CT_AbstractNum| (or None), ``startOverride`` and ``start_override``
        map (numId, ilvl) to the startOverride value (0 or None if there is
        none), ``lvl`` maps (numId, ilvl) to the |CT_Lvl| that applies, and
        ``labels`` maps a story element to the list items of its paragraphs
        (see ``_count_list_items``). Numbering definitions,
        paragraph properties and styles are treated as immutable during a
        document read (the same assumption Paragraph._number and the text
        cache already make); ``add_num`` and ``set_li_lvl`` invalidate, and
        a change to the paragraphs of a story forgets its list items (see
        ``forget_list_items``).
        Holding paragraph elements as keys keeps their lxml proxies alive,
        which keeps identity-based lookups stable. The cache itself lives on
        this element's proxy, so it survives only while the element is
        referenced from Python (e.g. via NumberingPart) and is silently
        rebuilt otherwise.
        """
        try:
            return self._num_caches_dict
        except AttributeError:
            caches = {
                'abstractNum': {}, 'startOverride': {},
                'start_override': {}, 'labels': {}, 'lvl': {},
            }
            self._num_caches_dict = caches
            return caches

    def _invalidate_num_caches(self):
        try:
            del self._num_caches_dict
        except AttributeError:
            pass

    def forget_list_items(self):
        """
        Forget the count of the list items. Call this after a paragraph is
        added to, removed from or moved in the document, or gets a different
        style: the count is for all the paragraphs of a story together.
        """
        self._num_caches['labels'].clear()

    def use_styles(self, styles_cache):
        """
        Remember *styles_cache*, a dict of the ``w:style`` elements by style
        id. A ``w:numStyleLink`` names a numbering style, and that style is
        in the styles, not in the numbering part.
        """
        if styles_cache:
            self._styles_cache = styles_cache

    def get_abstractNum(self, numId, _linked_from=()):
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

        # A definition with a numStyleLink has no levels of its own. They
        # are in the definition of its numbering style. If the link has no
        # target, the definition has only the levels that it holds.
        link = abstractNum.find(qn('w:numStyleLink')) if abstractNum is not None else None
        if link is not None:
            target, styles_known = self._abstractNum_of_numbering_style(
                link.get(qn('w:val')), _linked_from + (numId,))
            if target is not None:
                abstractNum = target
            elif not styles_known:
                # The styles can give the target later (see use_styles).
                return abstractNum

        cache[numId] = abstractNum
        return abstractNum

    def _abstractNum_of_numbering_style(self, style_id, linked_from):
        """
        Return ``(abstractNum, styles_known)`` for the numbering style having
        *style_id*. The definition is the one whose ``w:styleLink`` names
        the style, or else the one of the ``w:num`` that the style names in
        ``styles.xml`` (DCMR 21-522: no ``w:styleLink`` names the styles
        "RegsList2" and "RegsList3"). It is None if there is none.
        *styles_known* is False if the styles were necessary and are not
        known. *linked_from* has the numIds that led here, to stop a loop.
        """
        for el in self.abstractNum_lst:
            styleLink = el.find(qn('w:styleLink'))
            if styleLink is not None and styleLink.get(qn('w:val')) == style_id:
                return el, True
        styles_cache = getattr(self, '_styles_cache', None)
        if not styles_cache:
            return None, False
        try:
            numId = styles_cache[style_id].pPr.numPr.numId.val
        except (KeyError, AttributeError):
            return None, True
        if not numId or numId in linked_from:
            return None, True
        return self.get_abstractNum(numId, linked_from), True

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
        self.use_styles(styles_cache)
        try:
            numPr = p.pPr.numPr
            if styles_cache:
                # numId 0 turns off the numbering of the style, and with it
                # the indent of its level.
                if numPr is not None and numPr.numId is not None and numPr.numId.val == 0:
                    return None
                props = _ilvl_and_numId(_style_numPrs(p.pPr.style, styles_cache))
            else:
                props = _ilvl_and_numId([numPr] if numPr is not None else [])
        except AttributeError:
            return None
        if props is None:
            return None
        ilvl, numId = props
        return self._lvl_of(numId, ilvl)

    def get_lvl_from_style(self, style_id, styles_cache):
        """
        Return the ``w:lvl`` of the list that the style having *style_id*
        puts a paragraph in, or None if it puts it in no list. The numId and
        the ilvl each come from the style, or else from the nearest of the
        styles it is based on.
        """
        self.use_styles(styles_cache)
        props = _ilvl_and_numId(_style_numPrs(style_id, styles_cache or {}))
        if props is None:
            return None
        ilvl, numId = props
        return self._lvl_of(numId, ilvl)

    def get_num_for_p(self, p, styles_cache, append_suffix=True):
        """
        Returns list item for the given paragraph.
        The ``append_suffix`` flag is used to indicate the appending of
        numbered level suffix after the number
        (usually blank spaces between the number and the text of numbered paragraph).
        """
        self.use_styles(styles_cache)
        if self._get_ilvl_and_numId(p, styles_cache) is None:
            return None
        story = _story_of(p)
        labels_cache = self._num_caches['labels']
        items = labels_cache.get(story)
        if items is None or p not in items:
            items = labels_cache[story] = self._count_list_items(story, styles_cache)
        item = items.get(p)
        if item is None:
            return None
        numId, ilvl, counters = item
        lvl_el = self._lvl_of(numId, ilvl)

        numFmt = lvl_el.numFmt.get('{%s}val' % nsmap['w'])
        if numFmt not in self.fmt_map:
            return None
        # w:val is optional: a w:lvlText without it is an empty level text.
        lvlText = lvl_el.lvlText.get('{%s}val' % nsmap['w']) or ''

        isLgl = _is_on(lvl_el.find(qn('w:isLgl')))

        # Word shows each %N in the numFmt of its own level, or in decimal
        # if the current level has w:isLgl. A level with numFmt="none"
        # shows as nothing: "%12030.%2" shows "2030.1". A decimalZero level
        # stays decimalZero with w:isLgl, as in Aspose: "Section 1.01".
        def replace_token(m):
            target_ilvl = int(m.group(1)) - 1
            target_lvl_el = self._lvl_of(numId, target_ilvl)
            target_numFmt = numFmt
            if target_lvl_el is not None and target_lvl_el.numFmt is not None:
                target_numFmt = target_lvl_el.numFmt.get('{%s}val' % nsmap['w'])
            if target_numFmt == 'none':
                return ''
            if isLgl and target_numFmt != 'decimalZero':
                target_numFmt = 'decimal'
            value = counters.get(target_ilvl)
            if value is None:
                value = _start_value(target_lvl_el)
            formatter = self.fmt_map.get(target_numFmt, self.fmt_map[numFmt])
            return _format_value(formatter, value)

        suffix = lvl_el.suffix if append_suffix is True else ''
        return re.sub(r'%(\d)', replace_token, lvlText) + suffix

    def _count_list_items(self, story, styles_cache):
        """
        Count the list items of *story* in document order, as Word does.
        Returns a dict that maps each paragraph of *story* to
        ``(numId, ilvl, counters)``, or to None if it is not a list item.
        ``counters`` maps each started level of its list to the value that
        the label of the paragraph shows for it.

        Word keeps one counter per level for each abstract numbering
        definition; every ``w:num`` that uses the definition shares them.
        An item starts the levels above it that have not started, and
        restarts the levels below it (see ``w:lvlRestart``). A level starts
        at the ``w:startOverride`` of the item's ``w:num``, if it has one,
        or else at ``w:start``.

        A ``w:num`` restarts a counter one time only: at its first item at
        a level that has a ``w:startOverride``. That uses up its overrides
        at all the levels (DCMR 11-X305: after 305.1 restarts level 3, the
        first item of the same ``w:num`` at level 4 continues with "(d)").

        The label of an item shows a level above it that has no item of its
        own at the ``w:start`` of the definition, not at the value of its
        counter. With a ``w:startOverride`` of 5 on level 0 and of 3 on
        level 1, Word shows "1.3." for a first item at level 1, and "6."
        for the level 0 item after it. Aspose shows "5.3." and "6.".

        The counters belong to the definition that a ``w:num`` names, also
        when the levels come from a ``w:numStyleLink``, as in Aspose.

        A ``w:num`` can have a level that its definition does not have: its
        definition is not there, or the level is only in a ``w:lvlOverride``.
        Such a level starts at the ``w:start`` of that ``w:lvl``.
        """
        items = {}
        counters_of = {}
        with_item_of = {}
        list_ids = {}
        restarted = set()
        for p in story.iter(_w_p_tag):
            if _story_of(p) is not story:
                continue
            items[p] = None
            props = self._get_ilvl_and_numId(p, styles_cache)
            if props is None:
                continue
            ilvl, numId = props
            ilvl = self._level_in_list(numId, ilvl)
            if self._lvl_of(numId, ilvl) is None:
                continue
            abstractNum_el = self.get_abstractNum(numId)

            if numId not in list_ids:
                list_ids[numId] = self.num_having_numId(numId).abstractNumId.val
            counters = counters_of.setdefault(list_ids[numId], {})
            # the levels that have an item since their last restart
            with_item = with_item_of.setdefault(list_ids[numId], set())

            def definition_start_of(level):
                lvl_el = abstractNum_el.get_lvl(level) if abstractNum_el is not None else None
                if lvl_el is None:
                    lvl_el = self._override_lvl(numId, level)
                return _start_value(lvl_el)

            def start_of(level):
                startOverride = self._start_override(numId, level)
                if startOverride is None:
                    return definition_start_of(level)
                # With a w:lvl in the w:lvlOverride, Aspose restarts at its
                # w:start; LibreOffice at the startOverride.
                override_lvl_el = self._override_lvl(numId, level)
                if override_lvl_el is not None and override_lvl_el.start is not None:
                    return _start_value(override_lvl_el)
                return startOverride

            if (self._start_override(numId, ilvl) is not None
                    and numId not in restarted):
                restarted.add(numId)
                counters[ilvl] = start_of(ilvl)
            elif ilvl in counters:
                counters[ilvl] += 1
            else:
                counters[ilvl] = start_of(ilvl)
            with_item.add(ilvl)
            for upper in range(ilvl):
                if upper not in counters:
                    counters[upper] = start_of(upper)
            for lower in [i for i in counters if i > ilvl]:
                if ilvl < _restart_level(self._lvl_of(numId, lower), lower):
                    del counters[lower]
                    with_item.discard(lower)
            shown = dict(counters)
            for upper in range(ilvl):
                if upper not in with_item:
                    shown[upper] = definition_start_of(upper)
            items[p] = (numId, ilvl, shown)
        return items

    def _get_ilvl_and_numId(self, p, styles_cache):
        """
        Return ``(ilvl, numId)`` of the list *p* is in, or None if it is
        not in a list. Each comes from the paragraph, or else from the
        nearest of its styles that sets it.
        """
        pPr = p.pPr
        if pPr is None:
            return None
        numPrs = [pPr.numPr] if pPr.numPr is not None else []
        numPrs.extend(_style_numPrs(pPr.style, styles_cache or {}))
        return _ilvl_and_numId(numPrs)

    def _level_in_list(self, numId, ilvl):
        """
        Return the level that Word uses for an item of the ``<w:num>``
        having *numId* at *ilvl*. That is level 0 if the list is a
        ``singleLevel`` list and does not have *ilvl*: in DCMR 24-3306,
        the items at level 1 of such a list show "(1)" to "(5)". In a list
        of a different type, Word shows no label for an item at a level
        that the list does not have.
        """
        if ilvl and self._lvl_of(numId, ilvl) is None:
            abstractNum_el = self.get_abstractNum(numId)
            kind = (abstractNum_el.find(qn('w:multiLevelType'))
                    if abstractNum_el is not None else None)
            if kind is not None and kind.get(qn('w:val')) == 'singleLevel':
                return 0
        return ilvl

    def _lvl_of(self, numId, ilvl):
        """
        Return the ``w:lvl`` that the ``<w:num>`` having *numId* uses at
        *ilvl*: the one in its ``w:lvlOverride``, or else the one of its
        definition. Return None if there is none.
        """
        cache = self._num_caches['lvl']
        key = (numId, ilvl)
        if key not in cache:
            lvl_el = self._override_lvl(numId, ilvl)
            if lvl_el is None:
                abstractNum_el = self.get_abstractNum(numId)
                lvl_el = abstractNum_el.get_lvl(ilvl) if abstractNum_el is not None else None
                if numId not in self._num_caches['abstractNum']:
                    # The definition can change when the styles are known.
                    return lvl_el
            cache[key] = lvl_el
        return cache[key]

    def _override_lvl(self, numId, ilvl):
        """
        Return the ``w:lvl`` in the ``w:lvlOverride`` of the ``<w:num>``
        having *numId* for level *ilvl*, or None if there is none.
        """
        try:
            w_num = self.num_having_numId(numId)
        except KeyError:
            return None
        for lvlOverride in w_num.lvlOverride_lst:
            if lvlOverride.ilvl == ilvl:
                return lvlOverride.find(qn('w:lvl'))
        return None

    def _start_override(self, numId, ilvl):
        """
        Return the ``w:startOverride`` value of the ``<w:num>`` having
        *numId* for level *ilvl*, or None if there is none.
        """
        cache = self._num_caches['start_override']
        key = (numId, ilvl)
        if key not in cache:
            cache[key] = None
            for lvlOverride in self.num_having_numId(numId).lvlOverride_lst:
                if lvlOverride.ilvl == ilvl and lvlOverride.startOverride is not None:
                    cache[key] = int(lvlOverride.startOverride.get('{%s}val' % nsmap['w']))
                    break
        return cache[key]

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
