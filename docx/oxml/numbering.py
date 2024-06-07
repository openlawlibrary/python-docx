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
from .ns import nsmap


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
        return self._insert_num(num)

    def get_abstractNum(self, numId):
        """
        Returns |CT_AbstractNum| instance with corresponding
        paragraph ``pPr.numPr.numId`` if any
        """
        try:
            num_el = self.num_having_numId(numId)
        except KeyError:
            return None

        abstractNum_id = num_el.abstractNumId.val

        for el in self.abstractNum_lst:
            if el.abstractNumId == abstractNum_id:
                return el

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
        """

        def get_ilvl_and_numId(paragraph):
            """
            Return ``ilvl`` and ``numId`` for the given ``paragraph``.
            """
            para_numPr = paragraph.pPr.get_numPr(styles_cache)
            para_ilvl, para_numId = para_numPr.ilvl, para_numPr.numId.val
            para_ilvl = para_ilvl.val if para_ilvl is not None else 0
            return para_ilvl, para_numId

        def get_preceding_paragraphs_numIds(p, p_ilvl, p_numId):
            """
            Return preceding siblings ``numId`` that are in the same numbered list as the paragraph ``p``.
            Paragraphs are in the same list if they are on the same level (``p_ilvl``), and either
            have the same ``p_numId`` or have the same ``pStyle.val``.
            Skips unnumbered paragraphs within the numbering list.
            Stops on the paragraph that is on the lower level.
            """
            pStyle = p.pPr.pStyle
            for prev_p in p.itersiblings(preceding=True):
                try:
                    prev_p_ilvl, prev_p_numId = get_ilvl_and_numId(prev_p)
                    # skip unnumbered paragraphs within numbering list
                    if prev_p_numId == 0:
                        continue
                    prev_p_pStyle = prev_p.pPr.pStyle
                    if prev_p_ilvl < p_ilvl and (prev_p_numId == p_numId or
                                         (prev_p_pStyle is not None and prev_p_pStyle.val in linked_styles)):
                        break
                    if prev_p_ilvl == p_ilvl and (prev_p_numId == p_numId or pStyle.val == prev_p_pStyle.val or prev_p_pStyle.val in linked_styles):
                        yield prev_p_numId
                    # para `p` that has only style defined which is same as the `prev_p` style
                    # should be counted even though they have different `numId`s.
                    if prev_p_ilvl == p_ilvl and prev_p_numId != p_numId:
                        if p.pPr.numPr is None and prev_p.pPr.pStyle.val == pStyle.val:
                            startOverride = get_start_override(prev_p_numId)
                            if startOverride > 1:
                                yield prev_p_numId
                            else:
                                yield p_numId
                            break
                except AttributeError:
                    continue

        def get_preceding_paragraph(p, p_ilvl):
            """
            Returns the first sibling that has the same numbering format
            """
            pStyle = p.pPr.pStyle
            for prev_p in p.itersiblings(preceding=True):
                try:
                    prev_p_ilvl, prev_p_numId = get_ilvl_and_numId(prev_p)
                    # skip unnumbered paragraphs within numbering list
                    if prev_p_numId == 0:
                        continue
                    prev_p_pStyle = prev_p.pPr.pStyle
                    if prev_p_ilvl <= p_ilvl and pStyle.val == prev_p_pStyle.val:
                        return (prev_p, prev_p_ilvl, prev_p_numId)
                except AttributeError:
                    continue
            return None

        def count_same_numIds(preceding_paragraphs_numIds, numId, num):
            """
            Returns count of the preceding paragraphs having the same ``w:numId``.
            """
            for p_numId in preceding_paragraphs_numIds:
                try:
                    if numId == p_numId:
                        num += 1
                    else:
                        startOverride = get_start_override(p_numId)
                        if startOverride > 1:
                            num += startOverride
                        else:
                            prev_numId = next(preceding_paragraphs_numIds)
                            if prev_numId == numId:
                                num += 1
                        break
                except AttributeError:
                    continue
                except StopIteration:
                    break
            return num

        def get_start_override(for_numId):
            w_num = self.num_having_numId(for_numId)
            for lvlOverride in w_num.lvlOverride_lst:
                if lvlOverride.ilvl == ilvl:
                    return lvlOverride.startOverride.val
            return 0

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
            p_num = self.fmt_map[lvl_el.numFmt.get('{%s}val' % nsmap['w'])](p_num)
        except KeyError:
            return None

        suffix = ''
        if append_suffix is True:
            suffix = lvl_el.suffix
        lvlText = lvl_el.lvlText.get('{%s}val' % nsmap['w'])
        if lvlText.count('%') > 1:
            # `lvlText` is an custom defined list label that has multiple numbering values
            # e.g. `1.1.2`, `1.1.3`
            prev_p_tuple = get_preceding_paragraph(p, ilvl)
            if prev_p_tuple is not None:
                prev_p, prev_p_ilvl, _ = prev_p_tuple
                prev_num = self.get_num_for_p(prev_p, styles_cache, append_suffix=False)
                # get text that is before and after every number part of the list text
                lvl_text_split_by_num = str(re.sub(r'%(\d)', '$', lvlText)).split("$")
                pre_num_text = lvl_text_split_by_num[0]
                after_num_text = lvl_text_split_by_num[-1]
                if prev_p_ilvl < ilvl:
                    # first indented paragraph => on prev para num append new indent number
                    return f"{prev_num}{pre_num_text}{p_num}{after_num_text}{suffix}"
                # copy the prev list numbers and bump the last number to the correct value
                if len(after_num_text) > 0:
                    prev_num_after_text = prev_num.split(after_num_text)
                    prev_num_after_text[-2] = pre_num_text + str(p_num)
                    return after_num_text.join(prev_num_after_text) + suffix
                # the numbering style is something like `#1#1#2`
                if len(pre_num_text) > 0:
                    prev_num_pre_text = prev_num.split(pre_num_text)
                    prev_num_pre_text[-1] = str(p_num) + after_num_text
                    return pre_num_text.join(prev_num_pre_text) + suffix
                # the number style is like `1.2`, so no char before or after, only in the middle
                mid_num_text = lvl_text_split_by_num[1]
                prev_num_mid_text = prev_num.split(mid_num_text)
                prev_num_mid_text[-1] = str(p_num)
                if p_num == 1:
                    # increment the first number
                    # this is specific case for bylaw
                    prev_num_mid_text[0] = str(int(prev_num_mid_text[0])+1)
                return mid_num_text.join(prev_num_mid_text) + suffix
            # set all number parts to default value
            return re.sub(r'%(\d)', str(p_num), lvlText) + suffix
        return re.sub(r'%(\d)', str(p_num), lvlText, 1) + suffix

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
