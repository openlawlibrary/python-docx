# encoding: utf-8

"""Section-related custom element classes"""

from __future__ import absolute_import, division, print_function, unicode_literals

from copy import deepcopy

from warnings import warn

from docx.enum.section import WD_HEADER_FOOTER, WD_ORIENTATION, WD_SECTION_START
from docx.oxml.simpletypes import ST_SignedTwipsMeasure, ST_TwipsMeasure, XsdString, ST_FtnPos, ST_NumberFormat, ST_RestartNumber
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


class CT_FtnPos(BaseOxmlElement):
    """``<w:pos>`` element, footnote placement"""
    val = RequiredAttribute('w:val', ST_FtnPos)


class CT_FtnProps(BaseOxmlElement):
    """``<w:footnotePr>`` element, section wide footnote properties"""
    _tag_seq = (
        'w:pos', 'w:numFmt', 'w:numStart', 'w:numRestart'
    )
    pos = ZeroOrOne('w:pos', successors=_tag_seq)
    numFmt = ZeroOrOne('w:numFmt', successors=_tag_seq[1:])
    numStart = ZeroOrOne('w:numStart', successors=_tag_seq[2:])
    numRestart = ZeroOrOne('w:numRestart', successors=_tag_seq[3:])


class CT_HdrFtr(BaseOxmlElement):
    """`w:hdr` and `w:ftr`, the root element for header and footer part respectively"""

    p = ZeroOrMore('w:p', successors=())
    tbl = ZeroOrMore('w:tbl', successors=())


class CT_HdrFtrRef(BaseOxmlElement):
    """`w:headerReference` and `w:footerReference` elements"""

    type_ = RequiredAttribute('w:type', WD_HEADER_FOOTER)
    rId = RequiredAttribute('r:id', XsdString)


class CT_NumFmt(BaseOxmlElement):
    """``<w:numFmt>`` element, footnote numbering format"""
    val = RequiredAttribute('w:val', ST_NumberFormat)


class CT_NumRestart(BaseOxmlElement):
    """``<w:numStart>`` element, footnote numbering restart location"""
    val = RequiredAttribute('w:val', ST_RestartNumber)


class CT_PageMar(BaseOxmlElement):
    """
    ``<w:pgMar>`` element, defining page margins.
    """
    top = OptionalAttribute('w:top', ST_SignedTwipsMeasure)
    right = OptionalAttribute('w:right', ST_TwipsMeasure)
    bottom = OptionalAttribute('w:bottom', ST_SignedTwipsMeasure)
    left = OptionalAttribute('w:left', ST_TwipsMeasure)
    header = OptionalAttribute('w:header', ST_TwipsMeasure)
    footer = OptionalAttribute('w:footer', ST_TwipsMeasure)
    gutter = OptionalAttribute('w:gutter', ST_TwipsMeasure)


class CT_PageSz(BaseOxmlElement):
    """
    ``<w:pgSz>`` element, defining page dimensions and orientation.
    """
    w = OptionalAttribute('w:w', ST_TwipsMeasure)
    h = OptionalAttribute('w:h', ST_TwipsMeasure)
    orient = OptionalAttribute(
        'w:orient', WD_ORIENTATION, default=WD_ORIENTATION.PORTRAIT
    )


class CT_SectPr(BaseOxmlElement):
    """`w:sectPr` element, the container element for section properties"""

    _tag_seq = (
        'w:footnotePr', 'w:endnotePr', 'w:type', 'w:pgSz', 'w:pgMar', 'w:paperSrc',
        'w:pgBorders', 'w:lnNumType', 'w:pgNumType', 'w:cols', 'w:formProt', 'w:vAlign',
        'w:noEndnote', 'w:titlePg', 'w:textDirection', 'w:bidi', 'w:rtlGutter',
        'w:docGrid', 'w:printerSettings', 'w:sectPrChange',
    )
    headerReference = ZeroOrMore("w:headerReference", successors=_tag_seq)
    footerReference = ZeroOrMore("w:footerReference", successors=_tag_seq)
    type = ZeroOrOne("w:type", successors=_tag_seq[3:])
    pgSz = ZeroOrOne("w:pgSz", successors=_tag_seq[4:])
    pgMar = ZeroOrOne("w:pgMar", successors=_tag_seq[5:])
    titlePg = ZeroOrOne("w:titlePg", successors=_tag_seq[14:])
    footnotePr = ZeroOrOne("w:footnotePr", successors=_tag_seq[1:])
    del _tag_seq

    def add_footerReference(self, type_, rId):
        """Return newly added CT_HdrFtrRef element of *type_* with *rId*.

        The element tag is `w:footerReference`.
        """
        footerReference = self._add_footerReference()
        footerReference.type_ = type_
        footerReference.rId = rId
        return footerReference

    def add_headerReference(self, type_, rId):
        """Return newly added CT_HdrFtrRef element of *type_* with *rId*.

        The element tag is `w:headerReference`.
        """
        headerReference = self._add_headerReference()
        headerReference.type_ = type_
        headerReference.rId = rId
        return headerReference

    @property
    def bottom_margin(self):
        """
        The value of the ``w:bottom`` attribute in the ``<w:pgMar>`` child
        element, as a |Length| object, or |None| if either the element or the
        attribute is not present.
        """
        pgMar = self.pgMar
        if pgMar is None:
            return None
        return pgMar.bottom

    @bottom_margin.setter
    def bottom_margin(self, value):
        pgMar = self.get_or_add_pgMar()
        pgMar.bottom = value

    def clone(self):
        """
        Return an exact duplicate of this ``<w:sectPr>`` element tree
        suitable for use in adding a section break. All rsid* attributes are
        removed from the root ``<w:sectPr>`` element.
        """
        clone_sectPr = deepcopy(self)
        clone_sectPr.attrib.clear()
        return clone_sectPr

    @property
    def footer(self):
        """
        The value of the ``w:footer`` attribute in the ``<w:pgMar>`` child
        element, as a |Length| object, or |None| if either the element or the
        attribute is not present.
        """
        pgMar = self.pgMar
        if pgMar is None:
            return None
        return pgMar.footer

    @footer.setter
    def footer(self, value):
        pgMar = self.get_or_add_pgMar()
        pgMar.footer = value

    @property
    def footnote_number_format(self):
        """
        The value of the ``w:val`` attribute in the ``<w:numFmt>`` child
        element of ``<w:footnotePr>`` element, as a |String|, or |None| if either the element or the
        attribute is not present.
        """
        fPr = self.footnotePr
        if fPr is None or fPr.numFmt is None:
            return None
        return fPr.numFmt.val

    @footnote_number_format.setter
    def footnote_number_format(self, value):
        fPr = self.get_or_add_footnotePr()
        numFmt = fPr.get_or_add_numFmt()
        numFmt.val = value

    @property
    def footnote_numbering_restart_location(self):
        """
        The value of the ``w:val`` attribute in the ``<w:numRestart>`` child
        element of ``<w:footnotePr>`` element, as a |String|, or |None| if either the element or the
        attribute is not present.
        """
        fPr = self.footnotePr
        if fPr is None or fPr.numRestart is None:
            return None
        return fPr.numRestart.val

    @footnote_numbering_restart_location.setter
    def footnote_numbering_restart_location(self, value):
        fPr = self.get_or_add_footnotePr()
        numStart = fPr.get_or_add_numStart()
        numRestart = fPr.get_or_add_numRestart()
        numRestart.val = value
        if numStart is None or len(numStart.values()) == 0:
            numStart.val = 1
        elif value != 'continuous':
            numStart.val = 1
            msg = "When ``<w:numRestart> is not 'continuous', then ``<w:numStart>`` must be 1."
            warn(msg, UserWarning, stacklevel=2)

    @property
    def footnote_numbering_start_value(self):
        """
        The value of the ``w:val`` attribute in the ``<w:numStart>`` child
        element of ``<w:footnotePr>`` element, as a |Number|, or |None| if either the element or the
        attribute is not present.
        """
        fPr = self.footnotePr
        if fPr is None or fPr.numStart is None:
            return None
        return fPr.numStart.val

    @footnote_numbering_start_value.setter
    def footnote_numbering_start_value(self, value):
        fPr = self.get_or_add_footnotePr()
        numStart = fPr.get_or_add_numStart()
        numRestart = fPr.get_or_add_numRestart()
        numStart.val = value
        if numRestart is None or len(numRestart.values()) == 0:
            numRestart.val = 'continuous'
        elif value != 1:
            numRestart.val = 'continuous'
            msg = "When ``<w:numStart> is not 1, then ``<w:numRestart>`` must be 'continuous'."
            warn(msg, UserWarning, stacklevel=2)

    @property
    def footnote_position(self):
        """
        The value of the ``w:val`` attribute in the ``<w:pos>`` child
        element of ``<w:footnotePr>`` element, as a |String|, or |None| if either the element or the
        attribute is not present.
        """
        fPr = self.footnotePr
        if fPr is None or fPr.pos is None:
            return None
        return fPr.pos.val

    @footnote_position.setter
    def footnote_position(self, value):
        fPr = self.get_or_add_footnotePr()
        pos = fPr.get_or_add_pos()
        pos.val = value

    def get_footerReference(self, type_):
        """Return footerReference element of *type_* or None if not present."""
        path = "./w:footerReference[@w:type='%s']" % WD_HEADER_FOOTER.to_xml(type_)
        footerReferences = self.xpath(path)
        if not footerReferences:
            return None
        return footerReferences[0]

    def get_headerReference(self, type_):
        """Return headerReference element of *type_* or None if not present."""
        matching_headerReferences = self.xpath(
            "./w:headerReference[@w:type='%s']" % WD_HEADER_FOOTER.to_xml(type_)
        )
        if len(matching_headerReferences) == 0:
            return None
        return matching_headerReferences[0]

    @property
    def gutter(self):
        """
        The value of the ``w:gutter`` attribute in the ``<w:pgMar>`` child
        element, as a |Length| object, or |None| if either the element or the
        attribute is not present.
        """
        pgMar = self.pgMar
        if pgMar is None:
            return None
        return pgMar.gutter

    @gutter.setter
    def gutter(self, value):
        pgMar = self.get_or_add_pgMar()
        pgMar.gutter = value

    @property
    def header(self):
        """
        The value of the ``w:header`` attribute in the ``<w:pgMar>`` child
        element, as a |Length| object, or |None| if either the element or the
        attribute is not present.
        """
        pgMar = self.pgMar
        if pgMar is None:
            return None
        return pgMar.header

    @header.setter
    def header(self, value):
        pgMar = self.get_or_add_pgMar()
        pgMar.header = value

    @property
    def left_margin(self):
        """
        The value of the ``w:left`` attribute in the ``<w:pgMar>`` child
        element, as a |Length| object, or |None| if either the element or the
        attribute is not present.
        """
        pgMar = self.pgMar
        if pgMar is None:
            return None
        return pgMar.left

    @left_margin.setter
    def left_margin(self, value):
        pgMar = self.get_or_add_pgMar()
        pgMar.left = value

    @property
    def orientation(self):
        """
        The member of the ``WD_ORIENTATION`` enumeration corresponding to the
        value of the ``orient`` attribute of the ``<w:pgSz>`` child element,
        or ``WD_ORIENTATION.PORTRAIT`` if not present.
        """
        pgSz = self.pgSz
        if pgSz is None:
            return WD_ORIENTATION.PORTRAIT
        return pgSz.orient

    @orientation.setter
    def orientation(self, value):
        pgSz = self.get_or_add_pgSz()
        pgSz.orient = value

    @property
    def page_height(self):
        """
        Value in EMU of the ``h`` attribute of the ``<w:pgSz>`` child
        element, or |None| if not present.
        """
        pgSz = self.pgSz
        if pgSz is None:
            return None
        return pgSz.h

    @page_height.setter
    def page_height(self, value):
        pgSz = self.get_or_add_pgSz()
        pgSz.h = value

    @property
    def page_width(self):
        """
        Value in EMU of the ``w`` attribute of the ``<w:pgSz>`` child
        element, or |None| if not present.
        """
        pgSz = self.pgSz
        if pgSz is None:
            return None
        return pgSz.w

    @page_width.setter
    def page_width(self, value):
        pgSz = self.get_or_add_pgSz()
        pgSz.w = value

    @property
    def preceding_sectPr(self):
        """sectPr immediately preceding this one or None if this is the first."""
        # ---[1] predicate returns list of zero or one value---
        preceding_sectPrs = self.xpath("./preceding::w:sectPr[1]")
        return preceding_sectPrs[0] if len(preceding_sectPrs) > 0 else None

    def remove_footerReference(self, type_):
        """Return rId of w:footerReference child of *type_* after removing it."""
        footerReference = self.get_footerReference(type_)
        rId = footerReference.rId
        self.remove(footerReference)
        return rId

    def remove_headerReference(self, type_):
        """Return rId of w:headerReference child of *type_* after removing it."""
        headerReference = self.get_headerReference(type_)
        rId = headerReference.rId
        self.remove(headerReference)
        return rId

    @property
    def right_margin(self):
        """
        The value of the ``w:right`` attribute in the ``<w:pgMar>`` child
        element, as a |Length| object, or |None| if either the element or the
        attribute is not present.
        """
        pgMar = self.pgMar
        if pgMar is None:
            return None
        return pgMar.right

    @right_margin.setter
    def right_margin(self, value):
        pgMar = self.get_or_add_pgMar()
        pgMar.right = value

    @property
    def start_type(self):
        """
        The member of the ``WD_SECTION_START`` enumeration corresponding to
        the value of the ``val`` attribute of the ``<w:type>`` child element,
        or ``WD_SECTION_START.NEW_PAGE`` if not present.
        """
        type = self.type
        if type is None or type.val is None:
            return WD_SECTION_START.NEW_PAGE
        return type.val

    @start_type.setter
    def start_type(self, value):
        if value is None or value is WD_SECTION_START.NEW_PAGE:
            self._remove_type()
            return
        type = self.get_or_add_type()
        type.val = value

    @property
    def titlePg_val(self):
        """Value of `w:titlePg/@val` or |None| if not present"""
        titlePg = self.titlePg
        if titlePg is None:
            return False
        return titlePg.val

    @titlePg_val.setter
    def titlePg_val(self, value):
        if value in [None, False]:
            self._remove_titlePg()
        else:
            self.get_or_add_titlePg().val = value

    @property
    def top_margin(self):
        """
        The value of the ``w:top`` attribute in the ``<w:pgMar>`` child
        element, as a |Length| object, or |None| if either the element or the
        attribute is not present.
        """
        pgMar = self.pgMar
        if pgMar is None:
            return None
        return pgMar.top

    @top_margin.setter
    def top_margin(self, value):
        pgMar = self.get_or_add_pgMar()
        pgMar.top = value


class CT_SectType(BaseOxmlElement):
    """
    ``<w:sectType>`` element, defining the section start type.
    """
    val = OptionalAttribute('w:val', WD_SECTION_START)
