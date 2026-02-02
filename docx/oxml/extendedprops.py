# encoding: utf-8

"""Custom element classes for extended properties-related XML elements."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from docx.oxml import parse_xml
from docx.oxml.ns import nsmap
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrOne


class CT_ExtendedProperties(BaseOxmlElement):
    """
    ``<Properties>`` element, the root element of the Extended Properties
    part stored as ``/docProps/app.xml``. Implements the extended document
    properties such as TotalTime, Pages, Words, etc.
    """
    Template = ZeroOrOne('ep:Template', successors=())
    Manager = ZeroOrOne('ep:Manager', successors=())
    Company = ZeroOrOne('ep:Company', successors=())
    Pages = ZeroOrOne('ep:Pages', successors=())
    Words = ZeroOrOne('ep:Words', successors=())
    Characters = ZeroOrOne('ep:Characters', successors=())
    PresentationFormat = ZeroOrOne('ep:PresentationFormat', successors=())
    Lines = ZeroOrOne('ep:Lines', successors=())
    Paragraphs = ZeroOrOne('ep:Paragraphs', successors=())
    Slides = ZeroOrOne('ep:Slides', successors=())
    Notes = ZeroOrOne('ep:Notes', successors=())
    TotalTime = ZeroOrOne('ep:TotalTime', successors=())
    HiddenSlides = ZeroOrOne('ep:HiddenSlides', successors=())
    MMClips = ZeroOrOne('ep:MMClips', successors=())
    ScaleCrop = ZeroOrOne('ep:ScaleCrop', successors=())
    LinksUpToDate = ZeroOrOne('ep:LinksUpToDate', successors=())
    CharactersWithSpaces = ZeroOrOne('ep:CharactersWithSpaces', successors=())
    SharedDoc = ZeroOrOne('ep:SharedDoc', successors=())
    HyperlinkBase = ZeroOrOne('ep:HyperlinkBase', successors=())
    HyperlinksChanged = ZeroOrOne('ep:HyperlinksChanged', successors=())
    Application = ZeroOrOne('ep:Application', successors=())
    AppVersion = ZeroOrOne('ep:AppVersion', successors=())
    DocSecurity = ZeroOrOne('ep:DocSecurity', successors=())

    _extendedProperties_tmpl = (
        '<Properties xmlns="%s" xmlns:vt="%s"/>\n'
        % (nsmap['ep'], nsmap['vt'])
    )

    @classmethod
    def new(cls):
        """
        Return a new ``<Properties>`` element
        """
        xml = cls._extendedProperties_tmpl
        extendedProperties = parse_xml(xml)
        return extendedProperties

    # --- Integer properties ---

    @property
    def totalTime_int(self):
        """
        Integer value of TotalTime property (editing time in minutes).
        """
        totalTime = self.TotalTime
        if totalTime is None:
            return 0
        totalTime_str = totalTime.text
        if totalTime_str is None:
            return 0
        try:
            value = int(totalTime_str)
        except ValueError:
            value = 0
        if value < 0:
            value = 0
        return value

    @totalTime_int.setter
    def totalTime_int(self, value):
        """
        Set TotalTime property to string value of integer *value*.
        """
        if not isinstance(value, int) or value < 0:
            tmpl = "TotalTime property requires non-negative int, got '%s'"
            raise ValueError(tmpl % value)
        totalTime = self.get_or_add_TotalTime()
        totalTime.text = str(value)

    @property
    def pages_int(self):
        """Integer value of Pages property."""
        return self._int_of_element('Pages')

    @pages_int.setter
    def pages_int(self, value):
        self._set_element_int('Pages', value)

    @property
    def words_int(self):
        """Integer value of Words property."""
        return self._int_of_element('Words')

    @words_int.setter
    def words_int(self, value):
        self._set_element_int('Words', value)

    @property
    def characters_int(self):
        """Integer value of Characters property."""
        return self._int_of_element('Characters')

    @characters_int.setter
    def characters_int(self, value):
        self._set_element_int('Characters', value)

    @property
    def lines_int(self):
        """Integer value of Lines property."""
        return self._int_of_element('Lines')

    @lines_int.setter
    def lines_int(self, value):
        self._set_element_int('Lines', value)

    @property
    def paragraphs_int(self):
        """Integer value of Paragraphs property."""
        return self._int_of_element('Paragraphs')

    @paragraphs_int.setter
    def paragraphs_int(self, value):
        self._set_element_int('Paragraphs', value)

    @property
    def charactersWithSpaces_int(self):
        """Integer value of CharactersWithSpaces property."""
        return self._int_of_element('CharactersWithSpaces')

    @charactersWithSpaces_int.setter
    def charactersWithSpaces_int(self, value):
        self._set_element_int('CharactersWithSpaces', value)

    @property
    def docSecurity_int(self):
        """Integer value of DocSecurity property."""
        return self._int_of_element('DocSecurity')

    @docSecurity_int.setter
    def docSecurity_int(self, value):
        self._set_element_int('DocSecurity', value)

    # --- String properties ---

    @property
    def template_text(self):
        """String value of Template property."""
        return self._text_of_element('Template')

    @template_text.setter
    def template_text(self, value):
        self._set_element_text('Template', value)

    @property
    def manager_text(self):
        """String value of Manager property."""
        return self._text_of_element('Manager')

    @manager_text.setter
    def manager_text(self, value):
        self._set_element_text('Manager', value)

    @property
    def company_text(self):
        """String value of Company property."""
        return self._text_of_element('Company')

    @company_text.setter
    def company_text(self, value):
        self._set_element_text('Company', value)

    @property
    def application_text(self):
        """String value of Application property."""
        return self._text_of_element('Application')

    @application_text.setter
    def application_text(self, value):
        self._set_element_text('Application', value)

    @property
    def appVersion_text(self):
        """String value of AppVersion property."""
        return self._text_of_element('AppVersion')

    @appVersion_text.setter
    def appVersion_text(self, value):
        self._set_element_text('AppVersion', value)

    @property
    def hyperlinkBase_text(self):
        """String value of HyperlinkBase property."""
        return self._text_of_element('HyperlinkBase')

    @hyperlinkBase_text.setter
    def hyperlinkBase_text(self, value):
        self._set_element_text('HyperlinkBase', value)

    # --- Helper methods ---

    def _int_of_element(self, property_name):
        """
        Return the integer value of the element matching *property_name*,
        or 0 if the element is not present or contains invalid text.
        """
        element = getattr(self, property_name)
        if element is None:
            return 0
        text = element.text
        if text is None:
            return 0
        try:
            value = int(text)
        except ValueError:
            value = 0
        if value < 0:
            value = 0
        return value

    def _set_element_int(self, prop_name, value):
        """Set integer value of *prop_name* property to *value*."""
        if not isinstance(value, int) or value < 0:
            tmpl = "%s property requires non-negative int, got '%s'"
            raise ValueError(tmpl % (prop_name, value))
        get_or_add_method = getattr(self, 'get_or_add_%s' % prop_name)
        element = get_or_add_method()
        element.text = str(value)

    def _text_of_element(self, property_name):
        """
        Return the text in the element matching *property_name*, or an empty
        string if the element is not present or contains no text.
        """
        element = getattr(self, property_name)
        if element is None:
            return ''
        if element.text is None:
            return ''
        return element.text

    def _set_element_text(self, prop_name, value):
        """Set string value of *prop_name* property to *value*."""
        if value is None:
            value = ''
        if not isinstance(value, str):
            value = str(value)
        get_or_add_method = getattr(self, 'get_or_add_%s' % prop_name)
        element = get_or_add_method()
        element.text = value
