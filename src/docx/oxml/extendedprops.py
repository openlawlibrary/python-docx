"""Custom element class for the extended-properties part root element."""

from __future__ import annotations

from typing import Any, cast

from docx.oxml.ns import nsmap
from docx.oxml.parser import parse_xml
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrOne


class CT_ExtendedProperties(BaseOxmlElement):
    """`<Properties>` element, the root element of the Extended Properties part.

    Stored as `/docProps/app.xml`. Implements the extended document properties such as
    TotalTime, Pages, and Words.
    """

    Template = ZeroOrOne("ep:Template", successors=())
    Manager = ZeroOrOne("ep:Manager", successors=())
    Company = ZeroOrOne("ep:Company", successors=())
    Pages = ZeroOrOne("ep:Pages", successors=())
    Words = ZeroOrOne("ep:Words", successors=())
    Characters = ZeroOrOne("ep:Characters", successors=())
    Lines = ZeroOrOne("ep:Lines", successors=())
    Paragraphs = ZeroOrOne("ep:Paragraphs", successors=())
    TotalTime = ZeroOrOne("ep:TotalTime", successors=())
    CharactersWithSpaces = ZeroOrOne("ep:CharactersWithSpaces", successors=())
    HyperlinkBase = ZeroOrOne("ep:HyperlinkBase", successors=())
    Application = ZeroOrOne("ep:Application", successors=())
    AppVersion = ZeroOrOne("ep:AppVersion", successors=())
    DocSecurity = ZeroOrOne("ep:DocSecurity", successors=())

    _extendedProperties_tmpl = '<Properties xmlns="%s" xmlns:vt="%s"/>\n' % (
        nsmap["ep"],
        nsmap["vt"],
    )

    @classmethod
    def new(cls) -> CT_ExtendedProperties:
        """Return a new `<Properties>` element."""
        return cast(CT_ExtendedProperties, parse_xml(cls._extendedProperties_tmpl))

    # -- integer properties --------------------------------------------------------

    @property
    def totalTime_int(self) -> int:
        """Integer value of TotalTime property (editing time in minutes)."""
        return self._int_of_element("TotalTime")

    @totalTime_int.setter
    def totalTime_int(self, value: int) -> None:
        self._set_element_int("TotalTime", value)

    @property
    def pages_int(self) -> int:
        """Integer value of Pages property."""
        return self._int_of_element("Pages")

    @pages_int.setter
    def pages_int(self, value: int) -> None:
        self._set_element_int("Pages", value)

    @property
    def words_int(self) -> int:
        """Integer value of Words property."""
        return self._int_of_element("Words")

    @words_int.setter
    def words_int(self, value: int) -> None:
        self._set_element_int("Words", value)

    @property
    def characters_int(self) -> int:
        """Integer value of Characters property."""
        return self._int_of_element("Characters")

    @characters_int.setter
    def characters_int(self, value: int) -> None:
        self._set_element_int("Characters", value)

    @property
    def lines_int(self) -> int:
        """Integer value of Lines property."""
        return self._int_of_element("Lines")

    @lines_int.setter
    def lines_int(self, value: int) -> None:
        self._set_element_int("Lines", value)

    @property
    def paragraphs_int(self) -> int:
        """Integer value of Paragraphs property."""
        return self._int_of_element("Paragraphs")

    @paragraphs_int.setter
    def paragraphs_int(self, value: int) -> None:
        self._set_element_int("Paragraphs", value)

    @property
    def charactersWithSpaces_int(self) -> int:
        """Integer value of CharactersWithSpaces property."""
        return self._int_of_element("CharactersWithSpaces")

    @charactersWithSpaces_int.setter
    def charactersWithSpaces_int(self, value: int) -> None:
        self._set_element_int("CharactersWithSpaces", value)

    @property
    def docSecurity_int(self) -> int:
        """Integer value of DocSecurity property."""
        return self._int_of_element("DocSecurity")

    @docSecurity_int.setter
    def docSecurity_int(self, value: int) -> None:
        self._set_element_int("DocSecurity", value)

    # -- string properties ----------------------------------------------------------

    @property
    def template_text(self) -> str:
        """String value of Template property."""
        return self._text_of_element("Template")

    @template_text.setter
    def template_text(self, value: str) -> None:
        self._set_element_text("Template", value)

    @property
    def manager_text(self) -> str:
        """String value of Manager property."""
        return self._text_of_element("Manager")

    @manager_text.setter
    def manager_text(self, value: str) -> None:
        self._set_element_text("Manager", value)

    @property
    def company_text(self) -> str:
        """String value of Company property."""
        return self._text_of_element("Company")

    @company_text.setter
    def company_text(self, value: str) -> None:
        self._set_element_text("Company", value)

    @property
    def application_text(self) -> str:
        """String value of Application property."""
        return self._text_of_element("Application")

    @application_text.setter
    def application_text(self, value: str) -> None:
        self._set_element_text("Application", value)

    @property
    def appVersion_text(self) -> str:
        """String value of AppVersion property."""
        return self._text_of_element("AppVersion")

    @appVersion_text.setter
    def appVersion_text(self, value: str) -> None:
        self._set_element_text("AppVersion", value)

    @property
    def hyperlinkBase_text(self) -> str:
        """String value of HyperlinkBase property."""
        return self._text_of_element("HyperlinkBase")

    @hyperlinkBase_text.setter
    def hyperlinkBase_text(self, value: str) -> None:
        self._set_element_text("HyperlinkBase", value)

    # -- helpers ----------------------------------------------------------------

    def _get_or_add(self, prop_name: str) -> BaseOxmlElement:
        """Element returned by the "get_or_add_" method for `prop_name`."""
        get_or_add_method = getattr(self, "get_or_add_%s" % prop_name)
        return get_or_add_method()

    def _int_of_element(self, property_name: str) -> int:
        """Integer value of the element matching `property_name`.

        0 if the element is not present, contains invalid text, or is negative.
        """
        element = getattr(self, property_name)
        if element is None or element.text is None:
            return 0
        try:
            value = int(element.text)
        except ValueError:
            return 0
        return 0 if value < 0 else value

    def _set_element_int(self, prop_name: str, value: Any) -> None:
        """Set integer value of `prop_name` property to `value`."""
        if not isinstance(value, int) or value < 0:
            raise ValueError("%s property requires non-negative int, got '%s'" % (prop_name, value))
        element = self._get_or_add(prop_name)
        element.text = str(value)

    def _text_of_element(self, property_name: str) -> str:
        """Text in the element matching `property_name`.

        The empty string if the element is not present or contains no text.
        """
        element = getattr(self, property_name)
        if element is None or element.text is None:
            return ""
        return element.text

    def _set_element_text(self, prop_name: str, value: Any) -> None:
        """Set string value of `prop_name` property to `value`."""
        if not isinstance(value, str):
            value = str(value)
        element = self._get_or_add(prop_name)
        element.text = value
