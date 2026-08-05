"""Provides access to extended document properties (`/docProps/app.xml`)."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docx.oxml.extendedprops import CT_ExtendedProperties


class ExtendedProperties:
    """Corresponds to part named `/docProps/app.xml`.

    Contains the extended document properties for this document package.
    """

    def __init__(self, element: CT_ExtendedProperties):
        self._element = element

    @property
    def total_time(self) -> int:
        """Total editing time in minutes."""
        return self._element.totalTime_int

    @total_time.setter
    def total_time(self, value: int) -> None:
        self._element.totalTime_int = value

    @property
    def pages(self) -> int:
        """Number of pages in the document."""
        return self._element.pages_int

    @pages.setter
    def pages(self, value: int) -> None:
        self._element.pages_int = value

    @property
    def words(self) -> int:
        """Number of words in the document."""
        return self._element.words_int

    @words.setter
    def words(self, value: int) -> None:
        self._element.words_int = value

    @property
    def characters(self) -> int:
        """Number of characters in the document."""
        return self._element.characters_int

    @characters.setter
    def characters(self, value: int) -> None:
        self._element.characters_int = value

    @property
    def lines(self) -> int:
        """Number of lines in the document."""
        return self._element.lines_int

    @lines.setter
    def lines(self, value: int) -> None:
        self._element.lines_int = value

    @property
    def paragraphs(self) -> int:
        """Number of paragraphs in the document."""
        return self._element.paragraphs_int

    @paragraphs.setter
    def paragraphs(self, value: int) -> None:
        self._element.paragraphs_int = value

    @property
    def characters_with_spaces(self) -> int:
        """Number of characters including spaces."""
        return self._element.charactersWithSpaces_int

    @characters_with_spaces.setter
    def characters_with_spaces(self, value: int) -> None:
        self._element.charactersWithSpaces_int = value

    @property
    def doc_security(self) -> int:
        """Document security level."""
        return self._element.docSecurity_int

    @doc_security.setter
    def doc_security(self, value: int) -> None:
        self._element.docSecurity_int = value

    @property
    def template(self) -> str:
        """Name of the template used."""
        return self._element.template_text

    @template.setter
    def template(self, value: str) -> None:
        self._element.template_text = value

    @property
    def manager(self) -> str:
        """Manager of the document author."""
        return self._element.manager_text

    @manager.setter
    def manager(self, value: str) -> None:
        self._element.manager_text = value

    @property
    def company(self) -> str:
        """Company name."""
        return self._element.company_text

    @company.setter
    def company(self, value: str) -> None:
        self._element.company_text = value

    @property
    def application(self) -> str:
        """Application that created the document."""
        return self._element.application_text

    @application.setter
    def application(self, value: str) -> None:
        self._element.application_text = value

    @property
    def app_version(self) -> str:
        """Version of the application."""
        return self._element.appVersion_text

    @app_version.setter
    def app_version(self, value: str) -> None:
        self._element.appVersion_text = value

    @property
    def hyperlink_base(self) -> str:
        """Base URL for relative hyperlinks."""
        return self._element.hyperlinkBase_text

    @hyperlink_base.setter
    def hyperlink_base(self, value: str) -> None:
        self._element.hyperlinkBase_text = value
