# encoding: utf-8

"""
The :mod:`docx.opc.extendedprops` module provides access to extended
document properties (app.xml).
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)


class ExtendedProperties(object):
    """
    Corresponds to part named ``/docProps/app.xml``, containing the extended
    document properties for this document package.
    """
    def __init__(self, element):
        self._element = element

    @property
    def total_time(self):
        """
        Total editing time in minutes. Read/write. Returns an integer.
        """
        return self._element.totalTime_int

    @total_time.setter
    def total_time(self, value):
        self._element.totalTime_int = value

    @property
    def pages(self):
        """
        Number of pages in the document. Read/write. Returns an integer.
        """
        return self._element.pages_int

    @pages.setter
    def pages(self, value):
        self._element.pages_int = value

    @property
    def words(self):
        """
        Number of words in the document. Read/write. Returns an integer.
        """
        return self._element.words_int

    @words.setter
    def words(self, value):
        self._element.words_int = value

    @property
    def characters(self):
        """
        Number of characters in the document. Read/write. Returns an integer.
        """
        return self._element.characters_int

    @characters.setter
    def characters(self, value):
        self._element.characters_int = value

    @property
    def lines(self):
        """
        Number of lines in the document. Read/write. Returns an integer.
        """
        return self._element.lines_int

    @lines.setter
    def lines(self, value):
        self._element.lines_int = value

    @property
    def paragraphs(self):
        """
        Number of paragraphs in the document. Read/write. Returns an integer.
        """
        return self._element.paragraphs_int

    @paragraphs.setter
    def paragraphs(self, value):
        self._element.paragraphs_int = value

    @property
    def characters_with_spaces(self):
        """
        Number of characters including spaces. Read/write. Returns an integer.
        """
        return self._element.charactersWithSpaces_int

    @characters_with_spaces.setter
    def characters_with_spaces(self, value):
        self._element.charactersWithSpaces_int = value

    @property
    def doc_security(self):
        """
        Document security level. Read/write. Returns an integer.
        """
        return self._element.docSecurity_int

    @doc_security.setter
    def doc_security(self, value):
        self._element.docSecurity_int = value

    @property
    def template(self):
        """
        Name of the template used. Read/write. Returns a string.
        """
        return self._element.template_text

    @template.setter
    def template(self, value):
        self._element.template_text = value

    @property
    def manager(self):
        """
        Manager of the document author. Read/write. Returns a string.
        """
        return self._element.manager_text

    @manager.setter
    def manager(self, value):
        self._element.manager_text = value

    @property
    def company(self):
        """
        Company name. Read/write. Returns a string.
        """
        return self._element.company_text

    @company.setter
    def company(self, value):
        self._element.company_text = value

    @property
    def application(self):
        """
        Application that created the document. Read/write. Returns a string.
        """
        return self._element.application_text

    @application.setter
    def application(self, value):
        self._element.application_text = value

    @property
    def app_version(self):
        """
        Version of the application. Read/write. Returns a string.
        """
        return self._element.appVersion_text

    @app_version.setter
    def app_version(self, value):
        self._element.appVersion_text = value

    @property
    def hyperlink_base(self):
        """
        Base URL for relative hyperlinks. Read/write. Returns a string.
        """
        return self._element.hyperlinkBase_text

    @hyperlink_base.setter
    def hyperlink_base(self, value):
        self._element.hyperlinkBase_text = value
