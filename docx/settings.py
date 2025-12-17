# encoding: utf-8

"""Settings object, providing access to document-level settings"""

from __future__ import absolute_import, division, print_function, unicode_literals

from docx.shared import ElementProxy


class Settings(ElementProxy):
    """Provides access to document-level settings for a document.

    Accessed using the :attr:`.Document.settings` property.
    """

    __slots__ = ()

    @property
    def odd_and_even_pages_header_footer(self):
        """True if this document has distinct odd and even page headers and footers.

        Read/write.
        """
        return self._element.evenAndOddHeaders_val

    @odd_and_even_pages_header_footer.setter
    def odd_and_even_pages_header_footer(self, value):
        self._element.evenAndOddHeaders_val = value

    @property
    def footnote_position(self):
        """The document-level default footnote position.

        Can be 'pageBottom' or 'beneathText'.
        Returns None if not set. Read/write.
        """
        return self._element.footnote_position

    @footnote_position.setter
    def footnote_position(self, value):
        self._element.footnote_position = value

    @property
    def endnote_position(self):
        """The document-level default endnote position.

        Can be 'sectEnd' (end of section) or 'docEnd' (end of document).
        Returns None if not set. Read/write.
        """
        return self._element.endnote_position

    @endnote_position.setter
    def endnote_position(self, value):
        self._element.endnote_position = value
