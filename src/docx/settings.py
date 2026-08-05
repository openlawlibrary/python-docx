"""Settings object, providing access to document-level settings."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.shared import ElementProxy

if TYPE_CHECKING:
    import docx.types as t
    from docx.oxml.settings import CT_Settings
    from docx.oxml.xmlchemy import BaseOxmlElement


class Settings(ElementProxy):
    """Provides access to document-level settings for a document.

    Accessed using the :attr:`.Document.settings` property.
    """

    def __init__(self, element: BaseOxmlElement, parent: t.ProvidesXmlPart | None = None):
        super().__init__(element, parent)
        self._settings = cast("CT_Settings", element)

    @property
    def endnote_position(self) -> str | None:
        """Document-level default endnote position, e.g. `'docEnd'`.

        Read/write. |None| if not explicitly set, in which case each section's
        effective endnote position is determined by :attr:`.Section.endnote_position`.
        """
        return self._settings.endnote_position

    @endnote_position.setter
    def endnote_position(self, value: str | None):
        self._settings.endnote_position = value

    @property
    def footnote_position(self) -> str | None:
        """Document-level default footnote position, e.g. `'pageBottom'`.

        Read/write. |None| if not explicitly set, in which case each section's
        effective footnote position is determined by :attr:`.Section.footnote_position`.
        """
        return self._settings.footnote_position

    @footnote_position.setter
    def footnote_position(self, value: str | None):
        self._settings.footnote_position = value

    @property
    def odd_and_even_pages_header_footer(self) -> bool:
        """True if this document has distinct odd and even page headers and footers.

        Read/write.
        """
        return self._settings.evenAndOddHeaders_val

    @odd_and_even_pages_header_footer.setter
    def odd_and_even_pages_header_footer(self, value: bool):
        self._settings.evenAndOddHeaders_val = value
