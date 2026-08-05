"""|SdtBase| and closely related objects."""

from __future__ import annotations

from enum import Enum, auto
from typing import TYPE_CHECKING

from docx.blkcntnr import BlockItemContainer
from docx.shared import StoryChild

if TYPE_CHECKING:
    import docx.types as t
    from docx.oxml.sdts import CT_SdtBase, CT_SdtContentBase, CT_SdtPr
    from docx.shared import Length


class SdtType(Enum):
    """Initial list of available Structured Document Tag types."""

    RICH_TEXT = auto()
    PLAIN_TEXT = auto()
    DATE = auto()
    DROP_DOWN = auto()


class SdtBase(StoryChild):
    """Proxy for a `w:sdt` element, a structured document tag (content control).

    Provides access to the content control's properties (|SdtPr|) and content
    (|_SdtContentBase|).
    """

    def __init__(self, element: CT_SdtBase, parent: t.ProvidesStoryPart):
        super().__init__(parent)
        self._sdt = self._element = element
        self.__content: _SdtContentBase | None = None

    def add_paragraph(self, text: str = "", style: str | None = None):
        """Return paragraph newly added to the end of this content control's content."""
        return self._content.add_paragraph(text, style)

    def add_sdt(self, tag_name: str, alias_name: str = "") -> SdtBase:
        """Return a new content control tagged `tag_name`, nested in this one."""
        return self._content.add_sdt(tag_name, alias_name)

    def add_table(self, rows: int, cols: int, width: int | Length):
        """Return a table newly added to the end of this content control's content."""
        return self._content.add_table(rows, cols, width)  # pyright: ignore[reportArgumentType]

    def clear_content(self) -> None:
        """Remove all text from this content control's content, including placeholder
        text, without removing the content-holding elements themselves."""
        self._content.clear_content()

    @property
    def is_empty(self) -> bool:
        """True when this content control is showing placeholder text, or has no
        text content at all."""
        return self.properties.active_placeholder or not self._content.text

    @property
    def name(self) -> str | None:
        """The tag name of this content control."""
        return self.properties.tag

    @property
    def paragraphs(self):
        """The paragraphs directly contained in this content control's content."""
        return self._content.paragraphs

    @property
    def properties(self) -> SdtPr:
        """|SdtPr| object providing access to this content control's properties."""
        return SdtPr(self._sdt.get_or_add_sdtPr(), self)

    @property
    def sdts(self):
        """The content controls directly nested in this content control's content."""
        return self._content.sdts

    @property
    def tables(self):
        """The tables directly contained in this content control's content."""
        return self._content.tables

    @property
    def text(self) -> str:
        """The concatenated text of all runs in this content control's content."""
        return self._content.text

    @property
    def _content(self) -> _SdtContentBase:
        """|_SdtContentBase| object providing access to this content control's content."""
        if self.__content is None:
            self.__content = _SdtContentBase(self._sdt.get_or_add_sdtContent(), self)
        return self.__content


class SdtPr(StoryChild):
    """Proxy for a `w:sdtPr` element, the properties of a content control."""

    def __init__(self, element: CT_SdtPr, parent: t.ProvidesStoryPart):
        super().__init__(parent)
        self._sdtPr = self._element = element

    @property
    def active_placeholder(self) -> bool:
        """True when this content control is showing its placeholder text."""
        return self._sdtPr.active_placeholder

    @property
    def tag(self) -> str | None:
        """The tag name of this content control."""
        return self._sdtPr.name


class _SdtContentBase(BlockItemContainer):
    """Proxy for a `w:sdtContent` element, the content of a content control."""

    def __init__(self, element: CT_SdtContentBase, parent: t.ProvidesStoryPart):
        super().__init__(element, parent)
        self._sdtContent = element

    def clear_content(self) -> None:
        """Remove all text from this content, including placeholder text."""
        for r in self._sdtContent.iter_runs():
            r.clear_content()

    @property
    def text(self) -> str:
        """The concatenated text of all runs in this content, at any depth."""
        return "".join(r.text for r in self._sdtContent.iter_runs())
