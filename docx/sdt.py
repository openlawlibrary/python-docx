"""
|SdtBase| and closely related objects.
"""

from .shared import ElementProxy

from enum import Enum, auto
from .blkcntnr import BlockItemContainer

class SdtType(Enum):
    """
    Initial list of available Structure Document Types
    """
    RICH_TEXT = auto(),
    PLAIN_TEXT = auto(),
    DATE = auto(),
    DROP_DOWN = auto(),


class SdtBase(ElementProxy):
    """
    ``CT_SdtBase`` wrapper object, which contains references to
    parent object, content control content |SdtContentBase|
    and properties objects |SdtPr|.
    """
    def __init__(self, element, parent):
        super(SdtBase, self).__init__(element, parent)
        self.__content = None

    @property
    def _content(self):
        if self.__content is None:
            self.__content =_SdtContentBase(self._element.sdtContent, self)
        return self.__content

    @property
    def properties(self):
        return SdtPr(self._element.sdtPr, self)

    @property
    def is_empty(self):
        return any((self.properties.active_placeholder,
                    not self._content.text))

    @property
    def name(self):
        return self.properties.tag

    def add_sdt(self, tag_name):
        return self._content.add_sdt(tag_name)

    @property
    def sdts(self):
        return self._content.sdts

    def add_paragraph(self, text):
        return self._content.add_paragraph(text)

    @property
    def paragraphs(self):
        return self._content.paragraphs

    def add_table(self, rows, cols, width):
        return self._content.add_table(rows, cols, width)

    @property
    def tables(self):
        return self._content.tables

    @property
    def text(self):
        return self._content.text

class SdtPr(ElementProxy):
    """
    ``CT_SdtPr`` wrapper object which has references to
    parent object |SdtBase|.
    """
    def __init__(self, element, parent):
        super(SdtPr, self).__init__(element, parent)

    @property
    def tag(self):
        return self._element.name

    @property
    def active_placeholder(self):
        return self._element.active_placeholder is not None

class _SdtContentBase(BlockItemContainer):
    """
    ``CT_SdtContentBase`` wrapper object, which contains references
    to all paragraphs within the content control, and text property.
    """
    def __init__(self, element, parent):
        super(_SdtContentBase, self).__init__(element, parent)

    @property
    def text(self):
        text = ''
        for r in self._element.iter_runs():
            text += r.text
        return text
