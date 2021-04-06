"""
|SdtBase| and closely related objects.
"""

import docx.text.paragraph

from .shared import ElementProxy

from enum import Enum, auto

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

    @property
    def content(self):
        return SdtContentBase(self._element.sdtContent, self)

    @property
    def properties(self):
        return SdtPr(self._element.sdtPr, self)

    @property
    def is_empty(self):
        return any((self.properties.active_placeholder,
                    not self.content.text))

    @property
    def name(self):
        return self.properties.tag

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

class SdtContentBase(ElementProxy):
    """
    ``CT_SdtContentBase`` wrapper object, which contains references
    to all paragraphs within the content control, and text property.
    """
    def __init__(self, element, parent):
        super(SdtContentBase, self).__init__(element, parent)

    @property
    def paragraphs(self):
        return [docx.text.paragraph.Paragraph(p, self) for p in self._element.p_lst]

    @property
    def text(self):
        text = ''
        for r in self._element.iter_runs():
            text += r.text
        return text
