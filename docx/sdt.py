"""
|SdtBase| and closely related objects.
"""

from .shared import ElementProxy
from .text.paragraph import Paragraph

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
    def name(self):
        return self.properties.name

class SdtPr(ElementProxy):
    """
    ``CT_SdtPr`` wrapper object which has references to
    parent object |SdtBase|.
    """
    def __init__(self, element, parent):
        super(SdtPr, self).__init__(element, parent)

    @property
    def name(self):
        return self._element.name

class SdtContentBase(ElementProxy):
    """
    ``CT_SdtContentBase`` wrapper object, which contains references
    to all paragraphs within the content control, and text property.
    """
    def __init__(self, element, parent):
        super(SdtContentBase, self).__init__(element, parent)

    @property
    def paragraphs(self):
        return [Paragraph(p, self) for p in self._element.p_lst]

    @property
    def text(self):
        text = ''
        for r in self._element.get_runs():
            text += r.text
        return text
