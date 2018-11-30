
from .shared import ElementProxy
from .text.paragraph import Paragraph

class SdtBase(ElementProxy):
    """
    Instance of base structured document tag.
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
    Instance of structured document tag properties.
    """
    def __init__(self, element, parent):
        super(SdtPr, self).__init__(element, parent)

    @property
    def name(self):
        return self._element.name

class SdtContentBase(ElementProxy):
    def __init__(self, element, parent):
        super(SdtContentBase, self).__init__(element, parent)

    @property
    def paragraphs(self):
        return [Paragraph(p, self) for p in self._element.p_lst]

    @property
    def text(self):
        text = ''
        for p in self.paragraphs:
            text += p.text
        return text
