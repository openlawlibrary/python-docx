
from .shared import ElementProxy

class SdtBase(ElementProxy):
    """
    Instance of base structured document tag (block content control).
    """
    def __init__(self, element, parent):
        super(SdtBase, self).__init__(element, parent)

    @property
    def content(self):
        return SdtContentBase(self._element.sdtContent, self)


class SdtContentBase(ElementProxy):
    def __init__(self, element, parent):
        super(SdtContentBase).__init__(element, parent)

    @property
    def text(self):
        text = ''
        for run in self._element.runs:
            text += run.text
        return text
