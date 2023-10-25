# encoding: utf-8

"""
Footnotes-related proxy types.
"""

from .shared import (
    ElementProxy, Parented
)
from .text.paragraph import Paragraph


class Footnotes(ElementProxy):
    """
    Proxy object wrapping ``<w:footnotes>`` element.
    """

    __slots__ = ()

    def __getitem__(self, reference_id):
        """
        A |Footnote| for a specific footnote of reference id, defined with ``w:id`` argument of ``<w:footnoteReference>``.
        If reference id is invalid raises an |IndexError|
        """
        footnote = self._element.get_by_id(reference_id)
        if footnote == None:
            raise IndexError
        return Footnote(footnote, self)

    def __len__(self):
        return len(self._element)


class Footnote(Parented):
    """
    Proxy object wrapping ``<w:footnote>`` element.
    """
    def __init__(self, f, parent):
        super(Footnote, self).__init__(parent)
        self._f = self._element = f

    @property
    def id(self):
        return self._f.id

    @property
    def paragraphs(self):
        """
        Returns a list of |Paragraph| that are defined in this footnote,
        or return |None| if there is no paragraph in footnote.
        """
        if self._f.paragraphs == None:
            return None
        return [Paragraph(p, self) for p in self._f.paragraphs]
