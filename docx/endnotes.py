# encoding: utf-8

"""
Endnotes-related proxy types.
"""

from .shared import Parented
from .blkcntnr import BlockItemContainer


class Endnotes(Parented):
    """
    Proxy object wrapping ``<w:endnotes>`` element.
    """
    def __init__(self, endnotes, parent):
        super(Endnotes, self).__init__(parent)
        self._element = self._endnotes = endnotes

    def __getitem__(self, reference_id):
        """
        A |Endnote| for a specific endnote of reference id, defined with ``w:id`` argument of ``<w:endnoteReference>``.
        If reference id is invalid raises an |IndexError|
        """
        endnote = self._element.get_by_id(reference_id)
        if endnote is None:
            raise IndexError
        return Endnote(endnote, self)

    def __len__(self):
        return len(self._element)

    def add_endnote(self, endnote_reference_id):
        """
        Return a newly created |Endnote|, the new endnote will
        be inserted in the correct spot by `endnote_reference_id`.
        The endnotes are kept in order by `endnote_reference_id`.
        """
        elements = self._element # for easy access
        new_endnote = None
        if elements.get_by_id(endnote_reference_id) is not None:
            # When adding an endnote it can be inserted
            # in front of some other endnotes, so
            # we need to sort endnotes by `endnote_reference_id`
            # in |Endnotes| and in |Paragraph|
            #
            # resolve reference ids in |Endnotes|
            # iterate in reverse and compare the current
            # id with the inserted id. If there are the same
            # insert the new endnote in that place, if not
            # increment the current endnote id.
            for index in reversed(range(len(elements))):
                if elements[index].id == endnote_reference_id:
                    elements[index].id += 1
                    new_endnote = elements[index].add_endnote_before(endnote_reference_id)
                    break
                else:
                    elements[index].id += 1
        else:
            # append the newly created |Endnote| to |Endnotes|
            new_endnote = elements.add_endnote(endnote_reference_id)
        return Endnote(new_endnote, self)


class Endnote(BlockItemContainer):
    """
    Proxy object wrapping ``<w:endnote>`` element.
    """
    def __init__(self, e, parent):
        super(Endnote, self).__init__(e, parent)
        self._e = self._element = e

    def __eq__(self, other):
        if isinstance(other, Endnote):
            return self._e is other._e
        return False

    def __ne__(self, other):
        if isinstance(other, Endnote):
            return self._e is not other._e
        return True

    @property
    def id(self):
        return self._e.id
