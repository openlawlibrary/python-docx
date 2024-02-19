from ..shared import Parented, find_document_parent

class Hyperlink(Parented):

    def __init__(self, h, parent):
        super(Hyperlink, self).__init__(parent)
        self._h = h
        self._parent = parent

    @property
    def text(self):
        """
        String formed by concatenating the text of each run in the hyperlink.
        """
        if self._h.r_lst is None:
            return ''
        text = ''
        for r in self._h.r_lst:
            text += r.text
        return text

    @property
    def link(self):
        """
        String that can be either an URL or an |Bookmark| name.
        """
        if self._h.relationship_id:
            return find_document_parent(self).part.rels[self._h.relationship_id].target_ref
        else:
            return self._h.anchor
