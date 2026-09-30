from ..shared import Parented

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
        String that can be either an URL or an |Bookmark| name. The URL is a
        relationship of the part that holds the hyperlink, e.g. the footnotes
        part for a hyperlink in a footnote.
        """
        if self._h.relationship_id:
            return self.part.rels[self._h.relationship_id].target_ref
        else:
            return self._h.anchor
