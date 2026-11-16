from ..oxml.ns import qn
from ..shared import Parented, find_containing_document

#: A run in one of these elements is a run of the hyperlink. Tracked changes
#: are read as accepted, as for a paragraph: ``<w:del>`` and ``<w:moveFrom>``
#: are not here.
_RUN_CONTAINERS = (
    qn('w:sdt'), qn('w:sdtContent'), qn('w:smartTag'), qn('w:ins'), qn('w:moveTo'),
)


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
        def run_texts(element):
            for child in element:
                if child.tag == qn('w:r'):
                    yield child.text
                elif child.tag in _RUN_CONTAINERS:
                    yield from run_texts(child)

        return ''.join(run_texts(self._h))

    @property
    def link(self):
        """
        String that can be either an URL or an |Bookmark| name.
        """
        if self._h.relationship_id:
            return find_containing_document(self).part.rels[self._h.relationship_id].target_ref
        else:
            return self._h.anchor
