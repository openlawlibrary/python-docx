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
        String that can be either an URL or an |Bookmark| name. The URL is a
        relationship of the part that holds the hyperlink, e.g. the footnotes
        part for a hyperlink in a footnote. Earlier versions wrote the
        relationship of each hyperlink to the main document part, so the URL
        of a hyperlink whose part does not have the relationship comes from
        there.
        """
        rId = self._h.relationship_id
        if not rId:
            return self._h.anchor
        rels = self.part.rels
        if rId not in rels:
            rels = find_containing_document(self).part.rels
        return rels[rId].target_ref
