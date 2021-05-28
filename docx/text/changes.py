from ..shared import Parented

class TrackChange(Parented):
    """
    Proxy object wrapping ``<w:ins>`` and ``<w:del>`` elements.
    """
    def __init__(self, p, parent):
        super(TrackChange, self).__init__(parent)

    @property
    def text(self):
        return self._r.text

