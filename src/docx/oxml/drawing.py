"""Custom element-classes for DrawingML-related elements like `<w:drawing>`.

For legacy reasons, many DrawingML-related elements are in `docx.oxml.shape`. Expect
those to move over here as we have reason to touch them.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List

from docx.oxml.xmlchemy import BaseOxmlElement

if TYPE_CHECKING:
    from docx.oxml.shape import CT_Blip


class CT_Drawing(BaseOxmlElement):
    """`<w:drawing>` element, containing a DrawingML object like a picture or chart."""

    @property
    def blip(self) -> CT_Blip | None:
        """The `a:blip` descendant of this drawing, or |None| if not present.

        This drawing's picture data is either inline (`wp:inline`) or floating
        (`wp:anchor`) but the `a:blip` identifying its image is nested the same number
        of levels deep either way, so this is found by tag rather than by structural
        path.
        """
        blips: List[CT_Blip] = self.xpath(".//a:blip")
        return blips[0] if blips else None
