# pyright: reportImportCycles=false

"""Block item container, used by body, cell, header, etc.

Block level items are things like paragraph and table, although there are a few other
specialized ones like structured document tags.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, List, cast

from typing_extensions import TypeAlias

from docx.bookmark import BookmarkParent
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import StoryChild
from docx.text.paragraph import Paragraph

if TYPE_CHECKING:
    import docx.types as t
    from docx.oxml.comments import CT_Comment
    from docx.oxml.document import CT_Body
    from docx.oxml.endnote import CT_FtnEdn
    from docx.oxml.footnote import CT_FtnEnd
    from docx.oxml.sdts import CT_SdtBase, CT_SdtContentBase
    from docx.oxml.section import CT_HdrFtr
    from docx.oxml.table import CT_Tc
    from docx.sdt import SdtBase
    from docx.section import Sections
    from docx.shared import Length
    from docx.styles.style import ParagraphStyle
    from docx.table import Table

BlockItemElement: TypeAlias = (
    "CT_Body | CT_Comment | CT_FtnEdn | CT_FtnEnd | CT_HdrFtr | CT_SdtContentBase | CT_Tc"
)


class BlockItemContainer(StoryChild, BookmarkParent):
    """Base class for proxy objects that can contain block items.

    These containers include _Body, _Cell, header, footer, footnote, endnote, comment,
    and text box objects. Provides the shared functionality to add a block item like a
    paragraph or table.
    """

    def __init__(self, element: BlockItemElement, parent: t.ProvidesStoryPart):
        super(BlockItemContainer, self).__init__(parent)
        self._element = element

    def add_paragraph(self, text: str = "", style: str | ParagraphStyle | None = None) -> Paragraph:
        """Return paragraph newly added to the end of the content in this container.

        The paragraph has `text` in a single run if present, and is given paragraph
        style `style`.

        If `style` is |None|, no paragraph style is applied, which has the same effect
        as applying the 'Normal' style.
        """
        paragraph = self._add_paragraph()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    def add_sdt(self, tag_name: str, alias_name: str = "") -> SdtBase:
        """Return a content control (structured document tag) newly added to the end
        of the content in this container, tagged `tag_name`.

        `alias_name` defaults to `tag_name` when not provided.
        """
        from docx.sdt import SdtBase

        # -- `_new_sdt()` is present on `CT_Body`, `CT_HdrFtr`, and `CT_SdtContentBase`,
        # -- but not `CT_Comment`, `CT_Tc`, `CT_FtnEnd`, or `CT_FtnEdn` -- content
        # -- controls aren't supported inside comments, table cells, footnotes, or
        # -- endnotes, matching the original fork's exact scope.
        sdt = cast("CT_SdtBase", self._element._new_sdt())  # pyright: ignore
        sdtPr = sdt.get_or_add_sdtPr()
        sdtPr.name = tag_name
        sdtPr.alias_val = alias_name or tag_name
        sdt.get_or_add_sdtContent()
        self._element.append(sdt)
        return SdtBase(sdt, self)

    def add_table(self, rows: int, cols: int, width: Length) -> Table:
        """Return table of `width` having `rows` rows and `cols` columns.

        The table is appended appended at the end of the content in this container.

        `width` is evenly distributed between the table columns.
        """
        from docx.table import Table

        tbl = CT_Tbl.new_tbl(rows, cols, width)
        self._element._insert_tbl(tbl)  # pyright: ignore[reportPrivateUsage]
        return Table(tbl, self)

    def iter_inner_content(self) -> Iterator[Paragraph | Table]:
        """Generate each `Paragraph` or `Table` in this container in document order."""
        from docx.table import Table

        for element in self._element.inner_content_elements:
            yield (Paragraph(element, self) if isinstance(element, CT_P) else Table(element, self))

    @property
    def paragraphs(self):
        """A list containing the paragraphs in this container, in document order.

        Read-only.
        """
        return [Paragraph(p, self) for p in self._element.p_lst]

    @property
    def sdts(self) -> dict[str | None, SdtBase]:
        """The content controls (structured document tags) directly contained in this
        container, in document order, keyed by tag name.

        Read-only.
        """
        from docx.sdt import SdtBase

        return {sdt.name: SdtBase(sdt, self) for sdt in self._iter_sdts()}

    @property
    def sdts_all(self) -> dict[str | None, SdtBase]:
        """The content controls (structured document tags) contained anywhere in this
        container -- including nested inside other content controls, and, for the
        document body, in its headers and footers -- keyed by tag name.

        Read-only.
        """
        from docx.sdt import SdtBase

        return {sdt.name: SdtBase(sdt, self) for sdt in self._iter_sdts_all()}

    @property
    def tables(self):
        """A list containing the tables in this container, in document order.

        Read-only.
        """
        from docx.table import Table

        return [Table(tbl, self) for tbl in self._element.tbl_lst]

    def _add_paragraph(self):
        """Return paragraph newly added to the end of the content in this container."""
        return Paragraph(self._element.add_p(), self)

    def _iter_sdts(self) -> Iterator[CT_SdtBase]:
        """Generate each `w:sdt` element directly contained in this container."""
        yield from cast("List[CT_SdtBase]", self._element.sdt_lst)  # pyright: ignore

    def _iter_sdts_all(self) -> Iterator[CT_SdtBase]:
        """Generate each `w:sdt` element contained anywhere in this container.

        This includes content controls nested inside other content controls and,
        when this container is the document body, those in its headers and footers.
        """
        # -- `self._parent.sections` is only present when this container is the
        # -- document body (`_Body`, whose parent is `Document`); nested content
        # -- controls and other container types don't support this full-document scan.
        sections = cast("Sections", self._parent.sections)  # pyright: ignore
        for section in sections:
            for hdr_ftr in (
                section.header,
                section.first_page_header,
                section.even_page_header,
                section.footer,
                section.first_page_footer,
                section.even_page_footer,
            ):
                yield from cast(
                    "Iterator[CT_SdtBase]",
                    hdr_ftr._element.iterdescendants(qn("w:sdt")),  # pyright: ignore[reportPrivateUsage]
                )
        yield from cast("Iterator[CT_SdtBase]", self._element.iterdescendants(qn("w:sdt")))
