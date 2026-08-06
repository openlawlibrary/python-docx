"""|DocumentPart| and closely related objects."""

from __future__ import annotations

from itertools import chain
from typing import IO, TYPE_CHECKING, Iterator, cast

from docx.document import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.parts.comments import CommentsPart
from docx.parts.endnotes import EndnotesPart
from docx.parts.footnotes import FootnotesPart
from docx.parts.hdrftr import FooterPart, HeaderPart
from docx.parts.numbering import NumberingPart
from docx.parts.settings import SettingsPart
from docx.parts.story import StoryPart
from docx.parts.styles import StylesPart
from docx.shape import InlineShapes
from docx.shared import lazyproperty

if TYPE_CHECKING:
    from docx.comments import Comments
    from docx.endnotes import Endnotes
    from docx.enum.style import WD_STYLE_TYPE
    from docx.footnotes import Footnotes
    from docx.opc.coreprops import CoreProperties
    from docx.opc.customprops import CustomProperties
    from docx.opc.extendedprops import ExtendedProperties
    from docx.opc.part import Part
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.settings import Settings
    from docx.styles.style import BaseStyle


class DocumentPart(StoryPart):
    """Main document part of a WordprocessingML (WML) package, aka a .docx file.

    Acts as broker to other parts such as image, core properties, and style parts. It
    also acts as a convenient delegate when a mid-document object needs a service
    involving a remote ancestor. The `Parented.part` property inherited by many content
    objects provides access to this part object for that purpose.
    """

    def add_footer_part(self):
        """Return (footer_part, rId) pair for newly-created footer part."""
        footer_part = FooterPart.new(self.package)
        rId = self.relate_to(footer_part, RT.FOOTER)
        return footer_part, rId

    def add_header_part(self):
        """Return (header_part, rId) pair for newly-created header part."""
        header_part = HeaderPart.new(self.package)
        rId = self.relate_to(header_part, RT.HEADER)
        return header_part, rId

    @lazyproperty
    def cached_styles(self) -> dict[str, BaseOxmlElement]:
        """A `{style_id: w:style-element}` mapping for every style in this document.

        Building the |Styles| collection constructs a new proxy object for every
        style on each access, which is too expensive to repeat for every paragraph
        queried during numbering resolution -- this lazy, one-shot cache avoids that.
        Like the numbering part's own resolution caches, this assumes styles are not
        added or changed during a single read of the document.
        """
        return {
            s.style_id: cast("BaseOxmlElement", s._element)  # pyright: ignore[reportPrivateUsage]
            for s in self.styles
        }

    @property
    def comments(self) -> Comments:
        """|Comments| object providing access to the comments added to this document."""
        return self._comments_part.comments

    @property
    def core_properties(self) -> CoreProperties:
        """A |CoreProperties| object providing read/write access to the core properties
        of this document."""
        return self.package.core_properties

    @property
    def custom_properties(self) -> CustomProperties:
        """A |CustomProperties| object providing read/write access to the custom
        properties of this document."""
        return self.package.custom_properties

    @property
    def document(self):
        """A |Document| object providing access to the content of this document."""
        return Document(self._element, self)

    @property
    def extended_properties(self) -> ExtendedProperties:
        """An |ExtendedProperties| object providing read/write access to the extended
        properties of this document."""
        return self.package.extended_properties

    def drop_header_part(self, rId: str) -> None:
        """Remove related header part identified by `rId`."""
        self.drop_rel(rId)

    @property
    def endnotes(self) -> Endnotes:
        """|Endnotes| object providing access to the endnotes added to this document."""
        return self._endnotes_part.endnotes

    def footer_part(self, rId: str):
        """Return |FooterPart| related by `rId`."""
        return self.related_parts[rId]

    @property
    def footnotes(self) -> Footnotes:
        """|Footnotes| object providing access to the footnotes added to this document."""
        return self._footnotes_part.footnotes

    def get_style(self, style_id: str | None, style_type: WD_STYLE_TYPE) -> BaseStyle:
        """Return the style in this document matching `style_id`.

        Returns the default style for `style_type` if `style_id` is |None| or does not
        match a defined style of `style_type`.
        """
        return self.styles.get_by_id(style_id, style_type)

    def get_style_id(self, style_or_name, style_type):
        """Return the style_id (|str|) of the style of `style_type` matching
        `style_or_name`.

        Returns |None| if the style resolves to the default style for `style_type` or if
        `style_or_name` is itself |None|. Raises if `style_or_name` is a style of the
        wrong type or names a style not present in the document.
        """
        return self.styles.get_style_id(style_or_name, style_type)

    def header_part(self, rId: str):
        """Return |HeaderPart| related by `rId`."""
        return self.related_parts[rId]

    @lazyproperty
    def inline_shapes(self):
        """The |InlineShapes| instance containing the inline shapes in the document."""
        return InlineShapes(self._element.body, self)

    def iter_story_parts(self) -> Iterator[Part]:
        """Generate each story part of this document.

        A story is a sequence of block-level items (paragraphs and tables). Story parts
        include this main document part, headers, footers, footnotes, comments, and
        endnotes.
        """
        return chain(
            (self,),
            self.iter_parts_related_by(
                {RT.COMMENTS, RT.ENDNOTES, RT.FOOTER, RT.FOOTNOTES, RT.HEADER}
            ),
        )

    @lazyproperty
    def numbering_part(self) -> NumberingPart:
        """A |NumberingPart| object providing access to the numbering definitions for this document.

        Creates an empty numbering part if one is not present.
        """
        try:
            return cast(NumberingPart, self.part_related_by(RT.NUMBERING))
        except KeyError:
            numbering_part = NumberingPart.new()
            self.relate_to(numbering_part, RT.NUMBERING)
            return numbering_part

    def save(self, path_or_stream: str | IO[bytes]):
        """Save this document to `path_or_stream`, which can be either a path to a
        filesystem location (a string) or a file-like object."""
        self.package.save(path_or_stream)

    @property
    def settings(self) -> Settings:
        """A |Settings| object providing access to the settings in the settings part of
        this document."""
        return self._settings_part.settings

    @property
    def styles(self):
        """A |Styles| object providing access to the styles in the styles part of this
        document."""
        return self._styles_part.styles

    @property
    def _comments_part(self) -> CommentsPart:
        """A |CommentsPart| object providing access to the comments added to this document.

        Creates a default comments part if one is not present.
        """
        try:
            return cast(CommentsPart, self.part_related_by(RT.COMMENTS))
        except KeyError:
            assert self.package is not None
            comments_part = CommentsPart.default(self.package)
            self.relate_to(comments_part, RT.COMMENTS)
            return comments_part

    @property
    def _endnotes_part(self) -> EndnotesPart:
        """An |EndnotesPart| object providing access to the endnotes added to this
        document.

        Creates a default endnotes part if one is not present.
        """
        try:
            return cast(EndnotesPart, self.part_related_by(RT.ENDNOTES))
        except KeyError:
            assert self.package is not None
            endnotes_part = EndnotesPart.default(self.package)
            self.relate_to(endnotes_part, RT.ENDNOTES)
            return endnotes_part

    @property
    def _footnotes_part(self) -> FootnotesPart:
        """A |FootnotesPart| object providing access to the footnotes added to this
        document.

        Creates a default footnotes part if one is not present.
        """
        try:
            return cast(FootnotesPart, self.part_related_by(RT.FOOTNOTES))
        except KeyError:
            assert self.package is not None
            footnotes_part = FootnotesPart.default(self.package)
            self.relate_to(footnotes_part, RT.FOOTNOTES)
            return footnotes_part

    @property
    def _settings_part(self) -> SettingsPart:
        """A |SettingsPart| object providing access to the document-level settings for
        this document.

        Creates a default settings part if one is not present.
        """
        try:
            return cast(SettingsPart, self.part_related_by(RT.SETTINGS))
        except KeyError:
            settings_part = SettingsPart.default(self.package)
            self.relate_to(settings_part, RT.SETTINGS)
            return settings_part

    @property
    def _styles_part(self) -> StylesPart:
        """Instance of |StylesPart| for this document.

        Creates an empty styles part if one is not present.
        """
        try:
            return cast(StylesPart, self.part_related_by(RT.STYLES))
        except KeyError:
            package = self.package
            assert package is not None
            styles_part = StylesPart.default(package)
            self.relate_to(styles_part, RT.STYLES)
            return styles_part
