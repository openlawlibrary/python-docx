# pyright: reportPrivateUsage=false

"""Objects related to bookmarks."""

from __future__ import annotations

from collections.abc import Sequence
from itertools import chain
from typing import TYPE_CHECKING, Iterator, cast, overload

from typing_extensions import Protocol

from docx.oxml.ns import qn
from docx.shared import lazyproperty

if TYPE_CHECKING:
    from docx.opc.part import Part
    from docx.oxml.bookmark import CT_BookmarkEnd, CT_BookmarkStart


class _ProvidesStoryParts(Protocol):
    """An object, such as `DocumentPart`, that can enumerate the story parts of a
    document (main document story, headers, footers, footnotes, endnotes)."""

    def iter_story_parts(self) -> Iterator[Part]: ...


class _BookmarkableElement(Protocol):
    """An oxml element that can directly contain `w:bookmarkStart`/`w:bookmarkEnd`."""

    def _add_bookmarkStart(self) -> CT_BookmarkStart: ...
    def _add_bookmarkEnd(self) -> CT_BookmarkEnd: ...


class _BookmarkParentHost(Protocol):
    """The attributes `BookmarkParent` methods need from the class they're mixed into."""

    @property
    def _element(self) -> _BookmarkableElement: ...

    @property
    def part(self) -> _ProvidesStoryParts: ...


class BookmarkParent:
    """Mixin providing `start_bookmark()`/`end_bookmark()` for a block-item parent.

    Used by the proxy classes (`Paragraph`, `Run`, `Table`, `BlockItemContainer`) whose
    XML elements can directly contain a `w:bookmarkStart`/`w:bookmarkEnd`.
    """

    def start_bookmark(self, name: str) -> _Bookmark:
        """Start a bookmark named `name` at this location.

        Raises `KeyError` if a bookmark named `name` is already present in the
        document.
        """
        # -- `self.part` is statically typed as the abstract `StoryPart` by the proxy
        # -- classes this mixes into, but `Bookmarks` needs the `DocumentPart` it
        # -- actually is when the content lives in the main document body (as opposed
        # -- to a header, footer, or comment, where bookmark support was never fully
        # -- wired up in the original feature this was ported from either).
        host = cast(_BookmarkParentHost, self)
        bookmarks = Bookmarks(host.part)
        if name in (bookmark.name for bookmark in bookmarks):
            raise KeyError("Bookmark name already present in document.")

        # -- id must be computed before the new (as yet id-less) bookmarkStart is
        # -- inserted into the tree, since `next_id` scans for existing ids --
        next_id = bookmarks.next_id
        bookmarkStart = host._element._add_bookmarkStart()
        bookmarkStart.name = name
        bookmarkStart.id = next_id
        return _Bookmark((bookmarkStart, None))

    def end_bookmark(self, bookmark: _Bookmark) -> _Bookmark:
        """Close `bookmark` at this location."""
        host = cast(_BookmarkParentHost, self)
        bookmarkEnd = host._element._add_bookmarkEnd()
        bookmarkEnd.id = bookmark.id
        bookmark._bookmarkEnd = bookmarkEnd
        return bookmark


class Bookmarks(Sequence["_Bookmark"]):
    """Sequence of |Bookmark| objects for a document.

    Supports indexed access (including slices), `len()`, and iteration. Iteration
    performs significantly better than repeated indexed access.
    """

    def __init__(self, document_part: _ProvidesStoryParts):
        self._document_part = document_part

    @overload
    def __getitem__(self, idx: int) -> _Bookmark: ...
    @overload
    def __getitem__(self, idx: slice) -> list[_Bookmark]: ...

    def __getitem__(self, idx: int | slice) -> _Bookmark | list[_Bookmark]:
        """Supports indexed and sliced access."""
        bookmark_pairs = self._finder.bookmark_pairs
        if isinstance(idx, slice):
            return [_Bookmark(pair) for pair in bookmark_pairs[idx]]
        return _Bookmark(bookmark_pairs[idx])

    def __iter__(self) -> Iterator[_Bookmark]:
        """Supports iteration."""
        return (_Bookmark(pair) for pair in self._finder.bookmark_pairs)

    def __len__(self) -> int:
        return len(self._finder.bookmark_pairs)

    @lazyproperty
    def _finder(self) -> _DocumentBookmarkFinder:
        """`_DocumentBookmarkFinder` instance for this document."""
        return _DocumentBookmarkFinder(self._document_part)

    def get(self, name: str) -> _Bookmark:
        """Get bookmark based on its name.

        Raises `KeyError` when no bookmark named `name` is present.
        """
        for bookmark in self:
            if bookmark.name == name:
                return bookmark
        raise KeyError("Requested bookmark not found.")

    @property
    def next_id(self) -> int:
        """Lowest available bookmark id."""
        ids = {bmk.id for _, bmk in self._finder.bookmark_starts}
        id_range = set(range(len(ids) + 1))
        return min(id_range - ids)


class _Bookmark:
    """Proxy for a (`w:bookmarkStart`, `w:bookmarkEnd`) element pair."""

    def __init__(self, bookmark_pair: tuple[CT_BookmarkStart, CT_BookmarkEnd | None]) -> None:
        self._bookmarkStart, self._bookmarkEnd = bookmark_pair

    @property
    def id(self) -> int:
        """The bookmark id."""
        return self._bookmarkStart.id

    @property
    def name(self) -> str:
        """The bookmark name."""
        return self._bookmarkStart.name


class _DocumentBookmarkFinder:
    """Provides access to bookmark oxml elements in an overall document."""

    def __init__(self, document_part: _ProvidesStoryParts):
        self._document_part = document_part

    @property
    def bookmark_pairs(self) -> list[tuple[CT_BookmarkStart, CT_BookmarkEnd]]:
        """List of (bookmarkStart, bookmarkEnd) element pairs for document.

        All story parts of the document are searched, including the main document
        story, headers, footers, footnotes, and endnotes. The order of part searching
        is not guaranteed, but bookmarks appear in document order within a particular
        part. Only well-formed bookmarks appear. Any open bookmarks (start but no end),
        reversed bookmarks (end before start), or duplicate (name same as prior
        bookmark) bookmarks are ignored.
        """
        return list(
            chain.from_iterable(
                _PartBookmarkFinder.iter_start_end_pairs(part)
                for part in self._document_part.iter_story_parts()
            )
        )

    @property
    def bookmark_starts(self) -> list[tuple[int, CT_BookmarkStart]]:
        """List of (idx, bookmarkStart) pairs for document."""
        return list(
            chain.from_iterable(
                _PartBookmarkFinder.iter_starts(part)
                for part in self._document_part.iter_story_parts()
            )
        )


class _PartBookmarkFinder:
    """Provides access to bookmark oxml elements in a story part."""

    def __init__(self, part: Part) -> None:
        self._part = part

    @classmethod
    def iter_start_end_pairs(cls, part: Part) -> Iterator[tuple[CT_BookmarkStart, CT_BookmarkEnd]]:
        """Generate each (bookmarkStart, bookmarkEnd) in `part`."""
        return cls(part)._iter_start_end_pairs()

    @classmethod
    def iter_starts(cls, part: Part) -> Iterator[tuple[int, CT_BookmarkStart]]:
        """Generate each (idx, bookmarkStart) in `part`."""
        return cls(part)._iter_starts()

    def _iter_start_end_pairs(self) -> Iterator[tuple[CT_BookmarkStart, CT_BookmarkEnd]]:
        """Generate each (bookmarkStart, bookmarkEnd) in this part."""
        for idx, bookmarkStart in self._iter_starts():
            bookmarkEnd = self._matching_end(bookmarkStart, idx)
            # -- skip open pairs --
            if bookmarkEnd is None:
                continue
            # -- skip duplicate names --
            if self._name_already_used(bookmarkStart.name):
                continue
            yield (bookmarkStart, bookmarkEnd)

    @lazyproperty
    def _all_starts_and_ends(self) -> list[CT_BookmarkStart | CT_BookmarkEnd]:
        """List of all `w:bookmarkStart` and `w:bookmarkEnd` elements in part.

        Elements appear in document order.
        """
        part_el = getattr(self._part, "element", None)
        if part_el is None:
            return []
        return part_el.xpath("//w:bookmarkStart|//w:bookmarkEnd")

    def _iter_starts(self) -> Iterator[tuple[int, CT_BookmarkStart]]:
        """Generate (idx, bookmarkStart) elements in story.

        `idx` indicates the location of the bookmarkStart element among all the
        bookmarkStart and bookmarkEnd elements in the story.
        """
        for idx, element in enumerate(self._all_starts_and_ends):
            if element.tag == qn("w:bookmarkStart"):
                yield idx, cast("CT_BookmarkStart", element)

    def _matching_end(self, bookmarkStart: CT_BookmarkStart, idx: int) -> CT_BookmarkEnd | None:
        """The `w:bookmarkEnd` element corresponding to `bookmarkStart`.

        Returns `None` if no `w:bookmarkEnd` with matching id value is found. `idx` is
        the offset of `bookmarkStart` in the sequence of start and end elements in this
        story.
        """
        for element in self._all_starts_and_ends[idx + 1 :]:
            # -- skip bookmark starts --
            if element.tag == qn("w:bookmarkStart"):
                continue
            bookmarkEnd = cast("CT_BookmarkEnd", element)
            if bookmarkEnd.id == bookmarkStart.id:
                return bookmarkEnd
        return None

    def _name_already_used(self, name: str) -> bool:
        """True when bookmark `name` was already encountered, False otherwise."""
        names_so_far = self._names_so_far
        if name in names_so_far:
            return True
        names_so_far.add(name)
        return False

    @lazyproperty
    def _names_so_far(self) -> set[str]:
        """Set composed to track bookmark names encountered in document traversal."""
        return set()
