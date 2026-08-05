"""Unit-test suite for `docx.oxml.bookmark` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.bookmark import CT_BookmarkEnd, CT_BookmarkStart

from ..unitutil.cxml import element


class DescribeCT_BookmarkStart:
    """Unit-test suite for `docx.oxml.bookmark.CT_BookmarkStart`."""

    def it_provides_access_to_its_id_and_name(self):
        bookmarkStart = cast(CT_BookmarkStart, element("w:bookmarkStart{w:id=1,w:name=bmk1}"))
        assert bookmarkStart.id == 1
        assert bookmarkStart.name == "bmk1"


class DescribeCT_BookmarkEnd:
    """Unit-test suite for `docx.oxml.bookmark.CT_BookmarkEnd`."""

    def it_provides_access_to_its_id(self):
        bookmarkEnd = cast(CT_BookmarkEnd, element("w:bookmarkEnd{w:id=1}"))
        assert bookmarkEnd.id == 1
