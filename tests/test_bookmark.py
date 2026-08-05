"""Unit test suite for the `docx.bookmark` module."""

from __future__ import annotations

import pytest

from docx import Document


class DescribeBookmarks:
    """Integration-level test suite for bookmark support.

    Exercises the real `Document`/`Paragraph`/`Run`/`Table` object graph rather than
    mocks, since `Bookmarks` walks the actual part/relationship tree to find and
    number bookmarks.
    """

    def it_starts_with_no_bookmarks_in_a_default_document(self):
        document = Document()
        assert len(document.bookmarks) == 0

    def it_can_start_and_end_a_bookmark_at_the_document_level(self):
        document = Document()
        document.add_paragraph("paragraph one")

        bookmark = document.start_bookmark("bmk1")
        document.add_paragraph("paragraph two")
        document.end_bookmark(bookmark)

        assert bookmark.name == "bmk1"
        assert bookmark.id == 0
        assert len(document.bookmarks) == 1
        assert document.bookmarks.get("bmk1") is not None

    def it_raises_on_a_duplicate_bookmark_name(self):
        document = Document()
        bookmark = document.start_bookmark("bmk1")
        document.end_bookmark(bookmark)

        with pytest.raises(KeyError):
            document.start_bookmark("bmk1")

    def it_raises_getting_a_bookmark_by_a_name_not_present(self):
        document = Document()
        with pytest.raises(KeyError):
            document.bookmarks.get("does-not-exist")

    def it_assigns_the_lowest_available_id(self):
        document = Document()
        bmk1 = document.start_bookmark("bmk1")
        document.end_bookmark(bmk1)
        bmk2 = document.start_bookmark("bmk2")
        document.end_bookmark(bmk2)

        assert (bmk1.id, bmk2.id) == (0, 1)

    def it_can_start_and_end_a_bookmark_within_a_single_paragraph(self):
        document = Document()
        p = document.add_paragraph("test paragraph")

        bookmark = p.start_bookmark("bmk1")
        p.add_run(" more text")
        p.end_bookmark(bookmark)

        assert len(p.bookmark_starts) == 1
        assert len(p.bookmark_ends) == 1
        assert len(document.bookmarks) == 1

    def it_can_start_and_end_a_bookmark_spanning_paragraphs(self):
        document = Document()
        p1 = document.add_paragraph("paragraph one")

        bookmark = p1.start_bookmark("bmk1")
        p2 = document.add_paragraph("paragraph two")
        p2.end_bookmark(bookmark)

        assert len(p1.bookmark_starts) == 1
        assert len(p2.bookmark_ends) == 1
        assert len(document.bookmarks) == 1

    def it_can_start_and_end_a_bookmark_at_the_run_level(self):
        document = Document()
        p = document.add_paragraph()
        r1 = p.add_run("run one")

        bookmark = r1.start_bookmark("bmk1")
        r2 = p.add_run("run two")
        r2.end_bookmark(bookmark)

        assert len(document.bookmarks) == 1

    def it_can_start_and_end_a_bookmark_at_the_table_level(self):
        document = Document()
        table = document.add_table(rows=1, cols=1)

        bookmark = table.start_bookmark("bmk1")
        table.end_bookmark(bookmark)

        assert len(table.bookmark_starts) == 1
        assert len(table.bookmark_ends) == 1
        assert len(document.bookmarks) == 1

    def it_ignores_an_open_bookmark_with_no_matching_end(self):
        document = Document()
        document.start_bookmark("bmk1")

        assert len(document.bookmarks) == 0

    def it_supports_indexed_and_sliced_access(self):
        document = Document()
        bmk1 = document.start_bookmark("bmk1")
        document.end_bookmark(bmk1)
        bmk2 = document.start_bookmark("bmk2")
        document.end_bookmark(bmk2)

        bookmarks = document.bookmarks

        assert bookmarks[0].name == "bmk1"
        assert [b.name for b in bookmarks[0:2]] == ["bmk1", "bmk2"]

    def it_supports_iteration(self):
        document = Document()
        bmk1 = document.start_bookmark("bmk1")
        document.end_bookmark(bmk1)
        bmk2 = document.start_bookmark("bmk2")
        document.end_bookmark(bmk2)

        names = [bookmark.name for bookmark in document.bookmarks]

        assert names == ["bmk1", "bmk2"]
