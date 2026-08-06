"""Unit test suite for the `docx.sdt` module."""

from __future__ import annotations

from docx import Document


class DescribeSdt:
    """Integration-level test suite for content-control (structured document tag)
    support.

    Exercises the real `Document`/`Paragraph`/`SdtBase` object graph rather than
    mocks, since these span several collaborating proxy and oxml classes.
    """

    def it_can_add_a_content_control_to_the_document_body(self):
        document = Document()

        sdt = document.add_sdt("first_sdt")

        assert len(document.sdts) == 1
        assert document.sdts["first_sdt"] is not None
        assert sdt.name == "first_sdt"

    def it_can_add_multiple_content_controls_to_the_document_body(self):
        document = Document()
        document.add_sdt("first_sdt")
        document.add_sdt("second_sdt")

        assert len(document.sdts) == 2

    def it_can_nest_a_content_control_inside_another(self):
        document = Document()
        outer = document.add_sdt("outer_sdt")

        outer.add_sdt("inner_sdt")

        assert len(outer.sdts) == 1
        assert len(document.sdts) == 1
        assert len(document.sdts_all) == 2

    def it_can_add_a_paragraph_inside_a_content_control(self):
        document = Document()
        sdt = document.add_sdt("sdt")

        sdt.add_paragraph("paragraph one")
        sdt.add_paragraph("paragraph two")

        assert [p.text for p in sdt.paragraphs] == ["paragraph one", "paragraph two"]

    def it_can_add_a_table_inside_a_content_control(self):
        document = Document()
        sdt = document.add_sdt("sdt")

        sdt.add_table(rows=2, cols=2, width=90000)

        assert len(sdt.tables) == 1

    def it_creates_an_empty_content_control_with_placeholder_text_by_default(self):
        document = Document()
        p = document.add_paragraph()

        sdt = p.add_sdt("test_ord")

        assert sdt.text == "Click or tap here to enter text"
        assert sdt.properties.active_placeholder is True
        assert sdt.is_empty is True

    def it_creates_a_content_control_with_custom_placeholder_text(self):
        document = Document()
        p = document.add_paragraph()

        sdt = p.add_sdt("test_ord", placeholder_txt="double click to edit")

        assert sdt.text == "double click to edit"
        assert sdt.properties.active_placeholder is True

    def it_creates_a_content_control_with_actual_text(self):
        document = Document()
        p = document.add_paragraph()

        sdt = p.add_sdt("test_ord", text="some regular text")

        assert sdt.text == "some regular text"
        assert sdt.properties.active_placeholder is False
        assert sdt.is_empty is False

    def it_can_add_multiple_inline_content_controls_to_a_paragraph(self):
        document = Document()
        p = document.add_paragraph()

        p.add_sdt("first", text="one")
        p.add_run(" in between two content controls ")
        p.add_sdt("second", text="two")

        assert len(p.sdts) == 2
        assert len(document.sdts_all) == 2

    def it_includes_a_content_controls_text_in_its_paragraphs_text(self):
        document = Document()
        p = document.add_paragraph("before ")
        p.add_sdt("date", text="January 1, 2020")
        p.add_run(" after")

        assert p.text == "before January 1, 2020 after"

    def it_can_clear_the_content_of_a_content_control(self):
        document = Document()
        p = document.add_paragraph()
        sdt = p.add_sdt("test_ord", text="This is some text.")

        sdt.clear_content()

        assert sdt.text == ""
        assert len(document.sdts_all) == 1

    def it_can_clear_placeholder_text_from_a_content_control(self):
        document = Document()
        p = document.add_paragraph()
        sdt = p.add_sdt("test_ord")

        sdt.clear_content()

        assert sdt.text == ""
