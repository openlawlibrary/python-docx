"""Unit test suite for the `docx.opc.non_phys_pkg` module."""

from __future__ import annotations

import base64

import docx
from docx.opc.non_phys_pkg import StrPkgReader

# -- a 1x1 transparent PNG --
_PNG_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)

_FLAT_OPC_DOCUMENT = (
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
    '<pkg:part pkg:name="/_rels/.rels" '
    'pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">'
    "<pkg:xmlData>"
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument'
    '/2006/relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
    "</pkg:xmlData>"
    "</pkg:part>"
    '<pkg:part pkg:name="/word/document.xml" '
    'pkg:contentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml">'
    "<pkg:xmlData>"
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:body><w:p><w:r><w:t>hello from flat opc</w:t></w:r></w:p></w:body>"
    "</w:document>"
    "</pkg:xmlData>"
    "</pkg:part>"
    "</pkg:package>"
)


class DescribeDocumentFromFlatOpc:
    """Integration-level test: loading a real `docx.Document` from a flat-OPC string,
    exercising the full `Package.open(..., is_from_file=False)` chain end-to-end,
    unmocked."""

    def it_loads_a_document_from_a_flat_opc_xml_string(self):
        document = docx.Document(word_open_xml=_FLAT_OPC_DOCUMENT)

        assert [p.text for p in document.paragraphs] == ["hello from flat opc"]


class DescribeStrPkgReader:
    """Unit-test suite for `docx.opc.non_phys_pkg.StrPkgReader`."""

    def it_extracts_xml_part_blobs(self):
        xml_str = (
            '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
            '<pkg:part pkg:name="/word/document.xml" '
            'pkg:contentType="application/vnd.openxmlformats-officedocument'
            '.wordprocessingml.document.main+xml">'
            "<pkg:xmlData><w:document/></pkg:xmlData>"
            "</pkg:part>"
            "</pkg:package>"
        )
        reader = StrPkgReader(xml_str)

        blob = reader.blob_for("/word/document.xml")

        assert blob == "<w:document/>"

    def it_extracts_binary_part_blobs_decoded_from_base64(self):
        xml_str = (
            '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
            '<pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png">'
            "<pkg:binaryData>" + _PNG_B64 + "</pkg:binaryData>"
            "</pkg:part>"
            "</pkg:package>"
        )
        reader = StrPkgReader(xml_str)

        blob = reader.blob_for("/word/media/image1.png")

        assert blob == base64.b64decode(_PNG_B64)

    def it_returns_None_for_a_part_not_present(self):
        reader = StrPkgReader("<pkg:package/>")
        assert reader.blob_for("/word/document.xml") is None

    def it_generates_content_types_xml_from_the_package_parts(self):
        xml_str = (
            '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
            '<pkg:part pkg:name="/word/document.xml" '
            'pkg:contentType="application/vnd.openxmlformats-officedocument'
            '.wordprocessingml.document.main+xml">'
            "<pkg:xmlData><w:document xmlns:w="
            '"http://schemas.openxmlformats.org/wordprocessingml/2006/main"/></pkg:xmlData>'
            "</pkg:part>"
            '<pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png">'
            "<pkg:binaryData>" + _PNG_B64 + "</pkg:binaryData>"
            "</pkg:part>"
            "</pkg:package>"
        )
        reader = StrPkgReader(xml_str)

        content_types_xml = reader.content_types_xml.decode("utf-8")

        assert "/word/document.xml" in content_types_xml
        assert "png" in content_types_xml

    def it_is_a_no_op_to_close(self):
        StrPkgReader("<pkg:package/>").close()
