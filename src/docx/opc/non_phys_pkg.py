"""Provides a general interface to a *non-physical* OPC package.

A non-physical package is one contained in a string rather than a zip file, using the
"flat OPC" format -- a single XML document with `<pkg:package>`/`<pkg:part>` wrapper
elements around each part, part content stored as either `<pkg:xmlData>` (text parts) or
base64-encoded `<pkg:binaryData>` (binary parts such as images).
"""

from __future__ import annotations

import base64

from docx.image.image import Image
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.opc.pkgwriter import _ContentTypesItem  # pyright: ignore[reportPrivateUsage]
from docx.parts.image import ImagePart


class StrPkgReader:
    """Implements the `PhysPkgReader` interface for an OPC package contained in a
    "flat OPC" XML string."""

    def __init__(self, xml_str: str):
        self._xml_str = xml_str

    def blob_for(self, pack_uri: str) -> bytes | str | None:
        """Contents of the "file" corresponding to `pack_uri` in the package string."""
        is_image_part = "/word/media/image" in pack_uri
        pkg_part_start_tag = '<pkg:part pkg:name="' + pack_uri
        pkg_xml_start_tag = "<pkg:binaryData>" if is_image_part else "<pkg:xmlData>"
        pkg_xml_end_tag = "</" + pkg_xml_start_tag[1:]

        pkg_start_pos = self._xml_str.find(pkg_part_start_tag)
        if pkg_start_pos < 0:
            return None

        pkg_xml_start_pos = self._xml_str.find(pkg_xml_start_tag[:-1], pkg_start_pos)
        if pkg_xml_start_pos < 0:
            return None
        pkg_xml_start_pos += len(pkg_xml_start_tag) - 1
        if self._xml_str[pkg_xml_start_pos] == ">":
            pkg_xml_start_pos += 1
        else:
            pkg_xml_start_pos = self._xml_str.find(">", pkg_xml_start_pos) + 1

        pkg_end_pos = self._xml_str.find(pkg_xml_end_tag, pkg_xml_start_pos)
        if pkg_end_pos < 0:
            return None

        blob = self._xml_str[pkg_xml_start_pos:pkg_end_pos]
        return base64.b64decode(blob) if is_image_part else blob

    def close(self) -> None:
        """Provides interface consistency with `PhysPkgReader`, but does nothing,
        since a string doesn't need closing."""
        pass

    @property
    def content_types_xml(self) -> bytes:
        """The `[Content_Types].xml` blob for the package."""
        parts = self._get_parts()
        return _ContentTypesItem.from_parts(parts).blob

    def rels_xml_for(self, source_uri: PackURI) -> bytes | str | None:
        """Rels item XML for the source identified by `source_uri`.

        |None| if the source has no rels item.
        """
        return self.blob_for(source_uri.rels_uri)

    def _get_parts(self) -> list[XmlPart | ImagePart]:
        """A stub part -- carrying only `partname` and `content_type` -- for each part
        described in the package's metadata, used only to compute content types."""
        parts: list[XmlPart | ImagePart] = []
        for part_name, content_type in self._get_pkg_meta():
            blob = self.blob_for(part_name)
            if "image" in content_type:
                image = Image.from_blob(blob)  # pyright: ignore[reportArgumentType]
                parts.append(ImagePart.from_image(image, PackURI(part_name)))
            else:
                parts.append(
                    XmlPart.load(
                        PackURI(part_name),
                        content_type,
                        blob,  # pyright: ignore[reportArgumentType]
                        None,  # pyright: ignore[reportArgumentType]
                    )
                )
        return parts

    def _get_pkg_meta(self) -> list[tuple[str, str]]:
        """List of (partname, content-type) pairs, one for each `<pkg:part>` element
        described in the package string."""
        pkg_meta: list[tuple[str, str]] = []
        i = 0
        while i >= 0:
            pkg_name, i = self._harvest_substring('pkg:name="', '"', i)
            if i < 0 or pkg_name is None:
                break
            pkg_content_type, i = self._harvest_substring('pkg:contentType="', '"', i)
            if i < 0 or pkg_content_type is None:
                break
            pkg_meta.append((pkg_name, pkg_content_type))
        return pkg_meta

    def _harvest_substring(self, str_start: str, str_end: str, i: int) -> tuple[str | None, int]:
        """The substring between `str_start` and `str_end`, searching from offset `i`,
        and the offset immediately following it (or (None, -1) if not found)."""
        str_start_pos = self._xml_str.find(str_start, i)
        if str_start_pos < 0:
            return None, -1
        str_start_pos += len(str_start)
        str_end_pos = self._xml_str.find(str_end, str_start_pos)
        if str_end_pos < 0:
            return None, -1
        return self._xml_str[str_start_pos:str_end_pos], str_end_pos
