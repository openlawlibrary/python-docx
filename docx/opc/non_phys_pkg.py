# encoding: utf-8

"""
Provides a general interface to a *non-physical* OPC package, such as a zip file.
"""

from __future__ import absolute_import

from .exceptions import PackageNotFoundError
from .packuri import CONTENT_TYPES_URI, PackURI
from .part import XmlPart
from .pkgwriter import _ContentTypesItem


class StrPkgReader(object):
    """
    Implements |PkgReader| interface for an OPC package contained in a string.
    """
    def __init__(self, xml_str):
        super().__init__()
        self._xml_str = xml_str

    def blob_for(self, pack_uri):
        """
        Return contents of "file" corresponding to *pack_uri* in package
        string.
        """
        blob = None

        pkg_part_start_tag = '<pkg:part pkg:name="' + pack_uri
        pkg_xml_start_tag = '<pkg:xmlData>'
        pkg_xml_end_tag = '</pkg:xmlData>'

        pkg_start_pos = self._xml_str.find(pkg_part_start_tag)
        if pkg_start_pos >= 0:
            pkg_xml_start_pos = self._xml_str.find(
                pkg_xml_start_tag, pkg_start_pos
            )
            if pkg_xml_start_pos >= 0:
                pkg_xml_start_pos += len(pkg_xml_start_tag)
                pkg_end_pos = self._xml_str.find(
                    pkg_xml_end_tag, pkg_xml_start_pos
                )
                if pkg_end_pos >= 0:
                    return self._xml_str[pkg_xml_start_pos:pkg_end_pos]

        return None

    def close(self):
        """
        Provides interface consistency with |ZipFileSystem|, but does
        nothing, since a string doesn't need closing.
        """
        pass

    @property
    def content_types_xml(self):
        """
        Return the `[Content_Types].xml` blob from the package.
        """
        parts = self._get_parts()
        cti = _ContentTypesItem.from_parts(parts)
        return cti.blob

    def _get_parts(self):
        parts = []
        pkg_meta = self._get_pkg_meta()
        for part_name, content_type in pkg_meta:
            blob = self.blob_for(part_name)
            part = XmlPart.load(PackURI(part_name), content_type, blob, None)
            parts.append(part)
        return parts

    def _get_pkg_meta(self):
        pkg_meta = []
        i = 0
        while i >= 0:
            pkg_name, i = self._harvest_substring('pkg:name="', '"', i)
            if i >= 0:
                pkg_content_type, i = self._harvest_substring(
                    'pkg:contentType="', '"', i
                )
                if i >= 0:
                    pkg_meta.append((pkg_name, pkg_content_type))
        return pkg_meta

    def _harvest_substring(self, str_start, str_end, i):
        value = None
        str_start_pos = self._xml_str.find(str_start, i)
        if str_start_pos >= 0:
            str_start_pos += len(str_start)
            str_end_pos = self._xml_str.find(str_end, str_start_pos)
            if str_end_pos >= 0:
                value = self._xml_str[str_start_pos:str_end_pos]
                i = str_end_pos
        else:
            i = -1
        return value, i

    def rels_xml_for(self, source_uri):
        """
        Return rels item XML for source with *source_uri*, or None if the
        item has no rels item.
        """
        rels_xml = self.blob_for(source_uri.rels_uri)

        return rels_xml
