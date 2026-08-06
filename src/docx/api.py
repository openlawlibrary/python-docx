"""Directly exposed API functions and classes, :func:`Document` for now.

Provides a syntactically more convenient API for interacting with the OpcPackage graph.
"""

from __future__ import annotations

import os
from typing import IO, TYPE_CHECKING, cast

from docx.opc.constants import CONTENT_TYPE as CT
from docx.package import Package

if TYPE_CHECKING:
    from docx.document import Document as DocumentObject
    from docx.parts.document import DocumentPart


def Document(
    docx: str | IO[bytes] | None = None, word_open_xml: str | None = None
) -> DocumentObject:
    """Return a |Document| object loaded from `docx`, where `docx` can be either a path
    to a ``.docx`` file (a string) or a file-like object.

    If `docx` is missing or ``None`` and `word_open_xml` is also ``None``, the built-in
    default document "template" is loaded. `word_open_xml` may instead be specified as a
    "flat OPC" XML string -- `docx` and `word_open_xml` are mutually exclusive.
    """
    if docx and word_open_xml:
        raise ValueError("Must either specify docx or word_open_xml, but not both")

    if word_open_xml is not None:
        document_part = cast(
            "DocumentPart", Package.open(word_open_xml, is_from_file=False).main_document_part
        )
        if document_part.content_type != CT.WML_DOCUMENT_MAIN:
            tmpl = "string '%s' is not a Word document, content type is '%s'"
            raise ValueError(tmpl % (word_open_xml, document_part.content_type))
        return document_part.document

    docx = _default_docx_path() if docx is None else docx
    document_part = cast("DocumentPart", Package.open(docx).main_document_part)
    if document_part.content_type != CT.WML_DOCUMENT_MAIN:
        tmpl = "file '%s' is not a Word file, content type is '%s'"
        raise ValueError(tmpl % (docx, document_part.content_type))
    return document_part.document


def _default_docx_path():
    """Return the path to the built-in default .docx package."""
    _thisdir = os.path.split(__file__)[0]
    return os.path.join(_thisdir, "templates", "default.docx")
