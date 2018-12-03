# encoding: utf-8
from __future__ import absolute_import, unicode_literals
from .simpletypes import XsdString, XsdUnsignedInt
from .xmlchemy import BaseOxmlElement, RequiredAttribute


class CT_BookmarkStart(BaseOxmlElement):
    """
    Used for ``<w:bookmarkStart>`` element. Specifies the id and name of a
    Bookmark start.
    """
    id = RequiredAttribute('w:id', XsdUnsignedInt)
    name = RequiredAttribute('w:name', XsdString)


class CT_BookmarkEnd(BaseOxmlElement):
    """
    Used for ``<w:bookmarkEnd>`` element. Specifies the id and name of a
    Bookmark end.
    """
    id = RequiredAttribute('w:id', XsdUnsignedInt)
