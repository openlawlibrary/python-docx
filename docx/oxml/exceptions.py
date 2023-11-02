# encoding: utf-8

"""
Exceptions for oxml sub-package
"""


class XmlchemyError(Exception):
    """Generic error class."""


class InvalidXmlError(XmlchemyError):
    """
    Raised when invalid XML is encountered, such as on attempt to access a
    missing required child element
    """

class HiddenTextTC(Exception):
    """
    Raised when parsing hidden text that represents table content.
    This text is used as a link to Table of Contents, and shouldn't be return as viewable text.
    """
