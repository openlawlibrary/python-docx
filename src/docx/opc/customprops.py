"""Provides an object model for custom document properties."""

from __future__ import annotations

import numbers
from typing import TYPE_CHECKING, Any

from lxml import etree

from docx.oxml.ns import nsmap, qn

if TYPE_CHECKING:
    from docx.oxml.customprops import CT_CustomProperties


class CustomProperties:
    """Corresponds to part named `/docProps/custom.xml`.

    Provides read/write access to the custom document properties for a document, via a
    `dict`-like interface keyed by property name.
    """

    def __init__(self, element: CT_CustomProperties):
        self._element = element

    def __getitem__(self, key: str) -> str | int | bool | None:
        """Value of the custom property named `key`, or |None| if not found."""
        prop = self._lookup(key)
        if prop is None:
            return None
        elm = prop[0]
        if elm.tag == qn("vt:i4"):
            return 0 if elm.text is None else int(elm.text)
        if elm.tag == qn("vt:bool"):
            return elm.text == "1"
        return elm.text

    def __setitem__(self, key: str, value: Any) -> None:
        """Add or update the custom property named `key` to `value`."""
        prop = self._lookup(key)
        if prop is None:
            elm_tag, text = self._elm_tag_and_text_for(value)
            prop = etree.SubElement(self._element, "property")
            elm = etree.SubElement(prop, elm_tag, nsmap={"vt": nsmap["vt"]})
            elm.text = text
            prop.set("name", key)
            prop.set("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            prop.set("pid", str(len(self._element) + 1))
        else:
            elm = prop[0]
            _, elm.text = self._elm_tag_and_text_for(value)

    def __len__(self) -> int:
        return len(self._element)

    def _lookup(self, key: str):
        """The `<property>` child element named `key`, or |None| if not present."""
        for child in self._element:
            if child.get("name") == key:
                return child
        return None

    @staticmethod
    def _elm_tag_and_text_for(value: Any) -> tuple[str, str]:
        """The (tag, text) pair for the `vt:*` element representing `value`."""
        if isinstance(value, bool):
            return qn("vt:bool"), str(int(value))
        if isinstance(value, numbers.Number):
            return qn("vt:i4"), str(int(value))  # pyright: ignore[reportArgumentType]
        return qn("vt:lpwstr"), str(value)
