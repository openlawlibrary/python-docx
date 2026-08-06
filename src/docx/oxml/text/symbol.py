"""Custom element classes related to text symbols (CT_Sym)."""

from __future__ import annotations

import warnings

from docx.oxml.simpletypes import ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute

# -- maps the most-used symbols to their Unicode character, per font. The `w:char`
# -- value for these fonts is a hex character code rather than a Unicode code point,
# -- so it can't be decoded generically.
font_to_char_map: dict[str, dict[str, str]] = {
    "WP TypographicSymbols": {
        "0021": "°",
        "0022": "°",
        "0024": "•",
        "0026": "¶",
        "0027": "§",
        "0028": "¡",
        "0029": "¿",
        "0032": "½",
        "0033": "¼",
        "0034": "¢",
        "0038": "©",
        "003A": "¾",
        "003D": "’",
        "0040": "”",
        "0041": "“",
        "0042": "–",
        "0043": "—",
        "0049": "‡",
        "004A": "™",
        "0053": "—",
        "0059": "…",
        "005A": "$",
        "0061": "⅓",
        "0062": "⅔",
        "0063": "⅛",
        "0064": "⅜",
        "0065": "⅝",
        "0066": "⅞",
        "006E": "—",
    },
    "WP IconicSymbolsA": {
        "F046": "☎",
    },
    "Symbol": {
        "F0B0": "°",
    },
    "WP Phonetic": {
        "F05F": "C",
    },
}


class CT_Sym(BaseOxmlElement):
    """`<w:sym>` element, a symbol from a font's private-use glyph range."""

    char: str | None = OptionalAttribute("w:char", ST_String)  # pyright: ignore[reportAssignmentType]
    font: str | None = OptionalAttribute("w:font", ST_String)  # pyright: ignore[reportAssignmentType]

    def __str__(self) -> str:
        """The Unicode character this symbol represents, for text-equivalent purposes.

        Falls back to the raw `\\uXXXX`-escaped character code, with a warning, when
        `font` is not one of the fonts this module knows how to map.
        """
        char = "" if self.char is None else self.char.upper()
        try:
            return font_to_char_map[self.font][char]  # pyright: ignore[reportArgumentType]
        except KeyError:
            warnings.warn(
                f"Symbol <{self.char}> is not supported in font '{self.font}',"
                f" fallback to Uniode value",
                UserWarning,
                stacklevel=2,
            )
            return f"\\u{self.char}"
