# encoding: utf-8

"""
Custom element classes related to text symbols (CT_Sym).
"""

from ..xmlchemy import (
    BaseOxmlElement, OptionalAttribute
)
from ..simpletypes import (
    ST_String
)
from warnings import warn

# for a specific font map most used symbols to appropriate character
font_to_char_map = {
    'WP TypographicSymbols': {
        # font code symbol mapping
        '0027': "§",
        '0040': "”",
        '0038': "©",
        '003D': "’",
        '0041': "“",
        '0042': "–",
        '0043': "—",
        # aspose ascii char mapping
        '\'': "§",
        '@': "”",
        'C': "©",
        'A': "“",
        '-': "–",
    },
}

def convert_char_to_symbol(char, font):
    """
    Returns the proper symbol based on the character/character code and the font style.
    Throws an KeyError if the desired symbol is not mapped.
    """
    return font_to_char_map[font][char]

class CT_Sym(BaseOxmlElement):
    """
    ``<w:sym>`` element, containing the symbol.
    """

    # We are using type `ST_String` because we have that implemented, the official type should be `ST_ShortHexNumber`, maybe in the future add support for that type
    char = OptionalAttribute('w:char', ST_String)
    font = OptionalAttribute('w:font', ST_String)

    @property
    def readSymbol(self):
        """
        Returns the symbol value if the font is supported otherwise returns the Unicode value with a warning.
        """
        try:
            symbol = convert_char_to_symbol(self.char.upper(), self.font)
        except KeyError:
            msg = f"Symbol <{self.char}> is not supported in font '{self.font}', fallback to Uniode value"
            warn(msg, UserWarning, stacklevel=2)
            symbol = r"\u"+self.char

        return symbol

