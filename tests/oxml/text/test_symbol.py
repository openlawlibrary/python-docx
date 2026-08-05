"""Test suite for the docx.oxml.text.symbol module."""

from typing import cast

import pytest

from docx.oxml.text.symbol import CT_Sym

from ...unitutil.cxml import element


class DescribeCT_Sym:
    """Unit-test suite for the CT_Sym (symbol, <w:sym>) element."""

    def it_provides_the_Unicode_character_for_a_mapped_symbol(self):
        sym = self._sym("0038", "WP TypographicSymbols")
        assert str(sym) == "©"

    def it_matches_the_char_case_insensitively(self):
        sym = self._sym("0038".lower(), "WP TypographicSymbols")
        assert str(sym) == "©"

    def it_falls_back_to_an_escaped_char_code_for_an_unmapped_char(self):
        sym = self._sym("0013", "WP TypographicSymbols")
        with pytest.warns(UserWarning, match="Symbol <0013> is not supported"):
            actual = str(sym)
        assert actual == "\\u0013"

    def it_falls_back_to_an_escaped_char_code_for_an_unmapped_font(self):
        sym = self._sym("0027", "Calibri")
        with pytest.warns(UserWarning, match="Symbol <0027> is not supported"):
            actual = str(sym)
        assert actual == "\\u0027"

    # fixture --------------------------------------------------------

    def _sym(self, char: str, font: str) -> CT_Sym:
        sym = cast(CT_Sym, element("w:sym"))
        sym.char = char
        sym.font = font
        return sym
