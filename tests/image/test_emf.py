"""Test suite for docx.image.emf module."""

import io
import struct

import pytest

from docx.image.constants import MIME_TYPE
from docx.image.emf import Emf

from ..unitutil.mock import ANY, initializer_mock


class DescribeEmf:
    def it_can_construct_from_an_emf_stream(self, Emf__init__):
        # -- frame is (left=0, top=0, right=1000, bottom=500), in .01mm units --
        px_width, px_height, dpi = 118, 59, 300
        header = (
            b"filler\x00\x00"  # iType, nSize (unused)
            + b"boundsboundsbnds"  # rclBounds, 16 bytes (unused)
            + struct.pack("<4i", 0, 0, 1000, 500)  # rclFrame
            + b" EMF"  # dSignature
        )
        stream = io.BytesIO(header)

        emf = Emf.from_stream(stream)

        Emf__init__.assert_called_once_with(ANY, px_width, px_height, dpi, dpi)
        assert isinstance(emf, Emf)

    def it_knows_its_content_type(self):
        emf = Emf(None, None, None, None)
        assert emf.content_type == MIME_TYPE.EMF

    def it_knows_its_default_ext(self):
        emf = Emf(None, None, None, None)
        assert emf.default_ext == "emf"

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def Emf__init__(self, request):
        return initializer_mock(request, Emf)
