"""Provides Emf image header parser."""

from __future__ import annotations

import struct
from typing import IO

from docx.image.constants import MIME_TYPE
from docx.image.image import BaseImageHeader


class Emf(BaseImageHeader):
    """Image header parser for EMF (Enhanced Metafile) images.

    EMF is a vector format, so it has no inherent pixel dimensions; a fixed 300 dpi is
    assumed to convert its physical frame size to a pixel count.
    """

    @classmethod
    def from_stream(cls, stream: IO[bytes]) -> Emf:
        """Return a |Emf| instance having header properties parsed from the EMF image
        in `stream`."""
        # -- the EMR_HEADER structure this parses, in relevant part:
        # --   @0  DWORD  iType
        # --   @4  DWORD  nSize
        # --   @8  RECTL  rclBounds       (device units, unused here)
        # --   @24 RECTL  rclFrame        (.01 mm units: left, top, right, bottom)
        # --   @40 DWORD  dSignature      (already matched by the caller)
        stream.seek(24)
        left, top, right, bottom = struct.unpack("<4i", stream.read(16))

        dpi = 300
        mm_width = (right - left) / 100.0
        mm_height = (bottom - top) / 100.0
        # -- 1 inch == 25.4 mm --
        px_width = int(mm_width * dpi / 25.4)
        px_height = int(mm_height * dpi / 25.4)

        return cls(px_width, px_height, dpi, dpi)

    @property
    def content_type(self) -> str:
        """MIME content type for this image, unconditionally `image/x-emf` for EMF
        images."""
        return MIME_TYPE.EMF

    @property
    def default_ext(self) -> str:
        """Default filename extension, always 'emf' for EMF images."""
        return "emf"
