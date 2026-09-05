"""Small, tested utilities shared by the PPTX pipeline."""

from __future__ import annotations

import io
import logging
from pathlib import Path

import qrcode
from PIL import Image, ImageEnhance, ImageOps

logger = logging.getLogger(__name__)


def enhance_image(image: Image.Image, contrast: float = 1.08, brightness: float = 1.02) -> Image.Image:
    """Improve raster assets conservatively without changing their dimensions."""

    image = ImageEnhance.Contrast(image).enhance(contrast)
    return ImageEnhance.Brightness(image).enhance(brightness)


def fit_image(blob: bytes, width: int, height: int, enhance: bool = True) -> io.BytesIO:
    """Fit a raster image inside a canvas without cropping important content."""

    image = Image.open(io.BytesIO(blob)).convert("RGBA")
    if enhance:
        image = enhance_image(image)
    fitted = ImageOps.contain(image, (width, height), method=Image.Resampling.LANCZOS)
    canvas = Image.new("RGBA", (width, height), (255, 255, 255, 255))
    canvas.alpha_composite(fitted, ((width - fitted.width) // 2, (height - fitted.height) // 2))
    output = io.BytesIO()
    canvas.save(output, format="PNG", optimize=True)
    output.seek(0)
    return output


def safe_fit_image(blob: bytes, width: int, height: int) -> io.BytesIO:
    """Use raster fitting when possible and preserve unsupported vector bytes otherwise."""

    try:
        return fit_image(blob, width, height)
    except (OSError, ValueError):
        return io.BytesIO(blob)


def create_qr(link: str) -> io.BytesIO:
    qr = qrcode.QRCode(version=1, box_size=8, border=2)
    qr.add_data(link or "https://education.yandex.ru/handbook")
    qr.make(fit=True)
    image = qr.make_image(fill_color="black", back_color="white")
    output = io.BytesIO()
    image.save(output, format="PNG")
    output.seek(0)
    return output


def contrast_color(rgb: tuple[int, int, int]) -> tuple[int, int, int]:
    luminance = (rgb[0] * 299 + rgb[1] * 587 + rgb[2] * 114) / 1000
    return (255, 255, 255) if luminance < 145 else (21, 30, 47)


def validate_image_path(path: str | None) -> bool:
    return bool(path and Path(path).is_file())