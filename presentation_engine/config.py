"""Runtime configuration for the presentation redesign engine."""

from __future__ import annotations

import os
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_DIR / "generated_presentations"
SLIDE_RATIO = os.getenv("SLIDE_RATIO", "16:9")
RATIO_DIMENSIONS = {
    "16:9": (13.333, 7.5),
    "4:3": (10.0, 7.5),
    "16:10": (12.8, 8.0),
}
IMAGE_QUALITY = 88
IMAGE_MAX_WIDTH = 1920
IMAGE_MAX_HEIGHT = 1080
MAX_UPLOAD_BYTES = 80 * 1024 * 1024
FALLBACK_FONT = "Arial"


def slide_dimensions(ratio: str | None = None) -> tuple[float, float]:
    """Return a supported canvas size, falling back to the default ratio."""

    return RATIO_DIMENSIONS.get(ratio or SLIDE_RATIO, RATIO_DIMENSIONS["16:9"])