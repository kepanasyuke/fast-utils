"""Reusable layout primitives for educational PowerPoint redesigns."""

from __future__ import annotations

from dataclasses import dataclass, field
from math import ceil
from typing import Iterable

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


@dataclass(frozen=True)
class ThemeTokens:
    """A small design system shared by every slide template."""

    accent: RGBColor
    text: RGBColor = RGBColor(21, 30, 47)
    muted: RGBColor = RGBColor(85, 101, 126)
    white: RGBColor = RGBColor(255, 255, 255)
    surface: RGBColor = RGBColor(248, 250, 252)
    border: RGBColor = RGBColor(226, 232, 240)
    dark_surface: RGBColor = RGBColor(13, 19, 33)
    header_font: str = "Segoe UI"
    body_font: str = "Verdana"
    code_font: str = "Consolas"


@dataclass(frozen=True)
class Rect:
    """A slide rectangle expressed in inches."""

    x: float
    y: float
    width: float
    height: float

    @property
    def right(self) -> float:
        return self.x + self.width

    @property
    def bottom(self) -> float:
        return self.y + self.height

    def inset(self, amount: float) -> "Rect":
        return Rect(
            self.x + amount,
            self.y + amount,
            max(0, self.width - amount * 2),
            max(0, self.height - amount * 2),
        )


@dataclass
class TextBlock:
    """Normalized text ready to be placed into a template zone."""

    text: str
    role: str = "body"
    level: int = 0
    emphasis: bool = False


@dataclass
class AssetRef:
    """Metadata for an image or vector asset from the source deck."""

    blob: bytes
    mime_type: str = "image/png"
    source_index: int = 0
    alt_text: str = ""


@dataclass
class SlideContent:
    """Intermediate representation shared by classifiers and templates."""

    index: int
    title: str
    blocks: list[TextBlock] = field(default_factory=list)
    tables: list[list[list[str]]] = field(default_factory=list)
    images: list[AssetRef] = field(default_factory=list)
    kind: str = "content_overview"
    source_width: float = 13.333
    source_height: float = 7.5

    @property
    def body(self) -> str:
        return "\n".join(block.text for block in self.blocks)

    @property
    def plain_text(self) -> str:
        return " ".join([self.title, self.body]).strip()

    @property
    def has_table(self) -> bool:
        return bool(self.tables)

    @property
    def has_images(self) -> bool:
        return bool(self.images)


@dataclass(frozen=True)
class LayoutIssue:
    """A measurable issue found after a slide has been rendered."""

    message: str
    severity: str = "warning"
    rect: Rect | None = None


def estimate_line_capacity(zone: Rect, font_size: float, line_height: float = 1.2) -> int:
    """Estimate visible lines in a text zone without depending on PowerPoint."""

    if zone.height <= 0 or font_size <= 0:
        return 0
    return max(1, int((zone.height * 72) / (font_size * line_height)))


def estimate_characters_per_line(zone: Rect, font_size: float, average_width: float = 0.52) -> int:
    """Estimate characters per line for Latin and Cyrillic body text."""

    if zone.width <= 0 or font_size <= 0:
        return 1
    points = zone.width * 72
    return max(8, int(points / (font_size * average_width)))


def estimate_text_lines(text: str, zone: Rect, font_size: float) -> int:
    """Estimate wrapped line count, respecting explicit paragraph breaks."""

    per_line = estimate_characters_per_line(zone, font_size)
    return sum(max(1, ceil(len(part) / per_line)) for part in text.splitlines() or [""])


def estimate_text_height(text: str, zone: Rect, font_size: float, line_height: float = 1.2) -> float:
    """Return estimated text height in inches."""

    lines = estimate_text_lines(text, zone, font_size)
    return lines * font_size * line_height / 72


def choose_font_size(
    text: str,
    zone: Rect,
    preferred: float,
    minimum: float = 10,
    line_height: float = 1.2,
) -> float:
    """Choose the largest size that fits a text zone within a bounded range."""

    size = preferred
    while size > minimum and estimate_text_height(text, zone, size, line_height) > zone.height:
        size -= 1
    return max(minimum, size)


def split_text_to_fit(
    blocks: Iterable[TextBlock],
    zone: Rect,
    font_size: float,
    max_blocks: int | None = None,
) -> list[list[TextBlock]]:
    """Paginate blocks by estimated height instead of arbitrary character counts."""

    pages: list[list[TextBlock]] = []
    current: list[TextBlock] = []
    current_height = 0.0
    line_height = font_size * 1.2 / 72
    for block in blocks:
        block_height = estimate_text_height(block.text, zone, font_size) + line_height
        reaches_limit = current and current_height + block_height > zone.height
        reaches_count = max_blocks is not None and len(current) >= max_blocks
        if reaches_limit or reaches_count:
            pages.append(current)
            current = []
            current_height = 0.0
        current.append(block)
        current_height += block_height
    if current or not pages:
        pages.append(current)
    return pages


def keep_inside(rect: Rect, canvas: Rect) -> Rect:
    """Clamp a rectangle to the slide canvas while preserving its origin."""

    width = min(rect.width, canvas.width)
    height = min(rect.height, canvas.height)
    x = min(max(canvas.x, rect.x), canvas.right - width)
    y = min(max(canvas.y, rect.y), canvas.bottom - height)
    return Rect(x, y, width, height)


def validate_rect(rect: Rect, canvas: Rect) -> list[LayoutIssue]:
    """Validate that a generated object stays inside the slide."""

    issues = []
    if rect.width <= 0 or rect.height <= 0:
        issues.append(LayoutIssue("Размер объекта должен быть положительным", "error", rect))
    if rect.x < canvas.x or rect.y < canvas.y or rect.right > canvas.right or rect.bottom > canvas.bottom:
        issues.append(LayoutIssue("Объект выходит за границы слайда", "error", rect))
    return issues


def align_text(paragraph, alignment: PP_ALIGN = PP_ALIGN.LEFT) -> None:
    """Apply a predictable alignment to a python-pptx paragraph."""

    paragraph.alignment = alignment


def to_pptx_rect(rect: Rect) -> tuple[int, int, int, int]:
    """Convert an inch rectangle to python-pptx units."""

    return Inches(rect.x), Inches(rect.y), Inches(rect.width), Inches(rect.height)


def point_in_canvas(x: float, y: float, canvas: Rect) -> bool:
    return canvas.x <= x <= canvas.right and canvas.y <= y <= canvas.bottom


def distribute_evenly(count: int, start: float, end: float) -> list[float]:
    """Return evenly spaced coordinates for nodes or cards."""

    if count <= 0:
        return []
    if count == 1:
        return [(start + end) / 2]
    step = (end - start) / (count - 1)
    return [start + step * index for index in range(count)]