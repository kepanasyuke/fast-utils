"""Shared design tokens used by slide templates."""

from pptx.dml.color import RGBColor

FONT_PAIRS = {
    "modern": ("Segoe UI", "Verdana"),
    "classic": ("Times New Roman", "Georgia"),
    "code": ("Consolas", "Courier New"),
}
FONTS = {
    "header": "Segoe UI",
    "body": "Verdana",
    "code": "Consolas",
}
COLORS = {
    "white": RGBColor(255, 255, 255),
    "text": RGBColor(21, 30, 47),
    "muted": RGBColor(85, 101, 126),
    "light": RGBColor(248, 250, 252),
    "border": RGBColor(226, 232, 240),
    "terminal": RGBColor(13, 19, 33),
}
SPACING = {"margin": 0.85, "gutter": 0.35, "padding": 0.15}
TYPOGRAPHY = {"title_size": 30, "body_size": 18, "small_size": 12, "line_height": 1.2}