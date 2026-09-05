"""Каталог композиционных паттернов для учебных и IT-презентаций.

Каталог собран как 24 композиционных архетипа в пяти семантических семействах.
Итого 120 устойчивых вариантов. Паттерн описывает геометрию и декоративный
chrome, а содержание по-прежнему выбирается анализатором исходного слайда.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class LayoutPattern:
    identifier: str
    family: str
    composition: str
    alignment: str
    dark: bool
    density: str
    image_side: str
    accent_mode: str


COMPOSITIONS = (
    ("hero", "center", "low", "none"),
    ("split", "left", "medium", "right"),
    ("split_reverse", "right", "medium", "left"),
    ("stack", "left", "high", "none"),
    ("grid_2", "left", "medium", "both"),
    ("grid_3", "center", "medium", "both"),
    ("grid_4", "left", "high", "both"),
    ("timeline", "center", "medium", "none"),
    ("flow_vertical", "center", "high", "none"),
    ("flow_horizontal", "center", "medium", "none"),
    ("quote", "center", "low", "none"),
    ("metrics", "center", "low", "both"),
    ("table_focus", "left", "high", "none"),
    ("image_left", "right", "medium", "left"),
    ("image_right", "left", "medium", "right"),
    ("image_mosaic", "left", "medium", "both"),
    ("code_left", "right", "high", "left"),
    ("code_right", "left", "high", "right"),
    ("cards", "left", "high", "none"),
    ("diagonal", "left", "medium", "right"),
    ("sidebar", "left", "medium", "left"),
    ("magazine", "left", "medium", "both"),
    ("notebook", "left", "high", "none"),
    ("terminal", "left", "high", "none"),
)

FAMILIES = {
    "A": "code",
    "B": "interactive",
    "C": "definition",
    "D": "timeline",
    "E": "minimal",
}


LAYOUT_CATALOG: dict[str, LayoutPattern] = {}
for family, family_name in FAMILIES.items():
    for variation, (composition, alignment, density, image_side) in enumerate(COMPOSITIONS, 1):
        identifier = f"{family}-{variation:02d}"
        LAYOUT_CATALOG[identifier] = LayoutPattern(
            identifier=identifier,
            family=family_name,
            composition=composition,
            alignment=alignment,
            dark=family == "A" and variation % 2 == 1,
            density=density,
            image_side=image_side,
            accent_mode="solid" if variation % 3 else "split",
        )


def get_layout_pattern(identifier: str | None) -> LayoutPattern:
    """Вернуть безопасный паттерн с fallback на нейтральный вариант."""

    return LAYOUT_CATALOG.get(identifier or "E-01", LAYOUT_CATALOG["E-01"])


def catalog_stats() -> dict[str, int]:
    return {
        "patterns": len(LAYOUT_CATALOG),
        "compositions": len(COMPOSITIONS),
        "families": len(FAMILIES),
    }