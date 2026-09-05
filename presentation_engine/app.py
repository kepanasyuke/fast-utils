"""Web app that turns an uploaded PowerPoint into a branded presentation."""

from __future__ import annotations

import io
import os
import re
import uuid
from pathlib import Path

import qrcode
import uvicorn
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


app = FastAPI(title="AI Presentation Brand Engine")
OUTPUT_DIR = Path(__file__).resolve().parent / "generated_presentations"
OUTPUT_DIR.mkdir(exist_ok=True)

SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)
FONT_HEADER = "Segoe UI"
FONT_BODY = "Verdana"
FONT_CODE = "Consolas"

WHITE = RGBColor(255, 255, 255)
TEXT = RGBColor(21, 30, 47)
MUTED = RGBColor(85, 101, 126)
LIGHT = RGBColor(248, 250, 252)
BORDER = RGBColor(226, 232, 240)
TERMINAL = RGBColor(13, 19, 33)
PALETTES = {
    "cyber_blue": RGBColor(6, 182, 212),
    "matrix_green": RGBColor(16, 185, 129),
    "js_yellow": RGBColor(234, 179, 8),
    "python_blue": RGBColor(59, 130, 246),
    "ruby_red": RGBColor(225, 29, 72),
    "git_orange": RGBColor(249, 115, 22),
    "ai_purple": RGBColor(139, 92, 246),
}


def hex_color(color: RGBColor) -> str:
    return f"#{color[0]:02x}{color[1]:02x}{color[2]:02x}"


def text_box(slide, x, y, width, height, text, font=FONT_BODY, size=18,
             color=TEXT, bold=False, align=PP_ALIGN.LEFT):
    shape = slide.shapes.add_textbox(x, y, width, height)
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = True
    paragraph = frame.paragraphs[0]
    paragraph.text = text
    paragraph.alignment = align
    paragraph.font.name = font
    paragraph.font.size = Pt(size)
    paragraph.font.color.rgb = color
    paragraph.font.bold = bold
    return shape


def card(slide, x, y, width, height, fill=WHITE, line=BORDER, radius=True):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, x, y, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = line
    shape.line.width = Pt(1.5)
    return shape


def qr_stream(link: str) -> io.BytesIO:
    image = qrcode.make(link or "https://education.yandex.ru/handbook")
    stream = io.BytesIO()
    image.save(stream, format="PNG")
    stream.seek(0)
    return stream


def parse_slides(file_bytes: bytes) -> list[dict[str, str]]:
    try:
        presentation = Presentation(io.BytesIO(file_bytes))
    except Exception as error:
        raise HTTPException(status_code=400, detail="Не удалось прочитать PPTX-файл.") from error

    result = []
    for index, slide in enumerate(presentation.slides):
        texts = [shape.text.strip() for shape in slide.shapes
                 if getattr(shape, "has_text_frame", False) and shape.text.strip()]
        if not texts:
            continue
        title = texts[0]
        body = "\n".join(texts[1:]) or title
        lower = f"{title} {body}".lower()
        if index == 0:
            slide_type = "title_root"
        elif any(token in lower for token in ("def ", "import ", "print(", "python", "код")):
            slide_type = "code_dark_ide" if index % 2 else "code_light_ide"
        elif any(token in lower for token in ("задание", "тест", "вопрос", "практика", "?")):
            slide_type = "yandex_interactive" if index % 2 else "classic_timer"
        elif any(token in lower for token in ("это", "означает", "определение", "термин")):
            slide_type = "big_definition"
        elif any(token in lower for token in ("этап", "история", "шаг", "развитие")):
            slide_type = "step_timeline"
        elif any(token in lower for token in ("git", "репозитор", "коммит", "ветка")):
            slide_type = "git_workflow"
        elif len(body) < 100:
            slide_type = "minimal_quote"
        else:
            slide_type = "split_comparison"
        result.append({"title": title, "body": body, "type": slide_type})
    return result


def add_header(slide, title: str, accent: RGBColor) -> None:
    text_box(slide, Inches(0.85), Inches(0.42), Inches(11.6), Inches(0.7),
             title, font=FONT_HEADER, size=30, bold=True)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.85), Inches(1.2), Inches(1.0), Inches(0.08))
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent
    bar.line.fill.background()


def render_slide(slide, data: dict[str, str], accent: RGBColor, link: str, index: int) -> None:
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = WHITE
    kind = data["type"]
    title, body = data["title"], data["body"]

    if kind == "title_root":
        stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.45), SLIDE_HEIGHT)
        stripe.fill.solid(); stripe.fill.fore_color.rgb = accent; stripe.line.fill.background()
        card(slide, Inches(1.4), Inches(1.75), Inches(10.5), Inches(3.9), LIGHT, BORDER)
        text_box(slide, Inches(1.8), Inches(2.45), Inches(9.7), Inches(1.1), title,
                 font=FONT_HEADER, size=42, bold=True, align=PP_ALIGN.CENTER)
        text_box(slide, Inches(2.1), Inches(3.85), Inches(9.1), Inches(1.0), body,
                 size=17, color=MUTED, align=PP_ALIGN.CENTER)
        return

    add_header(slide, title, accent)
    if kind in ("code_dark_ide", "code_light_ide"):
        fill = TERMINAL if kind == "code_dark_ide" else LIGHT
        color = WHITE if kind == "code_dark_ide" else TEXT
        card(slide, Inches(0.9), Inches(1.75), Inches(6.3), Inches(4.8), fill, accent, False)
        text_box(slide, Inches(1.25), Inches(2.1), Inches(5.6), Inches(4.1), body,
                 font=FONT_CODE, size=14, color=color)
        text_box(slide, Inches(8.0), Inches(2.5), Inches(4.0), Inches(1.2),
                 "IDE / BUILD / TEST", font=FONT_HEADER, size=22, color=accent, bold=True)
        text_box(slide, Inches(8.0), Inches(3.8), Inches(3.9), Inches(1.5),
                 "Код превращается в понятный визуальный сценарий.", size=18, color=MUTED)
    elif kind in ("yandex_interactive", "classic_timer"):
        card(slide, Inches(0.9), Inches(1.8), Inches(7.4), Inches(4.7), WHITE, accent)
        text_box(slide, Inches(1.35), Inches(2.25), Inches(6.5), Inches(2.5), body, size=19, color=MUTED)
        label = "ЗАДАНИЕ В ЯНДЕКС УЧЕБНИКЕ" if kind == "yandex_interactive" else "ВРЕМЯ ПОШЛО"
        text_box(slide, Inches(1.35), Inches(5.2), Inches(5.2), Inches(0.5), label,
                 font=FONT_HEADER, size=16, color=accent, bold=True)
        slide.shapes.add_picture(qr_stream(link), Inches(10.0), Inches(2.2), Inches(1.8), Inches(1.8))
    elif kind == "big_definition":
        card(slide, Inches(1.0), Inches(2.0), Inches(11.2), Inches(2.8), WHITE, accent)
        text_box(slide, Inches(1.5), Inches(2.5), Inches(10.2), Inches(0.5), "КЛЮЧЕВОЙ ТЕРМИН",
                 font=FONT_HEADER, size=15, color=accent, bold=True)
        text_box(slide, Inches(1.5), Inches(3.15), Inches(10.2), Inches(1.1), body,
                 font=FONT_HEADER, size=24, bold=True)
    elif kind == "step_timeline":
        steps = [part.strip() for part in re.split(r"[.!?]\s*", body) if part.strip()][:3] or [body]
        for step_index, step in enumerate(steps):
            x = Inches(1.1 + step_index * 4.0)
            node = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, Inches(3.0), Inches(0.55), Inches(0.55))
            node.fill.solid(); node.fill.fore_color.rgb = accent; node.line.fill.background()
            text_box(slide, x - Inches(0.2), Inches(2.25), Inches(2.7), Inches(0.5),
                     f"ШАГ {step_index + 1}", font=FONT_HEADER, size=14, color=accent, bold=True)
            text_box(slide, x - Inches(0.2), Inches(3.8), Inches(3.2), Inches(1.4), step, size=15, color=MUTED)
            if step_index < len(steps) - 1:
                line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + Inches(0.55), Inches(3.24), Inches(3.4), Inches(0.06))
                line.fill.solid(); line.fill.fore_color.rgb = accent; line.line.fill.background()
    elif kind == "git_workflow":
        card(slide, Inches(1.0), Inches(1.9), Inches(11.2), Inches(4.5), LIGHT, accent, False)
        text_box(slide, Inches(1.5), Inches(2.4), Inches(10.0), Inches(0.5), "GIT WORKFLOW",
                 font=FONT_HEADER, size=18, color=accent, bold=True)
        text_box(slide, Inches(1.5), Inches(3.15), Inches(10.0), Inches(2.2), body,
                 font=FONT_CODE, size=17)
    elif kind == "minimal_quote":
        text_box(slide, Inches(1.0), Inches(2.2), Inches(11.0), Inches(2.8), f'« {body} »',
                 font=FONT_HEADER, size=30, color=accent, bold=True, align=PP_ALIGN.CENTER)
    else:
        card(slide, Inches(0.9), Inches(1.8), Inches(6.0), Inches(4.7), LIGHT, BORDER)
        card(slide, Inches(7.25), Inches(1.8), Inches(5.1), Inches(4.7), WHITE, accent)
        text_box(slide, Inches(1.35), Inches(2.3), Inches(5.1), Inches(0.5), "КОНТЕКСТ", font=FONT_HEADER, size=15, color=accent, bold=True)
        text_box(slide, Inches(1.35), Inches(3.0), Inches(5.0), Inches(2.5), body, size=17, color=MUTED)
        text_box(slide, Inches(7.75), Inches(2.4), Inches(4.1), Inches(1.5), "Единый дизайн-код", font=FONT_HEADER, size=24, color=TEXT, bold=True)
        text_box(slide, Inches(7.75), Inches(4.0), Inches(4.0), Inches(1.2), "Структура, акцент и ритм работают вместе.", size=16, color=MUTED)


def build_presentation(dataset: list[dict[str, str]], palette: str, link: str) -> Path:
    presentation = Presentation()
    presentation.slide_width = SLIDE_WIDTH
    presentation.slide_height = SLIDE_HEIGHT
    accent = PALETTES.get(palette, PALETTES["cyber_blue"])
    blank = presentation.slide_layouts[6]
    for index, data in enumerate(dataset):
        slide = presentation.slides.add_slide(blank)
        render_slide(slide, data, accent, link, index)
    output = OUTPUT_DIR / f"generated_{uuid.uuid4().hex[:10]}.pptx"
    presentation.save(output)
    return output


PALETTE_OPTIONS = "".join(
    f'<option value="{key}">{key.replace("_", " ").title()}</option>'
    for key in PALETTES
)
HTML_UI = f"""<!doctype html>
<html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Presentation Brand Engine</title><style>
body{{margin:0;padding:40px;background:#020617;color:#f8fafc;font-family:'Segoe UI',sans-serif}}
main{{max-width:760px;margin:auto;padding:42px;background:#0f172a;border:1px solid #1e293b;border-radius:24px;box-shadow:0 24px 70px #0008}}
h1{{margin:0 0 12px;color:#38bdf8;font-size:34px}}p{{color:#94a3b8;line-height:1.6}}
label{{display:block;margin:24px 0 8px;color:#cbd5e1;font-weight:700;font-size:13px;text-transform:uppercase}}
input,select,button{{width:100%;box-sizing:border-box;padding:14px;border-radius:10px;font:inherit}}
input,select{{background:#020617;color:#fff;border:1px solid #334155}}button{{margin-top:28px;border:0;background:#38bdf8;color:#082f49;font-weight:800;cursor:pointer}}
button:hover{{background:#7dd3fc}}small{{color:#64748b}}
</style></head><body><main><h1>AI Presentation Brand Engine</h1>
<p>Загрузите презентацию, выберите акцент и скачайте новую версию с единым визуальным стилем.</p>
<form action="/engine-process" method="post" enctype="multipart/form-data">
<label for="file">Исходная презентация</label><input id="file" name="file" type="file" accept=".pptx" required>
<label for="palette">Палитра</label><select id="palette" name="palette">{PALETTE_OPTIONS}</select>
<label for="yandex_link">Ссылка для QR-кода (необязательно)</label><input id="yandex_link" name="yandex_link" type="url" placeholder="https://...">
<button type="submit">Собрать редизайн</button></form><small>Поддерживается формат .pptx</small></main></body></html>"""


@app.get("/", response_class=HTMLResponse)
async def root() -> str:
    return HTML_UI


@app.post("/engine-process")
async def engine_process(
    palette: str = Form("cyber_blue"),
    yandex_link: str = Form(""),
    file: UploadFile = File(...),
):
    if not file.filename or not file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Загрузите файл в формате .pptx.")
    data = await file.read()
    slides = parse_slides(data)
    if not slides:
        raise HTTPException(status_code=400, detail="В презентации не найден текстовый контент.")
    output = build_presentation(slides, palette, yandex_link)
    return FileResponse(output, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        filename=f"redesigned_{Path(file.filename).stem}.pptx")


if __name__ == "__main__":
    uvicorn.run("app:app", host=os.getenv("HOST", "0.0.0.0"),
                port=int(os.getenv("PORT", "8000")), reload=True)