from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt


_DEFAULT_THEME = {
    "primary": "17365D",
    "accent": "2F75B5",
    "background": "FFFFFF",
    "text": "263238",
    "muted": "667085",
    "font": "Aptos",
}
_CHART_TYPES = {
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "line": XL_CHART_TYPE.LINE,
}
_SUPPORTED_SLIDES = {
    "title",
    "section",
    "bullets",
    "text",
    "two_column",
    "quote",
    "stats",
    "image",
    "chart",
}


def _color(value: str) -> RGBColor:
    value = value.strip().lstrip("#")
    if len(value) != 6:
        raise ValueError(f"颜色必须是 6 位十六进制值，例如 17365D：{value!r}")
    try:
        return RGBColor.from_string(value.upper())
    except ValueError as exc:
        raise ValueError(f"无效的十六进制颜色：{value!r}") from exc


class PresentationGenerator:
    """Build 16:9 PowerPoint files with reusable layouts and theme settings."""

    def __init__(self, theme: dict[str, Any] | None = None) -> None:
        self.theme = {**_DEFAULT_THEME, **(theme or {})}
        for key in ("primary", "accent", "background", "text", "muted"):
            _color(str(self.theme[key]))
        self.presentation = Presentation()
        self.presentation.slide_width = Inches(13.333)
        self.presentation.slide_height = Inches(7.5)

    def add_slide(self, data: dict[str, Any], index: int) -> None:
        if not isinstance(data, dict):
            raise ValueError(f"第 {index} 页必须是对象")
        kind = data.get("type")
        if kind not in _SUPPORTED_SLIDES:
            supported = "、".join(sorted(_SUPPORTED_SLIDES))
            raise ValueError(f"第 {index} 页类型无效：{kind!r}；支持类型：{supported}")

        slide = self.presentation.slides.add_slide(self.presentation.slide_layouts[6])
        self._set_background(slide)
        if kind == "title":
            self._title_slide(slide, data)
        elif kind == "section":
            self._section_slide(slide, data)
        else:
            self._add_heading(slide, data.get("title", ""))
            getattr(self, f"_layout_{kind}")(slide, data)
            self._add_footer(slide, index)

    def save(self, output_path: str | Path) -> Path:
        path = Path(output_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self.presentation.save(str(path))
        return path

    def _set_background(self, slide: Any) -> None:
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = _color(self.theme["background"])

    def _add_text(
        self,
        slide: Any,
        text: Any,
        x: float,
        y: float,
        width: float,
        height: float,
        *,
        size: int = 20,
        color: str | None = None,
        bold: bool = False,
        align: PP_ALIGN = PP_ALIGN.LEFT,
        valign: MSO_ANCHOR = MSO_ANCHOR.TOP,
    ) -> Any:
        box = slide.shapes.add_textbox(
            Inches(x), Inches(y), Inches(width), Inches(height)
        )
        frame = box.text_frame
        frame.clear()
        frame.word_wrap = True
        frame.margin_left = Inches(0.02)
        frame.margin_right = Inches(0.02)
        frame.margin_top = Inches(0.02)
        frame.margin_bottom = Inches(0.02)
        frame.vertical_anchor = valign
        paragraph = frame.paragraphs[0]
        paragraph.alignment = align
        run = paragraph.add_run()
        run.text = str(text)
        run.font.name = str(self.theme["font"])
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = _color(color or self.theme["text"])
        return box

    def _add_heading(self, slide: Any, title: str) -> None:
        self._add_text(
            slide,
            title or "未命名页面",
            0.7,
            0.42,
            11.9,
            0.65,
            size=28,
            color=self.theme["primary"],
            bold=True,
        )
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.72), Inches(1.2), Inches(0.75), Inches(0.06)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = _color(self.theme["accent"])
        line.line.fill.background()

    def _add_footer(self, slide: Any, index: int) -> None:
        self._add_text(
            slide,
            f"{index:02d}",
            12.15,
            7.08,
            0.45,
            0.2,
            size=9,
            color=self.theme["muted"],
            align=PP_ALIGN.RIGHT,
        )

    def _title_slide(self, slide: Any, data: dict[str, Any]) -> None:
        panel = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(4.0), Inches(7.5)
        )
        panel.fill.solid()
        panel.fill.fore_color.rgb = _color(self.theme["primary"])
        panel.line.fill.background()
        self._add_text(
            slide,
            data.get("title", "演示文稿标题"),
            4.65,
            2.2,
            7.7,
            1.35,
            size=36,
            color=self.theme["primary"],
            bold=True,
            valign=MSO_ANCHOR.MIDDLE,
        )
        self._add_text(
            slide,
            data.get("subtitle", ""),
            4.72,
            3.75,
            7.2,
            0.7,
            size=20,
            color=self.theme["muted"],
        )
        if data.get("author"):
            self._add_text(
                slide, data["author"], 4.72, 5.85, 7.2, 0.35, size=14,
                color=self.theme["muted"],
            )

    def _section_slide(self, slide: Any, data: dict[str, Any]) -> None:
        panel = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5)
        )
        panel.fill.solid()
        panel.fill.fore_color.rgb = _color(self.theme["primary"])
        panel.line.fill.background()
        self._add_text(
            slide, data.get("title", "章节标题"), 1.0, 2.55, 11.3, 1.0,
            size=34, color="FFFFFF", bold=True, align=PP_ALIGN.CENTER,
        )
        self._add_text(
            slide, data.get("subtitle", ""), 1.1, 3.7, 11.1, 0.65,
            size=18, color="EAF1F8", align=PP_ALIGN.CENTER,
        )

    def _layout_bullets(self, slide: Any, data: dict[str, Any]) -> None:
        items = data.get("items", [])
        if not isinstance(items, list):
            raise ValueError("bullets 页的 items 必须是数组")
        box = slide.shapes.add_textbox(
            Inches(0.95), Inches(1.65), Inches(11.5), Inches(4.95)
        )
        frame = box.text_frame
        frame.clear()
        frame.word_wrap = True
        for i, item in enumerate(items):
            p = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
            p.text = str(item)
            p.level = 0
            p.space_after = Pt(19)
            p.font.name = str(self.theme["font"])
            p.font.size = Pt(22)
            p.font.color.rgb = _color(self.theme["text"])
            bullet = OxmlElement("a:buChar")
            bullet.set("char", "•")
            p._p.get_or_add_pPr().insert(0, bullet)

    def _layout_text(self, slide: Any, data: dict[str, Any]) -> None:
        self._add_text(
            slide, data.get("body", ""), 0.95, 1.65, 11.4, 4.9,
            size=21, color=self.theme["text"],
        )

    def _layout_two_column(self, slide: Any, data: dict[str, Any]) -> None:
        for x, key, heading_key in (
            (0.9, "left_items", "left_title"),
            (6.85, "right_items", "right_title"),
        ):
            self._add_text(
                slide, data.get(heading_key, ""), x, 1.55, 5.2, 0.45,
                size=19, color=self.theme["accent"], bold=True,
            )
            items = data.get(key, [])
            if not isinstance(items, list):
                raise ValueError(f"{key} 必须是数组")
            self._add_text(
                slide, "\n".join(f"•  {item}" for item in items),
                x, 2.15, 5.35, 4.25, size=18,
            )

    def _layout_quote(self, slide: Any, data: dict[str, Any]) -> None:
        self._add_text(
            slide, f"“{data.get('quote', '')}”", 1.3, 2.0, 10.7, 2.6,
            size=28, color=self.theme["primary"], bold=True, align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
        )
        if data.get("author"):
            self._add_text(
                slide, f"— {data['author']}", 2.0, 5.0, 9.3, 0.45,
                size=16, color=self.theme["muted"], align=PP_ALIGN.CENTER,
            )

    def _layout_stats(self, slide: Any, data: dict[str, Any]) -> None:
        items = data.get("items", [])
        if not isinstance(items, list) or not items:
            raise ValueError("stats 页的 items 必须是非空数组")
        card_width = min(3.5, 11.5 / len(items))
        total_width = card_width * len(items)
        start_x = (13.333 - total_width) / 2
        for i, item in enumerate(items):
            if not isinstance(item, dict):
                raise ValueError("stats 的每个项目都必须是对象")
            x = start_x + i * card_width
            shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(x + 0.08), Inches(2.35), Inches(card_width - 0.16), Inches(2.7),
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = _color("F3F6FA")
            shape.line.color.rgb = _color("E4EAF1")
            self._add_text(
                slide, item.get("value", ""), x + 0.15, 2.85, card_width - 0.3, 0.75,
                size=30, color=self.theme["accent"], bold=True, align=PP_ALIGN.CENTER,
            )
            self._add_text(
                slide, item.get("label", ""), x + 0.15, 3.85, card_width - 0.3, 0.55,
                size=16, color=self.theme["muted"], align=PP_ALIGN.CENTER,
            )

    def _layout_image(self, slide: Any, data: dict[str, Any]) -> None:
        image_path = Path(str(data.get("_base_dir", "."))) / str(data.get("image", ""))
        if not data.get("image") or not image_path.is_file():
            raise ValueError(f"图片不存在：{image_path}")
        picture = slide.shapes.add_picture(
            str(image_path), Inches(1.0), Inches(1.55), height=Inches(4.95)
        )
        max_width, max_height = Inches(11.2), Inches(4.95)
        scale = min(max_width / picture.width, max_height / picture.height)
        picture.width = int(picture.width * scale)
        picture.height = int(picture.height * scale)
        picture.left = int((Inches(13.333) - picture.width) / 2)
        picture.top = Inches(1.55) + int((max_height - picture.height) / 2)
        if data.get("caption"):
            self._add_text(
                slide, data["caption"], 1.0, 6.55, 11.2, 0.3, size=12,
                color=self.theme["muted"], align=PP_ALIGN.CENTER,
            )

    def _layout_chart(self, slide: Any, data: dict[str, Any]) -> None:
        chart_type = data.get("chart_type", "column")
        if chart_type not in _CHART_TYPES:
            raise ValueError("chart_type 只支持 bar、column 或 line")
        categories = data.get("categories", [])
        series = data.get("series", [])
        if not categories or not series:
            raise ValueError("chart 页需要提供 categories 和 series")
        chart_data = CategoryChartData()
        chart_data.categories = categories
        for item in series:
            if not isinstance(item, dict) or "name" not in item or "values" not in item:
                raise ValueError("series 项目必须包含 name 和 values")
            if len(item["values"]) != len(categories):
                raise ValueError("每组 series.values 的数量必须与 categories 一致")
            chart_data.add_series(item["name"], item["values"])
        chart = slide.shapes.add_chart(
            _CHART_TYPES[chart_type],
            Inches(1.0), Inches(1.55), Inches(11.3), Inches(4.95),
            chart_data,
        ).chart
        chart.has_legend = len(series) > 1
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False


def build_presentation(
    document: dict[str, Any], output_path: str | Path, *, base_dir: str | Path = "."
) -> Path:
    """Create a presentation from a JSON-compatible dictionary."""
    slides = document.get("slides")
    if not isinstance(slides, list) or not slides:
        raise ValueError("文档必须包含非空 slides 数组")
    generator = PresentationGenerator(document.get("theme"))
    for index, slide_data in enumerate(slides, start=1):
        slide = dict(slide_data) if isinstance(slide_data, dict) else slide_data
        if isinstance(slide, dict):
            slide["_base_dir"] = str(base_dir)
        generator.add_slide(slide, index)
    return generator.save(output_path)


def build_from_json(input_path: str | Path, output_path: str | Path) -> Path:
    """Load a JSON presentation description and write the resulting PPTX."""
    source = Path(input_path)
    try:
        document = json.loads(source.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"JSON 格式错误（第 {exc.lineno} 行）：{exc.msg}") from exc
    if not isinstance(document, dict):
        raise ValueError("JSON 顶层必须是对象")
    return build_presentation(document, output_path, base_dir=source.parent)
