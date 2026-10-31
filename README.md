# SlideForge

SlideForge 是一个开源 Python 命令行工具和轻量级 Python 库：用 JSON 配置主题与页面内容，生成可在 PowerPoint、Keynote、LibreOffice 中继续编辑的 `.pptx` 文件。适合快速制作汇报、课程、方案、产品介绍、数据展示等演示文稿。

## 功能

- 16:9 页面与可配置配色、字体
- 标题页、章节页、要点列表、正文、双栏、引用、数据卡片、图片、原生图表
- 图表支持柱形、条形和折线类型
- 图片使用相对于 JSON 文件的路径
- 可通过命令行使用，也可在 Python 项目中调用

## 安装

需要 Python 3.10 或更新版本。

```bash
python -m pip install .
```

开发时可使用可编辑安装：

```bash
python -m pip install -e .
```

## 快速开始

```bash
slideforge examples/sample.json -o output/quarterly-review.pptx
```

也可以不安装命令行入口：

```bash
python -m pptforge.cli examples/sample.json -o output/demo.pptx
```

## JSON 格式

查看 [`examples/sample.json`](examples/sample.json) 获取可运行的完整示例。最小配置如下：

```json
{
  "theme": {
    "primary": "17365D",
    "accent": "2F75B5",
    "background": "FFFFFF",
    "text": "263238",
    "muted": "667085",
    "font": "Microsoft YaHei"
  },
  "slides": [
    { "type": "title", "title": "演示标题", "subtitle": "副标题", "author": "作者" },
    { "type": "bullets", "title": "要点", "items": ["要点一", "要点二"] },
    { "type": "text", "title": "正文", "body": "这里放一段文字。" },
    {
      "type": "two_column",
      "title": "对比",
      "left_title": "方案 A",
      "left_items": ["特点一"],
      "right_title": "方案 B",
      "right_items": ["特点二"]
    },
    {
      "type": "stats",
      "title": "指标",
      "items": [{ "label": "增长率", "value": "28%" }]
    },
    { "type": "quote", "title": "引用", "quote": "一句重要的话。", "author": "作者" },
    { "type": "section", "title": "第二部分", "subtitle": "章节说明" },
    { "type": "image", "title": "现场照片", "image": "images/photo.jpg", "caption": "图片说明" },
    {
      "type": "chart",
      "title": "月度趋势",
      "chart_type": "line",
      "categories": ["一月", "二月"],
      "series": [{ "name": "用户数", "values": [120, 160] }]
    }
  ]
}
```

支持的 `chart_type` 为 `bar`、`column`、`line`。颜色使用不带 `#` 的 6 位十六进制值。图片类型支持 `python-pptx` 可读取的图片格式；图片文件路径相对于 JSON 文件所在目录。

## Python API

```python
from pptforge import build_presentation

build_presentation(
    {
        "theme": {"primary": "17365D", "accent": "2F75B5"},
        "slides": [
            {"type": "title", "title": "用 Python 生成", "subtitle": "可编辑的演示文稿"},
            {"type": "bullets", "title": "优势", "items": ["可复用", "易修改"]},
        ],
    },
    "presentation.pptx",
)
```

## 测试

```bash
python -m unittest discover -s tests
```

## 发布前检查

建议在目标平台上试生成并打开 `examples/sample.json` 的输出文件；字体是否可用取决于查看者的操作系统。项目作者信息与 GitHub 仓库地址已配置在 `pyproject.toml` 中。

## License

MIT
