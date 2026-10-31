import json
import tempfile
import unittest
from pathlib import Path

from pptx import Presentation

from pptforge.generator import build_from_json, build_presentation


class PresentationGeneratorTests(unittest.TestCase):
    def test_builds_editable_presentation_from_json(self) -> None:
        document = {
            "slides": [
                {"type": "title", "title": "测试标题"},
                {"type": "bullets", "title": "要点", "items": ["第一条", "第二条"]},
                {
                    "type": "chart",
                    "title": "趋势",
                    "categories": ["一月", "二月"],
                    "series": [{"name": "用户", "values": [2, 3]}],
                },
            ]
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "nested" / "test.pptx"
            build_presentation(document, output)
            presentation = Presentation(str(output))
            self.assertEqual(len(presentation.slides), 3)
            self.assertTrue(output.is_file())

    def test_reads_json_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "slides.json"
            output = Path(temp_dir) / "slides.pptx"
            source.write_text(json.dumps({"slides": [{"type": "text", "body": "Hello"}]}), encoding="utf-8")
            build_from_json(source, output)
            self.assertTrue(output.is_file())

    def test_rejects_unknown_slide_type(self) -> None:
        with self.assertRaisesRegex(ValueError, "类型无效"):
            build_presentation({"slides": [{"type": "unknown"}]}, "unused.pptx")

    def test_rejects_mismatched_chart_data(self) -> None:
        document = {
            "slides": [{
                "type": "chart",
                "categories": ["A", "B"],
                "series": [{"name": "数据", "values": [1]}],
            }]
        }
        with self.assertRaisesRegex(ValueError, "数量必须"):
            build_presentation(document, "unused.pptx")


if __name__ == "__main__":
    unittest.main()
