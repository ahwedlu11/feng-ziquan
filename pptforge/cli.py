from __future__ import annotations

import argparse
import sys

from .generator import build_from_json


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="slideforge",
        description="从 JSON 描述文件生成可编辑的 PowerPoint 演示文稿。",
    )
    parser.add_argument("input", help="演示文稿 JSON 文件")
    parser.add_argument(
        "-o", "--output", default="presentation.pptx", help="输出文件（默认：presentation.pptx）"
    )
    args = parser.parse_args()
    try:
        result = build_from_json(args.input, args.output)
    except (OSError, ValueError, TypeError) as exc:
        print(f"生成失败：{exc}", file=sys.stderr)
        raise SystemExit(2) from exc
    print(f"已生成：{result.resolve()}")


if __name__ == "__main__":
    main()
