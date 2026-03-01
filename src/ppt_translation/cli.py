from __future__ import annotations

import argparse
from pathlib import Path
import sys

from .translator import translate_presentation


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Translate a .ppt/.pptx file from English to Chinese.",
    )
    parser.add_argument("input", type=Path, help="Input .ppt/.pptx file path")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output .pptx file path (default: <input>_zh.pptx)",
    )
    parser.add_argument(
        "--source-lang",
        default="en",
        help="Source language code (default: en)",
    )
    parser.add_argument(
        "--target-lang",
        default="zh-CN",
        help="Target language code (default: zh-CN)",
    )
    parser.add_argument(
        "--include-notes",
        action="store_true",
        help="Also translate speaker notes.",
    )
    parser.add_argument(
        "--glossary-file",
        type=Path,
        default=None,
        help="Optional CSV glossary file: source_term,target_term",
    )
    return parser


def _default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_zh.pptx")


def main() -> int:
    args = _build_parser().parse_args()
    input_path = args.input

    if input_path.suffix.lower() not in {".ppt", ".pptx"}:
        print("Error: input file must be a .ppt or .pptx file.", file=sys.stderr)
        return 2

    if not input_path.exists():
        print(f"Error: input file not found: {input_path}", file=sys.stderr)
        return 2

    output_path = args.output or _default_output_path(input_path)
    if output_path.suffix.lower() != ".pptx":
        print("Error: output file must end with .pptx.", file=sys.stderr)
        return 2

    glossary_path = args.glossary_file
    if glossary_path is None:
        default_glossary = Path.cwd() / "domain_glossary.csv"
        if default_glossary.exists():
            glossary_path = default_glossary

    try:
        stats = translate_presentation(
            input_path=input_path,
            output_path=output_path,
            source_lang=args.source_lang,
            target_lang=args.target_lang,
            include_notes=args.include_notes,
            glossary_file=glossary_path,
        )
    except Exception as exc:
        print(f"Translation failed: {exc}", file=sys.stderr)
        return 1

    print(f"Done. Saved translated file: {output_path}")
    print(f"Slides: {stats.slides}")
    print(f"Text segments seen: {stats.text_segments_seen}")
    print(f"Text segments translated: {stats.text_segments_translated}")
    print(f"Cache hits: {stats.cache_hits}")
    print(f"Glossary replacements: {stats.glossary_replacements}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
