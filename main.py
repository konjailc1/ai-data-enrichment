"""CLI: làm sạch và làm giàu dữ liệu bán lẻ / clean and enrich retail data."""

from __future__ import annotations

import argparse
import sys
import traceback
from pathlib import Path

from dotenv import load_dotenv

from enrich import MOCK_PATH, EnrichmentError, print_summary, redact, run_pipeline


def _utf8_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8")
        except Exception:
            pass


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="main.py",
        description=(
            "Làm sạch và làm giàu dữ liệu bán lẻ tiếng Việt thành báo cáo Excel song ngữ. "
            "Clean and enrich a messy Vietnamese retail sheet into a bilingual Excel report."
        ),
        epilog=(
            "Ví dụ / Example: python main.py sample_input/shop_data.xlsx "
            "-o output/enriched.xlsx --batch-size 15"
        ),
    )
    parser.add_argument("input", help="CSV hoặc Excel đầu vào / Input CSV or Excel")
    parser.add_argument(
        "-o",
        "--output",
        default="output/enriched.xlsx",
        help="File Excel kết quả / Output workbook (default: output/enriched.xlsx)",
    )
    parser.add_argument(
        "--model",
        default="gemini-3.5-flash-lite",
        help=(
            "Mô hình Gemini / Gemini model (default: gemini-3.5-flash-lite). "
            "gemini-2.5-flash-lite trả 404 với API key mới; nếu gặp lỗi đó, "
            "chương trình tự chuyển sang model Google gợi ý."
        ),
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=15,
        help="Số dòng mỗi lần gọi / Rows per Gemini call (default: 15)",
    )
    parser.add_argument(
        "--offline",
        action="store_true",
        help="Dùng mock, không gọi API / Use the bundled mock, no API calls",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Giống --offline / Same as --offline",
    )
    parser.add_argument(
        "--no-cache",
        action="store_true",
        help="Bỏ qua cache trên đĩa / Ignore the on-disk response cache",
    )
    parser.add_argument(
        "--preview-image",
        default=None,
        help="Xuất ảnh xem trước sheet dữ liệu sạch / PNG preview of the clean-data sheet",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    _utf8_stdio()
    load_dotenv()
    args = build_parser().parse_args(argv)
    preview = Path(args.preview_image) if args.preview_image else None
    try:
        stats = run_pipeline(
            Path(args.input),
            Path(args.output),
            model=args.model,
            batch_size=args.batch_size,
            offline=args.offline or args.dry_run,
            use_cache=not args.no_cache,
            mock_path=MOCK_PATH,
            preview_image=preview,
        )
    except EnrichmentError as exc:
        print(redact(str(exc)), file=sys.stderr)
        return exc.exit_code
    except Exception as exc:
        print(
            "Lỗi không xác định / Unexpected error: " + redact(str(exc)),
            file=sys.stderr,
        )
        traceback.print_exc()
        return 1
    print_summary(stats)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
