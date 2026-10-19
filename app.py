"""Streamlit demo for the Vietnamese retail enrichment pipeline.

The page calls enrich.py. It does not reimplement price, phone, Gemini, or Excel logic.
"""

from __future__ import annotations

import hashlib
import math
import numbers
import os
from urllib.parse import quote
from datetime import datetime
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from enrich import (
    CACHE_DIR,
    DEFAULT_MODEL,
    MOCK_PATH,
    REQUIRED_COLUMNS,
    STATUS_OK,
    EnrichmentError,
    RunStats,
    clean_frame,
    enrich_rows,
    load_table,
    redact,
    rows_from_frame,
    write_workbook,
)

ROOT = Path(__file__).resolve().parent
SAMPLE_PATH = ROOT / "sample_input" / "shop_data.xlsx"
VN_TZ = ZoneInfo("Asia/Ho_Chi_Minh")
REPO_URL = "https://github.com/konjailc1/ai-data-enrichment"
PROFILE_URL = "https://github.com/konjailc1"

MAX_UPLOAD_BYTES = 1_000_000
MAX_UPLOAD_ROWS = 50
MAX_SESSION_RUNS = 3
DAILY_API_CAP = 100
WEB_BATCH_SIZE = 10

GITHUB_MARK = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16" width="16" height="16" fill="currentColor" aria-hidden="true">'
    '<path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 '
    "0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15"
    "-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07"
    "-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 "
    "2.2.82.64-.18 1.32-.27 2-.27s1.36.09 2 .27c1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 "
    "2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 "
    '1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0 0 16 8c0-4.42-3.58-8-8-8z"/></svg>'
)

MODEL_CHOICES = list(
    dict.fromkeys(
        [
            DEFAULT_MODEL,
            "gemini-2.5-flash-lite",
            "gemini-2.5-flash",
            "gemini-3.5-flash",
        ]
    )
)

COPY = {
    "vi": {
        "title": "AI Data Enrichment — Chuẩn hoá dữ liệu bán lẻ",
        "tagline": "Chuẩn hoá giá, số điện thoại, sản phẩm, địa chỉ Hà Nội và phản hồi khách — ra một file Excel.",
        "privacy": (
            "**Lưu ý:**\n"
            "- Chỉ tải dữ liệu mẫu, không tải dữ liệu thật hoặc dữ liệu nhạy cảm.\n"
            "- Bản demo dùng gói miễn phí của Gemini AI. Dữ liệu mẫu có sẵn không tính lượt xử lý bằng AI."
        ),
        "language": "Ngôn ngữ",
        "model": "Mô hình AI",
        "remaining": "Lượt xử lý của bạn còn lại",
        "today": "Lượt xử lý AI còn lại hôm nay",
        "key_on": "Dữ liệu được chuẩn hoá bằng Gemini AI",
        "key_off": "Chế độ AI đang tắt nên chưa xử lý được file tải lên, vui lòng dùng dữ liệu mẫu",
        "source": "Nguồn dữ liệu",
        "sample": "Dùng dữ liệu mẫu",
        "upload": "Tải file của bạn",
        "sample_note": "Dữ liệu mẫu {n} dòng của một cửa hàng tạp hoá giả lập. Kết quả đã được xử lý sẵn, không tính lần xử lý bằng AI.",
        "expected": "Cột bắt buộc",
        "template": "Tải file mẫu",
        "uploader": "File .xlsx hoặc .csv",
        "preview_sample": "Dữ liệu mẫu: Cửa hàng bán lẻ",
        "preview_upload": "Xem trước dữ liệu của bạn",
        "rows_in_file": "Số dòng xử lý: {n}",
        "truncate": "File có {n} dòng, demo giữ 50 dòng đầu.",
        "too_big": "File vượt 1 MB.",
        "bad_type": "Chỉ nhận file .xlsx hoặc .csv.",
        "empty": "File không có dòng dữ liệu.",
        "run": "Chạy",
        "running": "Đang xử lý…",
        "progress": "Bước {done}/{total}",
        "session_limit": "Bạn đã dùng hết 3 lượt xử lý trong phiên này. Hãy dùng dữ liệu mẫu.",
        "no_key_run": "Chế độ AI đang tắt nên chưa xử lý được file tải lên, vui lòng dùng dữ liệu mẫu",
        "memory_cache": "File này đã được xử lý với cùng mô hình AI. Không tính thêm lượt.",
        "download": "Tải file Excel",
        "clean_tab": "Dữ liệu sạch",
        "summary_tab": "Tổng hợp",
        "review_tab": "Cần kiểm tra",
        "kpi_rows": "Dòng",
        "kpi_ok": "OK",
        "kpi_review": "Cần kiểm tra",
        "kpi_api": "API calls",
        "kpi_cache": "Cache hits",
        "kpi_tokens": "Tokens",
        "tech_details": "Chi tiết kỹ thuật",
        "categories": "Số dòng theo danh mục",
        "sentiment": "Cảm xúc phản hồi",
        "no_review": "Không có dòng cần kiểm tra.",
        "daily_limit": "Demo đã hết lượt hôm nay, hãy thử dữ liệu mẫu",
        "unexpected": "Lỗi không xác định: {exc}",
        "footer": "Được xây dựng bằng Python, pandas, Gemini API và Streamlit",
        "profile": "GitHub của Quang",
        "estimate": "Tối đa khoảng {n} lượt xử lý AI nếu file này chưa được xử lý.",
        "show_raw": "Hiện cột dữ liệu gốc",
    },
    "en": {
        "title": "AI Data Enrichment — Retail data standardization",
        "tagline": "Standardizes prices, phone numbers, products, Hanoi addresses and customer feedback into one Excel file.",
        "privacy": (
            "**Note:**\n"
            "- Upload sample data only, no real or sensitive data.\n"
            "- This demo uses the free tier of Gemini AI. The built-in sample data doesn't count as AI runs."
        ),
        "language": "Language",
        "model": "AI model",
        "remaining": "Your remaining runs",
        "today": "AI runs left today",
        "key_on": "Data is standardized with Gemini AI",
        "key_off": "AI mode is off, uploads can't be processed. Please use the sample data",
        "source": "Data source",
        "sample": "Use sample data",
        "upload": "Upload your file",
        "sample_note": "{n} sample rows from a fictional grocery store. Results are pre-processed and don't count as AI runs.",
        "expected": "Required columns",
        "template": "Download sample template",
        "uploader": ".xlsx or .csv file",
        "preview_sample": "Sample data: Retail store",
        "preview_upload": "Preview of your data",
        "rows_in_file": "Rows to process: {n}",
        "truncate": "The file has {n} rows; the demo keeps the first 50.",
        "too_big": "File is larger than 1 MB.",
        "bad_type": "Only .xlsx or .csv files are accepted.",
        "empty": "The file has no data rows.",
        "run": "Run",
        "running": "Processing…",
        "progress": "Step {done}/{total}",
        "session_limit": "You have used all 3 runs in this session. Please use the sample data.",
        "no_key_run": "AI mode is off, uploads can't be processed. Please use the sample data",
        "memory_cache": "This file was already processed with the same AI model. No extra run is counted.",
        "download": "Download Excel file",
        "clean_tab": "Clean data",
        "summary_tab": "Summary",
        "review_tab": "Review",
        "kpi_rows": "Rows",
        "kpi_ok": "OK",
        "kpi_review": "Needs review",
        "kpi_api": "API calls",
        "kpi_cache": "Cache hits",
        "kpi_tokens": "Tokens",
        "tech_details": "Technical details",
        "categories": "Rows by category",
        "sentiment": "Feedback sentiment",
        "no_review": "No rows need review.",
        "daily_limit": "Demo limit reached today, try the sample data",
        "unexpected": "Unexpected error: {exc}",
        "footer": "Built with Python, pandas, Gemini API and Streamlit",
        "profile": "Quang's GitHub",
        "estimate": "Up to about {n} AI runs if this file has not been processed yet.",
        "show_raw": "Show raw columns",
    },
}

# Source columns from clean_frame → (vi header, en header, flags).
# "pair" uses the VI or EN source column; "split" keeps one bilingual source column.
_DISPLAY_COLUMNS: list[tuple[str, str, str, str]] = [
    ("STT / No.", "STT", "No.", "keep"),
    ("Sản phẩm gốc / Raw product", "Sản phẩm gốc", "Raw product", "raw_always"),
    ("Tên sạch (VI) / Clean name (VI)", "Tên sạch", "Clean name", "vi"),
    ("Tên sạch (EN) / Clean name (EN)", "Tên sạch", "Clean name", "en"),
    ("Thương hiệu / Brand", "Thương hiệu", "Brand", "keep"),
    ("Danh mục (VI) / Category (VI)", "Danh mục", "Category", "vi"),
    ("Danh mục (EN) / Category (EN)", "Danh mục", "Category", "en"),
    ("Định lượng / Size", "Định lượng", "Size", "keep"),
    ("Đơn vị / Unit", "Đơn vị", "Unit", "keep"),
    ("Quy cách / Pack qty", "Quy cách", "Pack qty", "keep"),
    ("Giá gốc / Raw price", "Giá gốc", "Raw price", "raw_optional"),
    ("Giá (VND) / Price (VND)", "Giá (VND)", "Price (VND)", "keep"),
    ("Khách hàng / Customer", "Khách hàng", "Customer", "keep"),
    ("SĐT gốc / Raw phone", "SĐT gốc", "Raw phone", "raw_optional"),
    ("SĐT / Phone", "SĐT", "Phone", "keep"),
    ("Địa chỉ gốc / Raw address", "Địa chỉ gốc", "Raw address", "raw_optional"),
    ("Địa chỉ / Address", "Địa chỉ", "Address", "keep"),
    ("Phản hồi gốc / Raw feedback", "Phản hồi gốc", "Raw feedback", "raw_always"),
    ("Cảm xúc / Sentiment", "Cảm xúc", "Sentiment", "split"),
    ("Tóm tắt (VI) / Summary (VI)", "Tóm tắt", "Summary", "vi"),
    ("Tóm tắt (EN) / Summary (EN)", "Tóm tắt", "Summary", "en"),
    ("Khiếu nại / Complaint", "Khiếu nại", "Complaint", "split"),
    ("Độ tin cậy / Confidence", "Độ tin cậy", "Confidence", "keep"),
    ("Trạng thái / Status", "Trạng thái", "Status", "split"),
    ("Lý do / Reason", "Lý do", "Reason", "split"),
]


def ui_text() -> dict[str, str]:
    return COPY[st.session_state.get("lang_widget", "vi")]


def resolve_api_key() -> str:
    """Secrets win, then the environment or .env. The value is never displayed."""
    try:
        if "GEMINI_API_KEY" in st.secrets:
            secret = str(st.secrets["GEMINI_API_KEY"] or "").strip()
            if secret:
                return secret
    except Exception:
        pass
    load_dotenv(ROOT / ".env")
    return os.getenv("GEMINI_API_KEY", "").strip()


def apply_api_key() -> str:
    key = resolve_api_key()
    if key:
        os.environ["GEMINI_API_KEY"] = key
    return key


@st.cache_resource
def _daily_api_counter(day: str) -> dict[str, int]:
    """Process-wide call counter. The date argument is the cache key."""
    del day
    return {"calls": 0}


def daily_counter() -> dict[str, int]:
    day = datetime.now(VN_TZ).date().isoformat()
    return _daily_api_counter(day)


def init_state() -> None:
    st.session_state.setdefault("live_runs", 0)
    st.session_state.setdefault("upload_cache", {})
    st.session_state.setdefault("result", None)


def show_error(exc: BaseException) -> None:
    st.error(redact(str(exc)))


def frame_from_bytes(name: str, data: bytes) -> pd.DataFrame:
    suffix = Path(name).suffix.lower()
    if suffix not in {".csv", ".xlsx"}:
        raise EnrichmentError(ui_text()["bad_type"])
    buffer = BytesIO(data)
    try:
        if suffix == ".csv":
            sample = data[:4000].decode("utf-8-sig", errors="replace")
            sep = ";" if sample.count(";") > sample.count(",") else ","
            frame = pd.read_csv(buffer, dtype=str, sep=sep, encoding="utf-8-sig")
        else:
            frame = pd.read_excel(buffer, dtype=str, engine="openpyxl")
    except EnrichmentError:
        raise
    except Exception as exc:
        raise EnrichmentError(
            f"Không đọc được file / Could not read the file: {name} ({exc})"
        ) from exc
    missing = [column for column in REQUIRED_COLUMNS if column not in frame.columns]
    if missing:
        raise EnrichmentError(
            "Thiếu cột / Missing columns: "
            + ", ".join(missing)
            + "\nCần có / Expected: "
            + ", ".join(REQUIRED_COLUMNS)
        )
    return frame


def prepare_upload(name: str, data: bytes) -> tuple[pd.DataFrame, str, int]:
    """Return the (possibly truncated) frame, content hash, and original row count."""
    if len(data) > MAX_UPLOAD_BYTES:
        raise EnrichmentError(ui_text()["too_big"])
    digest = hashlib.sha256(data).hexdigest()
    cached = st.session_state.get("prepared")
    if cached and cached["digest"] == digest and cached["name"] == name:
        return cached["frame"], digest, int(cached["original_rows"])
    frame = frame_from_bytes(name, data)
    original_rows = len(frame)
    if original_rows > MAX_UPLOAD_ROWS:
        frame = frame.iloc[:MAX_UPLOAD_ROWS].copy()
    st.session_state.prepared = {
        "digest": digest,
        "name": name,
        "frame": frame,
        "original_rows": original_rows,
    }
    return frame, digest, original_rows


def run_enrichment(
    frame: pd.DataFrame,
    source_name: str,
    model: str,
    *,
    offline: bool,
    progress: object | None = None,
    progress_label: str = "Step {done}/{total}",
    charge: dict[str, int] | None = None,
) -> dict[str, object]:
    sources = rows_from_frame(frame)
    stats = RunStats(
        rows=len(sources),
        offline=offline,
        model=model,
        output_path="memory",
    )

    def on_progress(done: int, total: int) -> None:
        if progress is None or total <= 0:
            return
        progress.progress(done / total, text=progress_label.format(done=done, total=total))  # type: ignore[attr-defined]

    try:
        cleaned = enrich_rows(
            sources,
            model=model,
            batch_size=WEB_BATCH_SIZE,
            offline=offline,
            use_cache=True,
            mock_path=MOCK_PATH,
            stats=stats,
            cache_dir=CACHE_DIR,
            on_progress=on_progress,
        )
    finally:
        # A failed batch can still have spent calls. Count those against the daily cap.
        if charge is not None and stats.api_calls:
            charge["calls"] += stats.api_calls
            st.session_state.live_runs = int(st.session_state.live_runs) + 1
    stats.ok = sum(1 for row in cleaned if row.status == STATUS_OK)
    stats.needs_review = stats.rows - stats.ok
    table = clean_frame(cleaned)
    buffer = BytesIO()
    write_workbook(cleaned, buffer, stats, source_name)
    return {
        "frame": table,
        "excel": buffer.getvalue(),
        "stats": stats,
        "source_name": source_name,
        "from_memory_cache": False,
    }


def replay_cached(stored: dict[str, object]) -> dict[str, object]:
    original: RunStats = stored["stats"]  # type: ignore[assignment]
    stats = RunStats(
        rows=original.rows,
        ok=original.ok,
        needs_review=original.needs_review,
        api_calls=0,
        cache_hits=original.cache_hits + original.api_calls,
        input_tokens=0,
        output_tokens=0,
        total_tokens=0,
        offline=False,
        model=original.model,
        output_path="memory",
    )
    return {
        "frame": stored["frame"],
        "excel": stored["excel"],
        "stats": stats,
        "source_name": stored["source_name"],
        "from_memory_cache": True,
    }


_NUMERIC_HEADERS = {
    "vi": ("STT", "Định lượng", "Quy cách", "Giá (VND)", "Độ tin cậy"),
    "en": ("No.", "Size", "Pack qty", "Price (VND)", "Confidence"),
}
_PREVIEW_HEADERS = {
    "raw_product": ("Sản phẩm", "Product"),
    "raw_price": ("Giá", "Price"),
    "customer_name": ("Khách hàng", "Customer"),
    "phone": ("SĐT", "Phone"),
    "address": ("Địa chỉ", "Address"),
    "feedback": ("Phản hồi", "Feedback"),
}
_FORMAT_BY_SOURCE = {
    "STT / No.": "id",
    "Định lượng / Size": "measure",
    "Quy cách / Pack qty": "measure",
    "Giá (VND) / Price (VND)": "vnd",
    "Độ tin cậy / Confidence": "confidence",
}


def _is_blank(value: object) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(value, str) and value.strip() == ""


def _as_number(value: object) -> float | None:
    if _is_blank(value):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, numbers.Real):
        try:
            if pd.isna(value):
                return None
        except (TypeError, ValueError):
            pass
        return float(value)
    return None


def format_grouped_int(value: int | float, lang: str) -> str:
    text = f"{int(round(value)):,}"
    if lang == "vi":
        return text.replace(",", ".")
    return text


def format_measure(value: object, lang: str) -> str:
    number = _as_number(value)
    if number is None:
        return "N/A"
    if number.is_integer():
        return str(int(number))
    text = f"{number:.2f}".rstrip("0").rstrip(".")
    if lang == "vi":
        return text.replace(".", ",")
    return text


def format_vnd(value: object, lang: str) -> str:
    number = _as_number(value)
    if number is None:
        return "N/A"
    return format_grouped_int(number, lang)


def format_confidence(value: object) -> str:
    number = _as_number(value)
    if number is None:
        return "N/A"
    return f"{int(round(number * 100))}%"


def format_count(value: object, lang: str) -> str:
    number = _as_number(value)
    if number is None:
        return "N/A"
    return format_grouped_int(number, lang)


def _format_cell(value: object, kind: str, lang: str) -> object:
    if kind == "vnd":
        return format_vnd(value, lang)
    if kind == "measure":
        return format_measure(value, lang)
    if kind == "confidence":
        return format_confidence(value)
    if kind == "id":
        number = _as_number(value)
        return "N/A" if number is None else str(int(number))
    if _is_blank(value):
        return "N/A"
    return value


def _side_of_bilingual(value: object, lang: str) -> object:
    """Take the VI or EN half of pipeline labels like 'Trung lập / Neutral'."""
    if not isinstance(value, str) or " / " not in value:
        return value
    pieces = []
    for chunk in value.split("; "):
        if " / " in chunk:
            vi_part, en_part = chunk.split(" / ", 1)
            pieces.append(en_part if lang == "en" else vi_part)
        else:
            pieces.append(chunk)
    return "; ".join(pieces)


def localize_frame(frame: pd.DataFrame, lang: str, *, show_raw: bool) -> pd.DataFrame:
    """Language-only view of the bilingual clean_frame. Excel download is unchanged."""
    out = pd.DataFrame()
    for source, vi_header, en_header, kind in _DISPLAY_COLUMNS:
        if source not in frame.columns:
            continue
        if kind == "vi" and lang != "vi":
            continue
        if kind == "en" and lang != "en":
            continue
        if kind == "raw_optional" and not show_raw:
            continue
        header = en_header if lang == "en" else vi_header
        series = frame[source]
        if kind == "split":
            series = series.map(lambda cell, code=lang: _side_of_bilingual(cell, code))
        fmt = _FORMAT_BY_SOURCE.get(source, "text")
        out[header] = [
            _format_cell(cell, fmt, lang) for cell in series.to_numpy()
        ]
    return out.reset_index(drop=True)


def paint_status(value: object) -> str:
    if value == STATUS_OK:
        return "background-color: #C6EFCE; color: #196F3D; font-weight: 700"
    if isinstance(value, str) and value in {"CẦN KIỂM TRA", "NEEDS REVIEW"}:
        return "background-color: #FCE4D6; color: #C65911; font-weight: 700"
    return ""


def _status_header(lang: str) -> str:
    return "Status" if lang == "en" else "Trạng thái"


def _id_header(lang: str) -> str:
    return "No." if lang == "en" else "STT"


def preview_display(frame: pd.DataFrame, lang: str) -> pd.DataFrame:
    """Language-only copy of a source table. Pipeline input is not modified."""
    id_header = _id_header(lang)
    out = pd.DataFrame()
    out[id_header] = [str(i) for i in range(1, len(frame) + 1)]
    used = {id_header}
    for name in frame.columns:
        mapped = _PREVIEW_HEADERS.get(str(name))
        if mapped is None:
            header = str(name)
        else:
            header = mapped[1] if lang == "en" else mapped[0]
        if header in used:
            header = str(name)
        used.add(header)
        out[header] = [_format_cell(cell, "text", lang) for cell in frame[name].to_numpy()]
    return out.reset_index(drop=True)


def _show_preview(frame: pd.DataFrame, lang: str, *, key: str) -> None:
    view = preview_display(frame, lang)
    id_header = _id_header(lang)
    st.dataframe(
        view,
        width="stretch",
        height=240,
        hide_index=True,
        key=key,
        column_config={id_header: st.column_config.Column(alignment="right")},
    )


def _show_counts(
    counts: pd.Series, lang: str, label_header: str, count_header: str, *, key: str
) -> None:
    id_header = _id_header(lang)
    display = pd.DataFrame(
        {
            id_header: [str(i) for i in range(1, len(counts) + 1)],
            label_header: [str(x) if not _is_blank(x) else "N/A" for x in counts.index],
            count_header: [format_count(value, lang) for value in counts.to_numpy()],
        }
    )
    st.dataframe(
        display,
        width="stretch",
        hide_index=True,
        key=key,
        column_config={
            id_header: st.column_config.Column(alignment="right"),
            count_header: st.column_config.Column(alignment="right"),
        },
    )


def _show_table(frame: pd.DataFrame, lang: str, *, height: int, key: str) -> None:
    status_header = _status_header(lang)
    config = {
        name: st.column_config.Column(alignment="right")
        for name in _NUMERIC_HEADERS[lang]
        if name in frame.columns
    }
    if status_header in frame.columns:
        styled = frame.style.map(paint_status, subset=[status_header])
        st.dataframe(
            styled,
            width="stretch",
            height=height,
            hide_index=True,
            key=key,
            column_config=config,
        )
    else:
        st.dataframe(
            frame,
            width="stretch",
            height=height,
            hide_index=True,
            key=key,
            column_config=config,
        )


def render_results(result: dict[str, object], text: dict[str, str]) -> None:
    stats: RunStats = result["stats"]  # type: ignore[assignment]
    frame: pd.DataFrame = result["frame"]  # type: ignore[assignment]
    if result.get("from_memory_cache"):
        st.info(text["memory_cache"])

    lang = st.session_state.get("lang_widget", "vi")
    columns = st.columns(3)
    for column, (label, value) in zip(
        columns,
        [
            (text["kpi_rows"], format_count(stats.rows, lang)),
            (text["kpi_ok"], format_count(stats.ok, lang)),
            (text["kpi_review"], format_count(stats.needs_review, lang)),
        ],
    ):
        column.metric(label, value)
    with st.expander(text["tech_details"]):
        tech = st.columns(3)
        for column, (label, value) in zip(
            tech,
            [
                (text["kpi_api"], format_count(stats.api_calls, lang)),
                (text["kpi_cache"], format_count(stats.cache_hits, lang)),
                (text["kpi_tokens"], format_count(stats.total_tokens, lang)),
            ],
        ):
            column.metric(label, value)

    st.download_button(
        text["download"],
        data=result["excel"],
        file_name="enriched.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download-enriched",
    )

    show_raw = st.toggle(text["show_raw"], key="show_raw_columns")
    view = localize_frame(frame, lang, show_raw=show_raw)
    status_header = _status_header(lang)
    category_header = "Category" if lang == "en" else "Danh mục"
    sentiment_header = "Sentiment" if lang == "en" else "Cảm xúc"
    other_label = "Other" if lang == "en" else "Khác"
    neutral_label = "Neutral" if lang == "en" else "Trung lập"

    clean_tab, summary_tab, review_tab = st.tabs(
        [text["clean_tab"], text["summary_tab"], text["review_tab"]]
    )
    with clean_tab:
        if view.empty:
            st.caption(text["empty"])
        else:
            _show_table(view, lang, height=460, key=f"clean-table-{lang}-{int(show_raw)}")
    with summary_tab:
        if view.empty:
            st.caption(text["empty"])
        else:
            left, right = st.columns(2)
            categories = view[category_header].fillna(other_label).value_counts()
            sentiment = view[sentiment_header].fillna(neutral_label).value_counts()
            count_header = "Count" if lang == "en" else "Số lượng"
            with left:
                st.markdown(f"**{text['categories']}**")
                st.bar_chart(categories, height=280, color="#1F4E79")
                _show_counts(
                    categories, lang, category_header, count_header, key=f"cat-counts-{lang}"
                )
            with right:
                st.markdown(f"**{text['sentiment']}**")
                st.bar_chart(sentiment, height=280, color="#1F4E79")
                _show_counts(
                    sentiment, lang, sentiment_header, count_header, key=f"sent-counts-{lang}"
                )
    with review_tab:
        if view.empty or status_header not in view.columns:
            st.caption(text["no_review"])
            return
        flagged = view.loc[view[status_header] != STATUS_OK].copy()
        if flagged.empty:
            st.success(text["no_review"])
            return
        _show_table(flagged, lang, height=360, key=f"review-table-{lang}-{int(show_raw)}")


def render_sidebar(text: dict[str, str], has_key: bool):
    st.radio(
        text["language"],
        options=["vi", "en"],
        format_func=lambda code: "Tiếng Việt" if code == "vi" else "English",
        key="lang_widget",
    )
    model = st.selectbox(text["model"], MODEL_CHOICES)
    remaining_box = st.empty()
    today_box = st.empty()
    if has_key:
        st.success(text["key_on"])
    else:
        st.warning(text["key_off"])
    return model, remaining_box, today_box


def paint_limits(text: dict[str, str], remaining_box: st.delta_generator.DeltaGenerator, today_box: st.delta_generator.DeltaGenerator) -> None:
    lang = st.session_state.get("lang_widget", "vi")
    remaining = max(0, MAX_SESSION_RUNS - int(st.session_state.live_runs))
    remaining_box.metric(
        text["remaining"],
        f"{format_count(remaining, lang)}/{format_count(MAX_SESSION_RUNS, lang)}",
    )
    used = int(daily_counter()["calls"])
    left = max(0, DAILY_API_CAP - used)
    today_box.metric(
        text["today"],
        f"{format_count(left, lang)}/{format_count(DAILY_API_CAP, lang)}",
    )


def main() -> None:
    lang = st.session_state.get("lang_widget", "vi")
    st.set_page_config(
        page_title=COPY[lang]["title"],
        page_icon="📋",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown(
        """
        <style>
        div[data-testid="stHorizontalBlock"] { flex-wrap: wrap; }
        div[data-testid="stMetric"] {
            background: #F4F7FB;
            border: 1px solid #D6E2EE;
            border-radius: 10px;
            padding: 0.55rem 0.75rem;
        }
        a.github-btn {
            display: inline-flex;
            align-items: center;
            gap: 0.45rem;
            margin: 0.15rem 0 0.75rem;
            padding: 0.4rem 0.9rem;
            border: 1px solid #D6E2EE;
            border-radius: 0.5rem;
            background: #F4F7FB;
            color: #102A43 !important;
            text-decoration: none !important;
            font-weight: 600;
            width: fit-content;
        }
        a.github-btn:hover {
            border-color: #1F4E79;
            color: #1F4E79 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    init_state()
    has_key = bool(apply_api_key())
    st.session_state.setdefault("lang_widget", "vi")
    text = COPY[st.session_state.lang_widget]
    with st.sidebar:
        model, remaining_box, today_box = render_sidebar(text, has_key)
        text = COPY[st.session_state.lang_widget]

    st.title(text["title"], anchor=False)
    st.caption(text["tagline"])
    mark = "data:image/svg+xml;utf8," + quote(GITHUB_MARK)
    st.html(
        "<a class='github-btn' href='"
        + REPO_URL
        + "' target='_blank' rel='noopener noreferrer' style='"
        "display:inline-flex;align-items:center;gap:0.45rem;"
        "padding:0.4rem 0.9rem;border:1px solid #D6E2EE;border-radius:0.5rem;"
        "background:#F4F7FB;color:#102A43;text-decoration:none;font-weight:600;"
        "font-family:sans-serif;width:fit-content;'>"
        f"<img src='{mark}' width='16' height='16' alt=''>"
        " GitHub</a>"
    )
    st.info(text["privacy"])

    mode = st.radio(
        text["source"],
        options=["sample", "upload"],
        format_func=lambda item: text[item],
        horizontal=True,
    )

    frame: pd.DataFrame | None = None
    digest = ""
    source_name = SAMPLE_PATH.name
    if mode == "sample":
        try:
            frame = load_table(SAMPLE_PATH)
            st.caption(text["sample_note"].format(n=len(frame)))
        except EnrichmentError as exc:
            show_error(exc)
            frame = None
    else:
        st.markdown(f"**{text['expected']}**")
        st.code(", ".join(REQUIRED_COLUMNS))
        if SAMPLE_PATH.is_file():
            st.download_button(
                text["template"],
                data=SAMPLE_PATH.read_bytes(),
                file_name="shop_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        uploaded = st.file_uploader(text["uploader"], type=["xlsx", "csv"])
        if uploaded is not None:
            source_name = uploaded.name
            try:
                data = uploaded.getvalue()
                frame, digest, original_rows = prepare_upload(uploaded.name, data)
                if original_rows > MAX_UPLOAD_ROWS:
                    st.warning(text["truncate"].format(n=original_rows))
            except EnrichmentError as exc:
                show_error(exc)
                frame = None

    if frame is not None:
        heading = text["preview_sample"] if mode == "sample" else text["preview_upload"]
        st.subheader(heading, anchor=False)
        st.caption(text["rows_in_file"].format(n=len(frame)))
        lang = st.session_state.get("lang_widget", "vi")
        _show_preview(frame.head(8), lang, key=f"preview-table-{lang}-{mode}")
        if mode == "upload" and has_key and len(frame):
            estimate = math.ceil(len(frame) / WEB_BATCH_SIZE)
            st.caption(text["estimate"].format(n=estimate))

    run_clicked = st.button(text["run"], type="primary", disabled=frame is None or frame.empty)
    if run_clicked and frame is not None:
        if mode == "sample":
            with st.spinner(text["running"]):
                slot = st.progress(0, text=text["progress"].format(done=0, total=1))
                try:
                    st.session_state.result = run_enrichment(
                        frame,
                        source_name,
                        model,
                        offline=True,
                        progress=slot,
                        progress_label=text["progress"],
                    )
                except EnrichmentError as exc:
                    show_error(exc)
                except Exception as exc:
                    show_error(EnrichmentError(text["unexpected"].format(exc=exc)))
        else:
            cache_key = f"{digest}:{model}"
            stored = st.session_state.upload_cache.get(cache_key)
            if stored is not None:
                st.session_state.result = replay_cached(stored)
            elif not has_key:
                st.warning(text["no_key_run"])
            elif int(st.session_state.live_runs) >= MAX_SESSION_RUNS:
                st.warning(text["session_limit"])
            else:
                estimate = math.ceil(len(frame) / WEB_BATCH_SIZE) if len(frame) else 0
                counter = daily_counter()
                if estimate and counter["calls"] + estimate > DAILY_API_CAP:
                    st.warning(text["daily_limit"])
                else:
                    with st.spinner(text["running"]):
                        slot = st.progress(0, text=text["progress"].format(done=0, total=max(estimate, 1)))
                        try:
                            result = run_enrichment(
                                frame,
                                source_name,
                                model,
                                offline=False,
                                progress=slot,
                                progress_label=text["progress"],
                                charge=counter,
                            )
                        except EnrichmentError as exc:
                            show_error(exc)
                            result = None
                        except Exception as exc:
                            show_error(EnrichmentError(text["unexpected"].format(exc=exc)))
                            result = None
                        if result is not None:
                            st.session_state.upload_cache[cache_key] = result
                            st.session_state.result = result

    paint_limits(text, remaining_box, today_box)
    if st.session_state.result is not None:
        render_results(st.session_state.result, text)

    st.divider()
    st.caption(f"{text['footer']} [{text['profile']}]({PROFILE_URL})")


if __name__ == "__main__":
    main()
