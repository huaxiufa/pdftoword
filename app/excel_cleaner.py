from __future__ import annotations

import csv
import io
import math
import re
import unicodedata
from pathlib import Path
from typing import Any

import pandas as pd

INVISIBLE_RE = re.compile(r"[\u0000-\u001f\u007f\u0080-\u009f\u200b\u200c\u200d\ufeff]")
EMAIL_RE = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
PLACEHOLDERS = {"n/a", "na", "null", "none", "nil", "tbd", "unknown", "-", "--", "—", "暂无", "无", "未知"}


def _text(value: Any) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return str(value)


def _clean_text(value: Any) -> Any:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    if isinstance(value, (int, float, bool)):
        return value
    text = unicodedata.normalize("NFKC", str(value))
    text = text.replace("\xa0", " ")
    text = INVISIBLE_RE.sub("", text)
    text = re.sub(r"[ \t\r\n]+", " ", text).strip()
    if text.casefold() in PLACEHOLDERS:
        return ""
    return text


def _header_name(value: Any, index: int) -> str:
    text = _clean_text(value)
    text = _text(text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or f"Column_{index + 1}"


def _unique_headers(headers: list[str]) -> list[str]:
    counts: dict[str, int] = {}
    out = []
    for header in headers:
        counts[header] = counts.get(header, 0) + 1
        out.append(header if counts[header] == 1 else f"{header}_{counts[header]}")
    return out


def _looks_email_column(name: str) -> bool:
    return bool(re.search(r"email|e-mail|邮箱|邮件", name, re.I))


def _looks_phone_column(name: str) -> bool:
    return bool(re.search(r"phone|mobile|tel|telephone|电话|手机|手机号|联系电话", name, re.I))


def _looks_date_column(name: str) -> bool:
    return bool(re.search(r"date|日期|时间|生日|出生|created|updated|time", name, re.I))


def _normalize_email(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    value = value.strip().lower()
    return value


def _normalize_phone(value: Any) -> Any:
    if value is None or value == "":
        return value
    text = str(value).strip()
    if text.startswith("+"):
        digits = "+" + re.sub(r"\D", "", text[1:])
    else:
        digits = re.sub(r"\D", "", text)
    return digits


def _safe_csv_read(path: Path) -> pd.DataFrame:
    raw = path.read_bytes()
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "gb18030", "big5", "cp1252", "latin1"):
        try:
            text = raw.decode(encoding)
            sample = text[:10000]
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
                sep = dialect.delimiter
            except Exception:
                sep = ","
            return pd.read_csv(io.StringIO(text), sep=sep, dtype=object, keep_default_na=False)
        except Exception as exc:
            last_error = exc
    raise ValueError(f"CSV 无法读取：{last_error}")


def load_input(path: Path) -> tuple[pd.DataFrame, str]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return _safe_csv_read(path), "CSV"
    if suffix in {".xlsx", ".xlsm"}:
        return pd.read_excel(path, sheet_name=0, dtype=object), "XLSX"
    raise ValueError("只支持 .xlsx、.xlsm、.csv 文件")


def _key_for_row(row: pd.Series, columns: list[str]) -> str:
    parts = []
    for column in columns:
        value = row.get(column, "")
        text = _text(value).strip().casefold()
        parts.append(text)
    return "\x1f".join(parts)


def clean_dataframe(df: pd.DataFrame, options: dict[str, Any], progress=None) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    df = df.copy()
    original_rows, original_cols = len(df), len(df.columns)
    df.columns = _unique_headers([_header_name(v, i) for i, v in enumerate(df.columns)])

    issues: list[dict[str, Any]] = []
    changed_cells = 0
    placeholder_cells = 0

    if progress:
        progress(10, "正在清理空白字符和隐藏字符…")

    for column in df.columns:
        before = df[column].copy()
        df[column] = df[column].map(_clean_text)
        changed_cells += int((before.astype(str) != df[column].astype(str)).sum())
        placeholder_cells += int(((before.map(_text).str.strip().str.casefold()).isin(PLACEHOLDERS)).sum())

        if _looks_email_column(column):
            df[column] = df[column].map(_normalize_email)
        elif _looks_phone_column(column):
            df[column] = df[column].map(_normalize_phone)

    # Remove fully blank rows after normalization.
    blank_mask = df.apply(lambda row: all(_text(v).strip() == "" for v in row), axis=1)
    blank_rows = int(blank_mask.sum())
    for idx in df.index[blank_mask]:
        issues.append({"row": int(idx) + 2, "column": "", "type": "blank_row", "detail": "整行为空，已删除"})
    if options.get("remove_blank_rows", True):
        df = df.loc[~blank_mask].reset_index(drop=True)

    if progress:
        progress(30, "正在检查邮箱、手机号和日期…")

    invalid_emails = 0
    for column in df.columns:
        if _looks_email_column(column):
            for idx, value in df[column].items():
                text = _text(value).strip()
                if text and not EMAIL_RE.match(text):
                    invalid_emails += 1
                    issues.append({"row": int(idx) + 2, "column": column, "type": "invalid_email", "detail": text})
        if _looks_date_column(column):
            # Only validate obvious date-like values; do not rewrite arbitrary text columns.
            for idx, value in df[column].items():
                text = _text(value).strip()
                if not text:
                    continue
                parsed = pd.to_datetime(value, errors="coerce")
                if pd.isna(parsed) and re.search(r"\d", text):
                    issues.append({"row": int(idx) + 2, "column": column, "type": "invalid_date", "detail": text})

    duplicates_removed = 0
    dedupe_columns = options.get("dedupe_columns") or []
    dedupe_columns = [c for c in dedupe_columns if c in df.columns]
    if not dedupe_columns and options.get("smart_dedupe", True):
        candidates = [c for c in df.columns if _looks_email_column(c) or _looks_phone_column(c)]
        dedupe_columns = candidates[:3]

    if progress:
        progress(50, "正在检查重复记录…")

    if options.get("remove_duplicates", True):
        before_count = len(df)
        if dedupe_columns:
            # Normalize keys once more so spacing/case differences do not hide duplicates.
            keys = df.apply(lambda row: _key_for_row(row, dedupe_columns), axis=1)
            duplicate_mask = keys.duplicated(keep="first") & keys.ne("")
        else:
            normalized = df.applymap(lambda v: _text(v).strip().casefold())
            duplicate_mask = normalized.duplicated(keep="first")
        for idx in df.index[duplicate_mask]:
            cols = ", ".join(dedupe_columns) if dedupe_columns else "全部列"
            issues.append({"row": int(idx) + 2, "column": cols, "type": "duplicate", "detail": "重复记录，已删除"})
        df = df.loc[~duplicate_mask].reset_index(drop=True)
        duplicates_removed = before_count - len(df)

    if progress:
        progress(70, "正在检查缺失字段…")

    for column in df.columns:
        missing = df[column].map(lambda v: _text(v).strip() == "")
        count = int(missing.sum())
        if count:
            for idx in df.index[missing].tolist()[:500]:
                issues.append({"row": int(idx) + 2, "column": column, "type": "blank_cell", "detail": "空值"})

    if progress:
        progress(85, "正在生成清洗报告…")

    issues_df = pd.DataFrame(issues, columns=["row", "column", "type", "detail"])
    summary = {
        "original_rows": original_rows,
        "cleaned_rows": len(df),
        "removed_blank_rows": blank_rows if options.get("remove_blank_rows", True) else 0,
        "removed_duplicates": duplicates_removed,
        "columns": original_cols,
        "changed_cells": changed_cells,
        "placeholder_cells": placeholder_cells,
        "invalid_emails": invalid_emails,
        "issues": len(issues_df),
        "dedupe_keys": ", ".join(dedupe_columns) if dedupe_columns else "全部列",
    }
    return df, issues_df, summary


def write_result(cleaned: pd.DataFrame, issues: pd.DataFrame, summary: dict[str, Any], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    summary_df = pd.DataFrame([{"metric": k, "value": v} for k, v in summary.items()])
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cleaned.to_excel(writer, index=False, sheet_name="Cleaned")
        issues.to_excel(writer, index=False, sheet_name="Issues")
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        for sheet in ("Cleaned", "Issues", "Summary"):
            ws = writer.book[sheet]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                max_len = 0
                for cell in col[:200]:
                    value = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(value))
                ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 10), 42)


def clean_file(source: Path, output: Path, options: dict[str, Any], progress=None) -> dict[str, Any]:
    if progress:
        progress(5, "正在读取文件…")
    df, source_type = load_input(source)
    if len(df.columns) == 0:
        raise ValueError("文件没有可用的表头或数据列")
    cleaned, issues, summary = clean_dataframe(df, options, progress)
    summary["source_type"] = source_type
    write_result(cleaned, issues, summary, output)
    if progress:
        progress(100, "清洗完成")
    return summary
