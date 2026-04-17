from __future__ import annotations

import io
import os
import re
import threading
from datetime import datetime
from typing import Any

import pandas as pd
from flask import Flask, jsonify, render_template, request

app = Flask(__name__)
_state_lock = threading.Lock()


CANON_MAIN_COLS = ["province", "distributor", "store", "category", "month", "value"]
CANON_STOCK_COLS = ["province", "distributor", "category", "weight", "boxes"]
CANON_SCAN_COLS = ["province", "month", "scan_value", "outbound_value"]
CANON_OUT_COLS = [
    "province",
    "distributor",
    "store",
    "category",
    "weight",
    "product",
    "year",
    "month",
    "day",
    "month_label",
    "date",
    "boxes",
]
CANON_NEWCUST_COLS = ["province", "distributor", "store", "ym", "year", "month", "new_customers"]


def _empty_df(columns: list[str]) -> pd.DataFrame:
    return pd.DataFrame(columns=columns)


APP_STATE: dict[str, Any] = {
    "loaded": False,
    "file_name": "",
    "updated_at": None,
    "main": _empty_df(CANON_MAIN_COLS),
    "stock": _empty_df(CANON_STOCK_COLS),
    "scan": _empty_df(CANON_SCAN_COLS),
    "outbound": _empty_df(CANON_OUT_COLS),
    "newcust": _empty_df(CANON_NEWCUST_COLS),
    "meta": {},
}


PROVINCE_ALIASES = ["省区", "省区名称", "品牌省区名称", "省份", "区域", "province", "region"]
DIST_ALIASES = [
    "经销商名称",
    "经销商",
    "客户简称",
    "客户名称",
    "客户",
    "购货单位",
    "distributor",
    "dealer",
]
OUT_DIST_ALIASES = ["客户简称", "经销商简称", "购货单位", "经销商名称", "经销商", "客户名称", "客户"]
STOCK_DIST_ALIASES = ["客户简称", "经销商简称", "经销商名称", "经销商", "客户名称", "客户"]
STORE_ALIASES = ["门店名称", "门店", "发往门店", "终端门店", "store", "shop"]
CATEGORY_ALIASES = ["产品大类", "大分类", "大类", "归类", "品类", "category", "class"]
WEIGHT_ALIASES = ["重量", "产品小类", "规格重量", "净含量", "weight"]
PRODUCT_ALIASES = ["出库产品", "产品名称", "产品", "product"]
DATE_ALIASES = ["日期", "出库时间", "时间", "date", "period"]
YEAR_ALIASES = ["年份", "年", "year"]
MONTH_ALIASES = ["月份", "月", "month"]
DAY_ALIASES = ["日", "day"]
NEWCUST_ALIASES = ["新客", "新客数", "新客数量", "新增客户", "new"]

VALUE_ALIASES = [
    "发货金额",
    "业绩",
    "出库数量",
    "数量(箱)",
    "箱数",
    "销量",
    "金额",
    "value",
    "amount",
    "qty",
]
OUTBOUND_BOX_ALIASES = ["箱", "箱数", "数量(箱)", "发货箱数", "出库箱数", "shipment", "outbound"]
STOCK_BOX_ALIASES = ["箱数", "库存箱数", "库存数量", "库存数量(听)", "库存量", "stock", "inventory"]
SCAN_ALIASES = ["扫码", "扫码(箱)", "扫码箱数", "扫码量", "scan"]
OUTBOUND_ALIASES = ["出库", "出库(箱)", "发货", "发货箱数", "数量(箱)", "qty", "shipment"]


def _normalize_text(value: Any) -> str:
    s = str(value or "").strip().lower()
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"\s+", "", s)
    return s


def _clean_frame(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    return out


def _find_column(columns: list[str], aliases: list[str]) -> str | None:
    norm_map = {_normalize_text(c): c for c in columns}
    for alias in aliases:
        n = _normalize_text(alias)
        if n in norm_map:
            return norm_map[n]

    for col in columns:
        nc = _normalize_text(col)
        if any(_normalize_text(alias) in nc for alias in aliases):
            return col
    return None


def _safe_to_int_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0).astype(int)


def _safe_to_num_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0).astype(float)


def _month_sort_key(label: str) -> tuple[int, int]:
    if not isinstance(label, str):
        return (9_999, 9_999)
    m = re.match(r"^(20\d{2})-(\d{2})$", label)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    m2 = re.match(r"^M(\d{2})$", label)
    if m2:
        return (9_998, int(m2.group(1)))
    return (9_999, 9_999)


def _normalize_month_label(raw: Any) -> str:
    if pd.isna(raw):
        return "Unknown"
    if isinstance(raw, pd.Timestamp):
        return raw.strftime("%Y-%m")
    txt = str(raw).strip()
    if not txt:
        return "Unknown"

    m = re.search(r"(20\d{2})[-/年_. ]*(\d{1,2})", txt)
    if m:
        y = int(m.group(1))
        mm = int(m.group(2))
        if 1 <= mm <= 12:
            return f"{y}-{mm:02d}"

    m2 = re.search(r"(^|[^0-9])(\d{1,2})月", txt)
    if m2:
        mm = int(m2.group(2))
        if 1 <= mm <= 12:
            return f"M{mm:02d}"

    n = _normalize_text(txt)
    if re.fullmatch(r"\d{6}", n):
        y = int(n[:4])
        mm = int(n[4:6])
        if 2000 <= y <= 2100 and 1 <= mm <= 12:
            return f"{y}-{mm:02d}"

    dt = pd.to_datetime(txt, errors="coerce")
    if pd.notna(dt):
        return dt.strftime("%Y-%m")
    return txt[:20]


def _best_numeric_column(df: pd.DataFrame, excluded: set[str]) -> str | None:
    best_col = None
    best_score = -1
    for c in df.columns:
        if c in excluded:
            continue
        s = pd.to_numeric(df[c], errors="coerce")
        score = int(s.notna().sum())
        if score > best_score:
            best_score = score
            best_col = c
    return best_col if best_score > 0 else None


def _sorted_unique(series: pd.Series) -> list[str]:
    values = [str(v).strip() for v in series.fillna("Unknown").tolist()]
    values = [v for v in values if v]
    return sorted(set(values))


def _sheet_name_bonus(sheet_name: str, keywords: list[str]) -> int:
    n = _normalize_text(sheet_name)
    return 1 if any(_normalize_text(k) in n for k in keywords) else 0


def _extract_main_sheet(df: pd.DataFrame, sheet_name: str, sheet_index: int) -> tuple[pd.DataFrame, int]:
    if df is None or df.empty:
        return _empty_df(CANON_MAIN_COLS), 0

    frame = _clean_frame(df)
    cols = list(frame.columns)
    province_col = _find_column(cols, PROVINCE_ALIASES)
    dist_col = _find_column(cols, STOCK_DIST_ALIASES) or _find_column(cols, DIST_ALIASES)
    store_col = _find_column(cols, STORE_ALIASES)
    category_col = _find_column(cols, CATEGORY_ALIASES)
    date_col = _find_column(cols, DATE_ALIASES + MONTH_ALIASES + YEAR_ALIASES)
    metric_col = _find_column(cols, VALUE_ALIASES)
    month_cols = [c for c in cols if _normalize_month_label(c).startswith(("20", "M"))]

    score = 0
    score += 2 if province_col else 0
    score += 2 if dist_col else 0
    score += 1 if category_col else 0
    score += 2 if metric_col else 0
    score += 2 if len(month_cols) >= 3 else 1 if month_cols else 0
    score += _sheet_name_bonus(sheet_name, ["出库", "发货", "业绩", "销售", "shipment", "performance"])
    score += 1 if sheet_index in (0, 2, 3) else 0

    dim_cols = [c for c in [province_col, dist_col, store_col, category_col] if c is not None]
    if not dim_cols:
        return _empty_df(CANON_MAIN_COLS), 0

    try:
        if len(month_cols) >= 3:
            use_cols = list(dict.fromkeys(dim_cols + month_cols))
            long_df = frame[use_cols].melt(
                id_vars=dim_cols,
                value_vars=month_cols,
                var_name="month",
                value_name="value",
            )
        else:
            if metric_col is None:
                metric_col = _best_numeric_column(frame, set(dim_cols + ([date_col] if date_col else [])))
            if metric_col is None:
                return _empty_df(CANON_MAIN_COLS), 0
            keep_cols = list(dict.fromkeys(dim_cols + [metric_col] + ([date_col] if date_col else [])))
            long_df = frame[keep_cols].copy()
            long_df["month"] = long_df[date_col].map(_normalize_month_label) if date_col else "Total"
            long_df = long_df.rename(columns={metric_col: "value"})
    except Exception:
        return _empty_df(CANON_MAIN_COLS), 0

    rename_map: dict[str, str] = {}
    if province_col and province_col in long_df.columns:
        rename_map[province_col] = "province"
    if dist_col and dist_col in long_df.columns:
        rename_map[dist_col] = "distributor"
    if store_col and store_col in long_df.columns:
        rename_map[store_col] = "store"
    if category_col and category_col in long_df.columns:
        rename_map[category_col] = "category"

    long_df = long_df.rename(columns=rename_map)
    for c in ["province", "distributor", "store", "category"]:
        if c not in long_df.columns:
            long_df[c] = "Unknown"
        long_df[c] = long_df[c].fillna("Unknown").astype(str).str.strip()
        long_df.loc[long_df[c] == "", c] = "Unknown"

    long_df["month"] = long_df["month"].map(_normalize_month_label)
    long_df["value"] = _safe_to_num_series(long_df["value"])
    long_df = long_df[long_df["value"] != 0].copy()
    if long_df.empty:
        return _empty_df(CANON_MAIN_COLS), 0

    out = long_df[CANON_MAIN_COLS].copy()
    score += 1 if len(out) > 100 else 0
    return out, score


def _extract_outbound_sheet(df: pd.DataFrame, sheet_name: str, sheet_index: int) -> tuple[pd.DataFrame, int]:
    if df is None or df.empty:
        return _empty_df(CANON_OUT_COLS), 0

    frame = _clean_frame(df)
    cols = list(frame.columns)

    province_col = _find_column(cols, PROVINCE_ALIASES)
    dist_col = _find_column(cols, OUT_DIST_ALIASES) or _find_column(cols, DIST_ALIASES)
    store_col = _find_column(cols, STORE_ALIASES)
    category_col = _find_column(cols, CATEGORY_ALIASES)
    weight_col = _find_column(cols, WEIGHT_ALIASES)
    product_col = _find_column(cols, PRODUCT_ALIASES)

    year_col = _find_column(cols, YEAR_ALIASES)
    month_col = _find_column(cols, MONTH_ALIASES)
    day_col = _find_column(cols, DAY_ALIASES)
    date_col = _find_column(cols, DATE_ALIASES)
    box_col = _find_column(cols, OUTBOUND_BOX_ALIASES)

    score = 0
    score += 2 if box_col else 0
    score += 2 if province_col else 0
    score += 2 if dist_col else 0
    score += 1 if category_col else 0
    score += 1 if year_col else 0
    score += 1 if month_col else 0
    score += 1 if day_col else 0
    score += 1 if date_col else 0
    score += _sheet_name_bonus(sheet_name, ["出库", "发货", "shipment", "outbound", "sheet3"])
    score += 1 if sheet_index in (2, 3) else 0

    if box_col is None:
        box_col = _best_numeric_column(frame, excluded=set())
    if box_col is None:
        return _empty_df(CANON_OUT_COLS), 0

    out = pd.DataFrame()
    out["province"] = frame[province_col] if province_col else "Unknown"
    out["distributor"] = frame[dist_col] if dist_col else "Unknown"
    out["store"] = frame[store_col] if store_col else "Unknown"
    out["category"] = frame[category_col] if category_col else "Unknown"
    out["weight"] = frame[weight_col] if weight_col else "Unknown"
    out["product"] = frame[product_col] if product_col else "Unknown"
    out["boxes"] = _safe_to_num_series(frame[box_col])

    date_series = pd.to_datetime(frame[date_col], errors="coerce") if date_col else pd.Series(pd.NaT, index=frame.index)
    year_series = _safe_to_int_series(frame[year_col].astype(str).str.extract(r"(\d{4})")[0]) if year_col else pd.Series(0, index=frame.index)
    month_series = _safe_to_int_series(frame[month_col].astype(str).str.extract(r"(\d{1,2})")[0]) if month_col else pd.Series(0, index=frame.index)
    day_series = _safe_to_int_series(frame[day_col].astype(str).str.extract(r"(\d{1,2})")[0]) if day_col else pd.Series(0, index=frame.index)

    if year_series.eq(0).any():
        year_series = year_series.where(year_series != 0, date_series.dt.year.fillna(0).astype(int))
    if month_series.eq(0).any():
        month_series = month_series.where(month_series != 0, date_series.dt.month.fillna(0).astype(int))
    if day_series.eq(0).any():
        day_series = day_series.where(day_series != 0, date_series.dt.day.fillna(0).astype(int))

    year_series = year_series.where((year_series >= 2000) & (year_series <= 2100), 0)
    month_series = month_series.where((month_series >= 1) & (month_series <= 12), 0)
    day_series = day_series.where((day_series >= 1) & (day_series <= 31), 0)

    out["year"] = year_series.astype(int)
    out["month"] = month_series.astype(int)
    out["day"] = day_series.astype(int)
    out["month_label"] = out.apply(
        lambda r: f"{int(r['year'])}-{int(r['month']):02d}" if r["year"] > 0 and r["month"] > 0 else "Unknown", axis=1
    )

    normalized_date = date_series
    fill_mask = normalized_date.isna() & (out["year"] > 0) & (out["month"] > 0) & (out["day"] > 0)
    if fill_mask.any():
        normalized_date.loc[fill_mask] = pd.to_datetime(
            out.loc[fill_mask, ["year", "month", "day"]].rename(columns={"year": "year", "month": "month", "day": "day"}),
            errors="coerce",
        )
    out["date"] = normalized_date

    for c in ["province", "distributor", "store", "category", "product"]:
        out[c] = out[c].fillna("Unknown").astype(str).str.strip()
        out.loc[out[c] == "", c] = "Unknown"
    out["distributor"] = out["distributor"].astype(str).str.replace(r"\s+", "", regex=True)
    out["weight"] = out["weight"].fillna("Unknown").astype(str).str.strip()
    out.loc[out["weight"] == "", "weight"] = "Unknown"
    out["weight"] = out["weight"].map(_normalize_weight_value)

    out = out[out["boxes"] != 0].copy()
    out = out[out["month_label"] != "Unknown"].copy()
    if out.empty:
        return _empty_df(CANON_OUT_COLS), 0

    score += 1 if len(out) > 1000 else 0
    return out[CANON_OUT_COLS].reset_index(drop=True), score


def _extract_stock_sheet(df: pd.DataFrame, sheet_name: str, sheet_index: int) -> tuple[pd.DataFrame, int]:
    if df is None or df.empty:
        return _empty_df(CANON_STOCK_COLS), 0

    frame = _clean_frame(df)
    cols = list(frame.columns)
    province_col = _find_column(cols, PROVINCE_ALIASES)
    dist_col = _find_column(cols, STOCK_DIST_ALIASES) or _find_column(cols, DIST_ALIASES)
    category_col = _find_column(cols, CATEGORY_ALIASES)
    weight_col = _find_column(cols, WEIGHT_ALIASES)
    box_col = _find_column(cols, STOCK_BOX_ALIASES)

    score = 0
    score += 2 if box_col else 0
    score += 1 if province_col else 0
    score += 1 if dist_col else 0
    score += 1 if category_col else 0
    score += 1 if weight_col else 0
    score += _sheet_name_bonus(sheet_name, ["库存", "stock", "sheet2"])
    score += 1 if sheet_index == 1 else 0

    if box_col is None:
        return _empty_df(CANON_STOCK_COLS), 0

    out = pd.DataFrame()
    out["boxes"] = _safe_to_num_series(frame[box_col])
    out["province"] = frame[province_col] if province_col else "Unknown"
    out["distributor"] = frame[dist_col] if dist_col else "Unknown"
    out["category"] = frame[category_col] if category_col else "Unknown"
    out["weight"] = frame[weight_col] if weight_col else "Unknown"

    for c in ["province", "distributor", "category"]:
        out[c] = out[c].fillna("Unknown").astype(str).str.strip()
        out.loc[out[c] == "", c] = "Unknown"
    out["distributor"] = out["distributor"].astype(str).str.replace(r"\s+", "", regex=True)
    out["weight"] = out["weight"].fillna("Unknown").astype(str).str.strip()
    out.loc[out["weight"] == "", "weight"] = "Unknown"
    out["weight"] = out["weight"].map(_normalize_weight_value)

    out = out[out["boxes"] != 0].copy()
    if out.empty:
        return _empty_df(CANON_STOCK_COLS), 0
    return out[CANON_STOCK_COLS].reset_index(drop=True), score


def _extract_scan_sheet(df: pd.DataFrame, sheet_name: str, sheet_index: int) -> tuple[pd.DataFrame, int]:
    if df is None or df.empty:
        return _empty_df(CANON_SCAN_COLS), 0

    frame = _clean_frame(df)
    cols = list(frame.columns)
    province_col = _find_column(cols, PROVINCE_ALIASES)
    date_col = _find_column(cols, DATE_ALIASES + MONTH_ALIASES)
    scan_col = _find_column(cols, SCAN_ALIASES)
    outbound_col = _find_column(cols, OUTBOUND_ALIASES)

    score = 0
    score += 2 if scan_col else 0
    score += 1 if date_col else 0
    score += 1 if province_col else 0
    score += _sheet_name_bonus(sheet_name, ["扫码", "scan", "sheet6"])
    score += 1 if sheet_index == 5 else 0

    if scan_col is None:
        return _empty_df(CANON_SCAN_COLS), 0

    out = pd.DataFrame()
    out["province"] = frame[province_col] if province_col else "Unknown"
    out["month"] = frame[date_col].map(_normalize_month_label) if date_col else "Unknown"
    out["scan_value"] = _safe_to_num_series(frame[scan_col])
    out["outbound_value"] = _safe_to_num_series(frame[outbound_col]) if outbound_col else 0.0
    out["province"] = out["province"].fillna("Unknown").astype(str).str.strip()
    out.loc[out["province"] == "", "province"] = "Unknown"
    out = out[(out["scan_value"] != 0) | (out["outbound_value"] != 0)].copy()
    if out.empty:
        return _empty_df(CANON_SCAN_COLS), 0
    return out[CANON_SCAN_COLS].reset_index(drop=True), score


def _extract_newcust_sheet(df: pd.DataFrame, sheet_name: str, sheet_index: int) -> tuple[pd.DataFrame, int]:
    if df is None or df.empty:
        return _empty_df(CANON_NEWCUST_COLS), 0

    frame = _clean_frame(df)
    cols = list(frame.columns)
    province_col = _find_column(cols, PROVINCE_ALIASES)
    dist_col = _find_column(cols, OUT_DIST_ALIASES) or _find_column(cols, DIST_ALIASES)
    store_col = _find_column(cols, STORE_ALIASES)
    value_col = _find_column(cols, NEWCUST_ALIASES)
    year_col = _find_column(cols, YEAR_ALIASES)
    month_col = _find_column(cols, MONTH_ALIASES)
    date_col = _find_column(cols, DATE_ALIASES)
    ym_col = _find_column(cols, ["年月", "年月份", "统计月份", "month"])

    score = 0
    score += 2 if value_col else 0
    score += 1 if province_col else 0
    score += 1 if dist_col else 0
    score += 1 if store_col else 0
    score += 1 if (date_col or ym_col or (year_col and month_col)) else 0
    score += _sheet_name_bonus(sheet_name, ["新客", "new", "sheet5", "sheet8"])

    if value_col is None and frame.shape[1] > 7:
        # Fixed-layout fallback from the original base template:
        # A省区, E门店, F日期, G新客数, J客户名称
        province_col = frame.columns[0] if province_col is None and frame.shape[1] > 0 else province_col
        store_col = frame.columns[4] if store_col is None and frame.shape[1] > 4 else store_col
        date_col = frame.columns[5] if date_col is None and frame.shape[1] > 5 else date_col
        value_col = frame.columns[6] if value_col is None and frame.shape[1] > 6 else value_col
        dist_col = frame.columns[9] if dist_col is None and frame.shape[1] > 9 else dist_col
        score += 1

    if value_col is None:
        return _empty_df(CANON_NEWCUST_COLS), 0

    out = pd.DataFrame()
    out["province"] = frame[province_col] if province_col is not None else "Unknown"
    out["distributor"] = frame[dist_col] if dist_col is not None else "Unknown"
    out["store"] = frame[store_col] if store_col is not None else "Unknown"
    out["new_customers"] = _safe_to_num_series(frame[value_col])

    if year_col is not None and month_col is not None:
        yy = _safe_to_int_series(frame[year_col].astype(str).str.extract(r"(\d{4})")[0])
        mm = _safe_to_int_series(frame[month_col].astype(str).str.extract(r"(\d{1,2})")[0])
    else:
        yy = pd.Series(0, index=frame.index)
        mm = pd.Series(0, index=frame.index)

    if ym_col is not None:
        ym_raw = frame[ym_col].astype(str)
        ym_yyyy_mm = ym_raw.str.extract(r"(20\d{2})\D{0,3}(\d{1,2})")
        yy2 = _safe_to_int_series(ym_yyyy_mm[0])
        mm2 = _safe_to_int_series(ym_yyyy_mm[1])
        yyyymm = _safe_to_int_series(ym_raw.str.extract(r"(\d{6})")[0])
        yy3 = (yyyymm // 100).astype(int)
        mm3 = (yyyymm % 100).astype(int)
        yy = yy.where(yy != 0, yy2).where(yy != 0, yy3)
        mm = mm.where(mm != 0, mm2).where(mm != 0, mm3)

    if date_col is not None:
        dt = pd.to_datetime(frame[date_col], errors="coerce")
        yy = yy.where(yy != 0, dt.dt.year.fillna(0).astype(int))
        mm = mm.where(mm != 0, dt.dt.month.fillna(0).astype(int))

    yy = yy.where((yy >= 2000) & (yy <= 2100), 0)
    mm = mm.where((mm >= 1) & (mm <= 12), 0)
    out["year"] = yy.astype(int)
    out["month"] = mm.astype(int)
    out["ym"] = (out["year"] * 100 + out["month"]).astype(int)

    for c in ["province", "distributor", "store"]:
        out[c] = out[c].fillna("Unknown").astype(str).str.strip()
        out.loc[out[c] == "", c] = "Unknown"
    out["distributor"] = out["distributor"].astype(str).str.replace(r"\s+", "", regex=True)

    out = out[out["new_customers"] != 0].copy()
    out = out[out["ym"].between(200001, 209912)].copy()
    if out.empty:
        return _empty_df(CANON_NEWCUST_COLS), 0
    return out[CANON_NEWCUST_COLS].reset_index(drop=True), score


def _read_uploaded_sheets(file_bytes: bytes, file_name: str) -> list[tuple[str, pd.DataFrame]]:
    lower_name = (file_name or "").lower()
    if lower_name.endswith(".csv"):
        for enc in ["utf-8-sig", "gb18030", "gbk", "latin-1"]:
            try:
                df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
                return [("CSV", df)]
            except Exception:
                continue
        raise ValueError("CSV 无法识别编码，请尝试 UTF-8 或 GBK。")

    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    out: list[tuple[str, pd.DataFrame]] = []
    for sheet_name in xl.sheet_names:
        try:
            out.append((str(sheet_name), xl.parse(sheet_name)))
        except Exception:
            continue
    if not out:
        raise ValueError("Excel 中未读取到可用工作表。")
    return out


def _parse_uploaded_dataset(file_bytes: bytes, file_name: str) -> dict[str, Any]:
    sheets = _read_uploaded_sheets(file_bytes, file_name)

    main_candidates: list[tuple[pd.DataFrame, int, str]] = []
    stock_candidates: list[tuple[pd.DataFrame, int, str]] = []
    scan_candidates: list[tuple[pd.DataFrame, int, str]] = []
    outbound_candidates: list[tuple[pd.DataFrame, int, str]] = []
    newcust_candidates: list[tuple[pd.DataFrame, int, str]] = []

    for idx, (sheet_name, df) in enumerate(sheets):
        main_df, main_score = _extract_main_sheet(df, sheet_name, idx)
        if main_score > 0 and not main_df.empty:
            main_candidates.append((main_df, main_score, sheet_name))

        stock_df, stock_score = _extract_stock_sheet(df, sheet_name, idx)
        if stock_score > 0 and not stock_df.empty:
            stock_candidates.append((stock_df, stock_score, sheet_name))

        scan_df, scan_score = _extract_scan_sheet(df, sheet_name, idx)
        if scan_score > 0 and not scan_df.empty:
            scan_candidates.append((scan_df, scan_score, sheet_name))

        out_df, out_score = _extract_outbound_sheet(df, sheet_name, idx)
        if out_score > 0 and not out_df.empty:
            outbound_candidates.append((out_df, out_score, sheet_name))

        nc_df, nc_score = _extract_newcust_sheet(df, sheet_name, idx)
        if nc_score > 0 and not nc_df.empty:
            newcust_candidates.append((nc_df, nc_score, sheet_name))

    if not outbound_candidates and not main_candidates:
        raise ValueError("未识别到可分析的数据，请确认底表包含出库/发货或省区经销商指标。")

    main_df = _empty_df(CANON_MAIN_COLS)
    stock_df = _empty_df(CANON_STOCK_COLS)
    scan_df = _empty_df(CANON_SCAN_COLS)
    out_df = _empty_df(CANON_OUT_COLS)
    newcust_df = _empty_df(CANON_NEWCUST_COLS)

    selected_main_sheet = ""
    selected_stock_sheet = ""
    selected_scan_sheet = ""
    selected_out_sheet = ""
    selected_newcust_sheet = ""

    if outbound_candidates:
        outbound_candidates.sort(key=lambda x: (x[1], len(x[0])), reverse=True)
        out_df, _, selected_out_sheet = outbound_candidates[0]
        main_df = out_df[["province", "distributor", "store", "category", "month_label", "boxes"]].rename(
            columns={"month_label": "month", "boxes": "value"}
        )
        selected_main_sheet = selected_out_sheet
    elif main_candidates:
        main_candidates.sort(key=lambda x: (x[1], len(x[0])), reverse=True)
        main_df, _, selected_main_sheet = main_candidates[0]

    if stock_candidates:
        stock_candidates.sort(key=lambda x: (x[1], len(x[0])), reverse=True)
        stock_df, _, selected_stock_sheet = stock_candidates[0]

    if scan_candidates:
        scan_candidates.sort(key=lambda x: (x[1], len(x[0])), reverse=True)
        scan_df, _, selected_scan_sheet = scan_candidates[0]

    if newcust_candidates:
        newcust_candidates.sort(key=lambda x: (x[1], len(x[0])), reverse=True)
        newcust_df, _, selected_newcust_sheet = newcust_candidates[0]

    return {
        "main": main_df.reset_index(drop=True),
        "stock": stock_df.reset_index(drop=True),
        "scan": scan_df.reset_index(drop=True),
        "outbound": out_df.reset_index(drop=True),
        "newcust": newcust_df.reset_index(drop=True),
        "meta": {
            "sheet_names": [name for name, _ in sheets],
            "main_sheet": selected_main_sheet,
            "stock_sheet": selected_stock_sheet,
            "scan_sheet": selected_scan_sheet,
            "outbound_sheet": selected_out_sheet,
            "newcust_sheet": selected_newcust_sheet,
        },
    }


def _filter_main(df: pd.DataFrame, province: str, distributor: str, category: str) -> pd.DataFrame:
    out = df
    if province and province != "全部":
        out = out[out["province"] == province]
    if distributor and distributor != "全部":
        out = out[out["distributor"] == distributor]
    if category and category != "全部":
        out = out[out["category"] == category]
    return out


def _safe_round(value: float, digits: int = 2) -> float:
    try:
        return round(float(value), digits)
    except Exception:
        return 0.0


def _build_dashboard_payload(
    dataset: dict[str, Any],
    province: str = "全部",
    distributor: str = "全部",
    category: str = "全部",
) -> dict[str, Any]:
    main_df = dataset["main"]
    stock_df = dataset["stock"]
    scan_df = dataset["scan"]

    filtered_main = _filter_main(main_df, province, distributor, category)
    filtered_stock = _filter_main(stock_df, province, distributor, category)

    filtered_scan = scan_df
    if province and province != "全部":
        filtered_scan = filtered_scan[filtered_scan["province"] == province]

    total_value = float(filtered_main["value"].sum()) if not filtered_main.empty else 0.0
    store_count = int(filtered_main["store"].nunique()) if not filtered_main.empty else 0
    dist_count = int(filtered_main["distributor"].nunique()) if not filtered_main.empty else 0
    prov_count = int(filtered_main["province"].nunique()) if not filtered_main.empty else 0
    stock_boxes = float(filtered_stock["boxes"].sum()) if not filtered_stock.empty else 0.0

    trend_df = pd.DataFrame(columns=["month", "value"])
    if not filtered_main.empty:
        trend_df = filtered_main.groupby("month", as_index=False)["value"].sum()
        trend_df["sort_key"] = trend_df["month"].map(_month_sort_key)
        trend_df = trend_df.sort_values("sort_key").drop(columns=["sort_key"])

    mom_rate = None
    if len(trend_df) >= 2:
        prev = float(trend_df.iloc[-2]["value"])
        curr = float(trend_df.iloc[-1]["value"])
        if abs(prev) > 1e-9:
            mom_rate = (curr - prev) / abs(prev)

    scan_rate = None
    if not filtered_scan.empty:
        scan_total = float(filtered_scan["scan_value"].sum())
        out_total = float(filtered_scan["outbound_value"].sum())
        if out_total > 0:
            scan_rate = scan_total / out_total

    province_rank = pd.DataFrame(columns=["province", "value"])
    category_share = pd.DataFrame(columns=["category", "value"])
    dist_rank = pd.DataFrame(columns=["distributor", "value"])

    if not filtered_main.empty:
        province_rank = (
            filtered_main.groupby("province", as_index=False)["value"].sum().sort_values("value", ascending=False).head(10)
        )
        category_share = (
            filtered_main.groupby("category", as_index=False)["value"].sum().sort_values("value", ascending=False).head(8)
        )
        dist_rank = (
            filtered_main.groupby("distributor", as_index=False)["value"].sum().sort_values("value", ascending=False).head(10)
        )

    board = pd.DataFrame(columns=["province", "distributor", "value"])
    if not filtered_main.empty:
        board = (
            filtered_main.groupby(["province", "distributor"], as_index=False)["value"]
            .sum()
            .sort_values("value", ascending=False)
            .head(12)
            .reset_index(drop=True)
        )

    table_rows: list[dict[str, Any]] = []
    for i, row in board.iterrows():
        table_rows.append(
            {
                "rank": int(i + 1),
                "province": str(row["province"]),
                "distributor": str(row["distributor"]),
                "value": _safe_round(float(row["value"]), 2),
            }
        )

    return {
        "meta": {
            "file_name": dataset.get("file_name", ""),
            "updated_at": dataset.get("updated_at", ""),
            "main_sheet": dataset.get("meta", {}).get("main_sheet", ""),
            "stock_sheet": dataset.get("meta", {}).get("stock_sheet", ""),
            "scan_sheet": dataset.get("meta", {}).get("scan_sheet", ""),
            "outbound_sheet": dataset.get("meta", {}).get("outbound_sheet", ""),
        },
        "filters": {
            "province_options": ["全部"] + _sorted_unique(main_df["province"]) if not main_df.empty else ["全部"],
            "distributor_options": ["全部"] + _sorted_unique(main_df["distributor"]) if not main_df.empty else ["全部"],
            "category_options": ["全部"] + _sorted_unique(main_df["category"]) if not main_df.empty else ["全部"],
            "selected": {
                "province": province or "全部",
                "distributor": distributor or "全部",
                "category": category or "全部",
            },
        },
        "kpis": {
            "total_value": _safe_round(total_value, 2),
            "store_count": store_count,
            "distributor_count": dist_count,
            "province_count": prov_count,
            "inventory_boxes": _safe_round(stock_boxes, 2),
            "mom_rate": _safe_round(mom_rate * 100.0, 2) if mom_rate is not None else None,
            "scan_rate": _safe_round(scan_rate * 100.0, 2) if scan_rate is not None else None,
        },
        "trend": {
            "labels": trend_df["month"].tolist() if not trend_df.empty else [],
            "values": [_safe_round(v, 2) for v in trend_df["value"].tolist()] if not trend_df.empty else [],
        },
        "province_rank": {
            "labels": province_rank["province"].tolist() if not province_rank.empty else [],
            "values": [_safe_round(v, 2) for v in province_rank["value"].tolist()] if not province_rank.empty else [],
        },
        "category_share": {
            "labels": category_share["category"].tolist() if not category_share.empty else [],
            "values": [_safe_round(v, 2) for v in category_share["value"].tolist()] if not category_share.empty else [],
        },
        "distributor_rank": {
            "labels": dist_rank["distributor"].tolist() if not dist_rank.empty else [],
            "values": [_safe_round(v, 2) for v in dist_rank["value"].tolist()] if not dist_rank.empty else [],
        },
        "table_rows": table_rows,
    }


def _parse_year_filter(value: str) -> int | None:
    if not value or value == "全部":
        return None
    m = re.search(r"(\d{4})", str(value))
    if not m:
        return None
    y = int(m.group(1))
    return y if 2000 <= y <= 2100 else None


def _parse_month_filter(value: str) -> int | None:
    if not value or value == "全部":
        return None
    m = re.search(r"(\d{1,2})", str(value))
    if not m:
        return None
    mm = int(m.group(1))
    return mm if 1 <= mm <= 12 else None


def _trend_type(values: list[float]) -> str:
    if len(values) < 2:
        return "样本不足"
    prev = float(values[-2])
    curr = float(values[-1])
    if prev <= 0 and curr > 0:
        return "新增长"
    if prev <= 0 and curr <= 0:
        return "平稳"
    ratio = curr / prev if prev else 1.0
    if ratio >= 1.15:
        return "上升"
    if ratio <= 0.85:
        return "下降"
    return "平稳"


def _weight_sort_key(v: str) -> tuple[int, float, str]:
    s = str(v or "").strip()
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    if m:
        return (0, float(m.group(1)), s)
    return (1, 0.0, s)


def _normalize_weight_value(v: Any) -> str:
    s = str(v or "").strip()
    if not s:
        return "Unknown"
    try:
        n = float(s)
        if abs(n - round(n)) < 1e-9:
            return str(int(round(n)))
        return f"{n:.2f}".rstrip("0").rstrip(".")
    except Exception:
        return s


def _build_outbound_payload(
    dataset: dict[str, Any],
    province: str = "全部",
    distributor: str = "全部",
    category: str = "全部",
    weight: str = "全部",
    year: str = "全部",
    month: str = "全部",
) -> dict[str, Any]:
    out_df = dataset["outbound"]
    stock_df = dataset.get("stock", _empty_df(CANON_STOCK_COLS))
    newcust_df = dataset.get("newcust", _empty_df(CANON_NEWCUST_COLS))

    if out_df.empty:
        return {
            "meta": {"file_name": dataset.get("file_name", ""), "updated_at": dataset.get("updated_at", "")},
            "filters": {
                "province_options": ["全部"],
                "distributor_options": ["全部"],
                "category_options": ["全部"],
                "weight_options": ["全部"],
                "year_options": ["全部"],
                "month_options": ["全部"],
                "selected": {
                    "province": province,
                    "distributor": distributor,
                    "category": category,
                    "weight": weight,
                    "year": year,
                    "month": month,
                },
            },
            "kpis": {
                "total_boxes": 0.0,
                "monthly_boxes": 0.0,
                "latest_day_boxes": 0.0,
                "record_count": 0,
                "store_count": 0,
                "mom_rate": None,
                "selected_month_label": "全部",
            },
            "monthly_trend": {"labels": [], "values": []},
            "daily_trend": {"labels": [], "values": []},
            "drill": {
                "level": "province",
                "path": {"province": None, "distributor": None},
                "month_columns": [],
                "metrics_columns": [],
                "rows": [],
            },
        }

    year_value = _parse_year_filter(year)
    month_value = _parse_month_filter(month)

    def _norm_key(v: Any) -> str:
        return re.sub(r"\s+", "", str(v or "").strip())

    base_for_dist = out_df if province == "全部" else out_df[out_df["province"] == province]
    dist_options = ["全部"] + _sorted_unique(base_for_dist["distributor"]) if not base_for_dist.empty else ["全部"]

    base_for_cat = base_for_dist if distributor == "全部" else base_for_dist[base_for_dist["distributor"] == distributor]
    cat_options = ["全部"] + _sorted_unique(base_for_cat["category"]) if not base_for_cat.empty else ["全部"]

    base_for_weight = base_for_cat.copy()
    weight_options = ["全部"] + sorted(_sorted_unique(base_for_weight["weight"]), key=_weight_sort_key) if not base_for_weight.empty else ["全部"]

    base_for_month = base_for_weight if weight == "全部" else base_for_weight[base_for_weight["weight"] == weight]
    if year_value is not None:
        base_for_month = base_for_month[base_for_month["year"] == year_value]
    month_options_vals = sorted(set(int(x) for x in base_for_month["month"].tolist() if int(x) > 0))
    month_options = ["全部"] + [f"{m:02d}" for m in month_options_vals]

    filtered = out_df.copy()
    if province != "全部":
        filtered = filtered[filtered["province"] == province]
    if distributor != "全部":
        filtered = filtered[filtered["distributor"] == distributor]
    if category != "全部":
        filtered = filtered[filtered["category"] == category]
    if weight != "全部":
        filtered = filtered[filtered["weight"] == weight]
    if year_value is not None:
        filtered = filtered[filtered["year"] == year_value]

    monthly_df = pd.DataFrame(columns=["month_label", "boxes"])
    if not filtered.empty:
        monthly_df = filtered.groupby("month_label", as_index=False)["boxes"].sum()
        monthly_df["sort_key"] = monthly_df["month_label"].map(_month_sort_key)
        monthly_df = monthly_df.sort_values("sort_key").drop(columns=["sort_key"])

    selected_month_label = "全部"
    selected_year = year_value
    selected_month = month_value
    if selected_year is not None and selected_month is not None:
        selected_month_label = f"{selected_year}-{selected_month:02d}"
    elif not monthly_df.empty:
        selected_month_label = str(monthly_df.iloc[-1]["month_label"])
        m = re.match(r"^(20\d{2})-(\d{2})$", selected_month_label)
        if m:
            selected_year = int(m.group(1))
            selected_month = int(m.group(2))

    selected_month_df = filtered.iloc[0:0].copy()
    if selected_year is not None and selected_month is not None:
        selected_month_df = filtered[(filtered["year"] == selected_year) & (filtered["month"] == selected_month)].copy()

    daily_df = pd.DataFrame(columns=["day", "boxes"])
    if not selected_month_df.empty:
        daily_df = (
            selected_month_df[selected_month_df["day"] > 0]
            .groupby("day", as_index=False)["boxes"]
            .sum()
            .sort_values("day")
        )

    mom_rate = None
    if len(monthly_df) >= 2:
        prev = float(monthly_df.iloc[-2]["boxes"])
        curr = float(monthly_df.iloc[-1]["boxes"])
        if abs(prev) > 1e-9:
            mom_rate = (curr - prev) / abs(prev)

    total_boxes = float(filtered["boxes"].sum()) if not filtered.empty else 0.0
    monthly_boxes = float(selected_month_df["boxes"].sum()) if not selected_month_df.empty else 0.0
    latest_day_boxes = float(daily_df.iloc[-1]["boxes"]) if not daily_df.empty else 0.0
    store_count = int(filtered["store"].nunique()) if not filtered.empty else 0

    drill_level = "province"
    drill_path = {"province": None, "distributor": None}
    if province == "全部":
        drill_level = "province"
        key_col = "province"
        name_label = "省区"
    elif distributor == "全部":
        drill_level = "distributor"
        drill_path["province"] = province
        key_col = "distributor"
        name_label = "客户简称"
    else:
        drill_level = "store"
        drill_path["province"] = province
        drill_path["distributor"] = distributor
        key_col = "store"
        name_label = "门店"

    detail_source = filtered.copy()
    month_labels = []
    if not detail_source.empty:
        month_labels = sorted(detail_source["month_label"].dropna().astype(str).unique().tolist(), key=_month_sort_key)

    month_columns: list[dict[str, str]] = []
    all_years = sorted(set(int(y) for y in detail_source["year"].tolist() if int(y) > 0)) if not detail_source.empty else []
    single_year_view = len(all_years) == 1
    for mlabel in month_labels:
        mm = re.match(r"^(20\d{2})-(\d{2})$", mlabel)
        display = mlabel
        if mm and (year_value is not None or single_year_view):
            display = f"{int(mm.group(2))}月"
        month_columns.append({"key": mlabel, "label": display})

    if detail_source.empty or not month_labels:
        drill_rows: list[dict[str, Any]] = []
        month_pivot = pd.DataFrame(columns=["name"])
    else:
        agg = detail_source.groupby([key_col, "month_label"], as_index=False)["boxes"].sum()
        pvt = agg.pivot(index=key_col, columns="month_label", values="boxes").fillna(0.0)
        pvt = pvt.reindex(columns=month_labels, fill_value=0.0).reset_index().rename(columns={key_col: "name"})
        pvt["name"] = pvt["name"].astype(str).str.strip()
        pvt = pvt[pvt["name"] != ""].copy()
        pvt["total_boxes"] = pvt[month_labels].sum(axis=1)
        pvt = pvt.sort_values("total_boxes", ascending=False).reset_index(drop=True)
        month_pivot = pvt
        drill_rows = []

    target_year = selected_year
    if target_year is None and all_years:
        target_year = max(all_years)

    q1_labels = []
    if target_year is not None:
        q1_labels = [f"{int(target_year)}-01", f"{int(target_year)}-02", f"{int(target_year)}-03"]
    q1_labels = [x for x in q1_labels if x in month_labels]

    q1_avg_total = 0.0
    if target_year is not None and not filtered.empty:
        q1_src = filtered[(filtered["year"] == int(target_year)) & (filtered["month"].isin([1, 2, 3]))]
        q1_avg_total = float(q1_src["boxes"].sum()) / 3.0 if not q1_src.empty else 0.0

    if not month_pivot.empty:
        for mlabel in [f"{int(target_year)}-01" if target_year else "", f"{int(target_year)}-02" if target_year else "", f"{int(target_year)}-03" if target_year else ""]:
            if mlabel and mlabel not in month_pivot.columns:
                month_pivot[mlabel] = 0.0
        if target_year is not None:
            m1 = pd.to_numeric(month_pivot.get(f"{int(target_year)}-01", 0), errors="coerce").fillna(0.0)
            m2 = pd.to_numeric(month_pivot.get(f"{int(target_year)}-02", 0), errors="coerce").fillna(0.0)
            m3 = pd.to_numeric(month_pivot.get(f"{int(target_year)}-03", 0), errors="coerce").fillna(0.0)
            month_pivot["q1_avg_boxes"] = (m1 + m2 + m3) / 3.0
        else:
            month_pivot["q1_avg_boxes"] = 0.0

    inv_total = 0.0
    if not stock_df.empty:
        inv_total_src = stock_df.copy()
        if province != "全部":
            inv_total_src = inv_total_src[inv_total_src["province"] == province]
        if distributor != "全部":
            inv_total_src = inv_total_src[
                inv_total_src["distributor"].astype(str).str.replace(r"\s+", "", regex=True) == _norm_key(distributor)
            ]
        if category != "全部" and "category" in inv_total_src.columns:
            inv_total_src = inv_total_src[inv_total_src["category"] == category]
        if weight != "全部" and "weight" in inv_total_src.columns:
            inv_total_src = inv_total_src[
                inv_total_src["weight"].map(_normalize_weight_value) == _normalize_weight_value(weight)
            ]
        inv_total = float(inv_total_src["boxes"].sum()) if not inv_total_src.empty else 0.0

    sellable_total = (inv_total / q1_avg_total) if q1_avg_total > 0 else 0.0

    # Inventory merge: province / customer level available, store level kept 0.
    inv_map: dict[str, float] = {}
    if not stock_df.empty and drill_level in ("province", "distributor"):
        inv_src = stock_df.copy()
        if province != "全部":
            inv_src = inv_src[inv_src["province"] == province]
        if distributor != "全部":
            inv_src = inv_src[inv_src["distributor"].astype(str).str.replace(r"\s+", "", regex=True) == _norm_key(distributor)]
        if category != "全部" and "category" in inv_src.columns:
            inv_src = inv_src[inv_src["category"] == category]
        if weight != "全部" and "weight" in inv_src.columns:
            inv_src = inv_src[inv_src["weight"].map(_normalize_weight_value) == _normalize_weight_value(weight)]

        if drill_level == "province":
            inv_agg = inv_src.groupby("province", as_index=False)["boxes"].sum()
        else:
            inv_src["distributor"] = inv_src["distributor"].astype(str).str.replace(r"\s+", "", regex=True)
            inv_agg = inv_src.groupby("distributor", as_index=False)["boxes"].sum()
        inv_map = {_norm_key(r.iloc[0]): float(r.iloc[1]) for _, r in inv_agg.iterrows()}

    # New-customer metrics: current month / Jan-Feb-Mar / cumulative.
    nc_scope = _empty_df(CANON_NEWCUST_COLS)
    nc_cur_map: dict[str, float] = {}
    nc_prev3_map: dict[str, float] = {}
    nc_cum_map: dict[str, float] = {}
    if not newcust_df.empty:
        nc = newcust_df.copy()
        if province != "全部":
            nc = nc[nc["province"] == province]
        if distributor != "全部":
            nc = nc[nc["distributor"].astype(str).str.replace(r"\s+", "", regex=True) == _norm_key(distributor)]
        if not nc.empty:
            nc_scope = nc.copy()
            cur_ym = int(selected_year * 100 + selected_month) if selected_year and selected_month else int(nc["ym"].max())
            base_year = int(cur_ym // 100) if cur_ym > 0 else (int(target_year) if target_year else 0)
            prev3_yms = [int(base_year * 100 + 1), int(base_year * 100 + 2), int(base_year * 100 + 3)] if base_year > 0 else []

            nc_key_col = "province" if drill_level == "province" else ("distributor" if drill_level == "distributor" else "store")
            nc[nc_key_col] = nc[nc_key_col].astype(str).str.strip()
            if nc_key_col == "distributor":
                nc[nc_key_col] = nc[nc_key_col].str.replace(r"\s+", "", regex=True)

            cur_df = nc[nc["ym"] == cur_ym].groupby(nc_key_col, as_index=False)["new_customers"].sum()
            p3_df = nc[nc["ym"].isin(prev3_yms)].groupby(nc_key_col, as_index=False)["new_customers"].sum() if prev3_yms else pd.DataFrame(columns=[nc_key_col, "new_customers"])
            cum_df = nc[nc["ym"] <= cur_ym].groupby(nc_key_col, as_index=False)["new_customers"].sum()

            nc_cur_map = {_norm_key(r.iloc[0]): float(r.iloc[1]) for _, r in cur_df.iterrows()}
            nc_prev3_map = {_norm_key(r.iloc[0]): float(r.iloc[1]) for _, r in p3_df.iterrows()}
            nc_cum_map = {_norm_key(r.iloc[0]): float(r.iloc[1]) for _, r in cum_df.iterrows()}

    kpi_current_newcust = float(sum(nc_cur_map.values())) if nc_cur_map else 0.0
    kpi_prev3_newcust = float(sum(nc_prev3_map.values())) if nc_prev3_map else 0.0
    kpi_cumulative_newcust = float(sum(nc_cum_map.values())) if nc_cum_map else 0.0

    newcust_monthly_df = pd.DataFrame(columns=["ym", "new_customers"])
    if not nc_scope.empty:
        nc_trend = nc_scope.copy()
        if year_value is not None:
            nc_trend = nc_trend[nc_trend["year"] == int(year_value)]
        newcust_monthly_df = nc_trend.groupby("ym", as_index=False)["new_customers"].sum().sort_values("ym")

    count_map_1: dict[str, int] = {}
    count_map_2: dict[str, int] = {}
    if not detail_source.empty:
        if drill_level == "province":
            g = detail_source.groupby("province", as_index=False).agg(customer_count=("distributor", "nunique"), store_count=("store", "nunique"))
            count_map_1 = {_norm_key(r["province"]): int(r["customer_count"]) for _, r in g.iterrows()}
            count_map_2 = {_norm_key(r["province"]): int(r["store_count"]) for _, r in g.iterrows()}
        elif drill_level == "distributor":
            g = detail_source.groupby("distributor", as_index=False).agg(store_count=("store", "nunique"))
            count_map_1 = {_norm_key(r["distributor"]): int(r["store_count"]) for _, r in g.iterrows()}
        else:
            g = detail_source.groupby("store", as_index=False).agg(product_count=("product", "nunique"))
            count_map_1 = {_norm_key(r["store"]): int(r["product_count"]) for _, r in g.iterrows()}

    metrics_columns = [
        {"key": "total_boxes", "label": "累计出库(箱)"},
        {"key": "q1_avg_boxes", "label": f"{int(target_year)}年1-3月均出库" if target_year else "1-3月均出库"},
        {"key": "inventory_boxes", "label": "库存(箱)"},
        {"key": "sellable_months", "label": "可销月"},
        {"key": "current_new_customers", "label": "本月新客"},
        {"key": "prev3_new_customers", "label": "近三月新客(1-3月)"},
        {"key": "cumulative_new_customers", "label": "累计新客"},
    ]
    if drill_level == "province":
        metrics_columns += [
            {"key": "customer_count", "label": "客户数"},
            {"key": "store_count", "label": "门店数"},
        ]
    elif drill_level == "distributor":
        metrics_columns += [{"key": "store_count", "label": "门店数"}]
    else:
        metrics_columns += [{"key": "product_count", "label": "产品数"}]

    drill_rows: list[dict[str, Any]] = []
    if not month_pivot.empty:
        for _, r in month_pivot.iterrows():
            row_name = str(r["name"])
            key = _norm_key(row_name)
            q1_avg = float(pd.to_numeric(r.get("q1_avg_boxes", 0), errors="coerce"))
            inv = float(inv_map.get(key, 0.0)) if drill_level in ("province", "distributor") else 0.0
            sellable = (inv / q1_avg) if q1_avg > 0 else 0.0

            item: dict[str, Any] = {
                "name": row_name,
                "name_label": name_label,
                "total_boxes": _safe_round(float(r.get("total_boxes", 0.0)), 2),
                "q1_avg_boxes": _safe_round(q1_avg, 2),
                "inventory_boxes": _safe_round(inv, 2),
                "sellable_months": _safe_round(sellable, 1),
                "current_new_customers": _safe_round(float(nc_cur_map.get(key, 0.0)), 2),
                "prev3_new_customers": _safe_round(float(nc_prev3_map.get(key, 0.0)), 2),
                "cumulative_new_customers": _safe_round(float(nc_cum_map.get(key, 0.0)), 2),
            }
            if drill_level == "province":
                item["customer_count"] = int(count_map_1.get(key, 0))
                item["store_count"] = int(count_map_2.get(key, 0))
            elif drill_level == "distributor":
                item["store_count"] = int(count_map_1.get(key, 0))
            else:
                item["product_count"] = int(count_map_1.get(key, 0))

            for mdef in month_columns:
                mk = mdef["key"]
                item[mk] = _safe_round(float(pd.to_numeric(r.get(mk, 0), errors="coerce")), 2)
            drill_rows.append(item)

    # Keep KPI inventory/sellable aligned with the visible drill table for province/customer views.
    kpi_inventory_boxes = inv_total
    kpi_q1_avg_boxes = q1_avg_total
    kpi_sellable_months = sellable_total
    if drill_level in ("province", "distributor") and drill_rows:
        kpi_inventory_boxes = float(sum(float(r.get("inventory_boxes", 0.0) or 0.0) for r in drill_rows))
        kpi_q1_avg_boxes = float(sum(float(r.get("q1_avg_boxes", 0.0) or 0.0) for r in drill_rows))
        kpi_sellable_months = (kpi_inventory_boxes / kpi_q1_avg_boxes) if kpi_q1_avg_boxes > 0 else 0.0

    year_options_vals = sorted(set(int(y) for y in out_df["year"].tolist() if int(y) > 0))
    year_options = ["全部"] + [str(y) for y in year_options_vals]

    return {
        "meta": {
            "file_name": dataset.get("file_name", ""),
            "updated_at": dataset.get("updated_at", ""),
            "outbound_sheet": dataset.get("meta", {}).get("outbound_sheet", ""),
            "newcust_sheet": dataset.get("meta", {}).get("newcust_sheet", ""),
        },
        "filters": {
            "province_options": ["全部"] + _sorted_unique(out_df["province"]),
            "distributor_options": dist_options,
            "category_options": cat_options,
            "weight_options": weight_options,
            "year_options": year_options,
            "month_options": month_options,
            "selected": {
                "province": province if province else "全部",
                "distributor": distributor if distributor else "全部",
                "category": category if category else "全部",
                "weight": weight if weight else "全部",
                "year": year if year else "全部",
                "month": month if month else "全部",
            },
        },
        "kpis": {
            "total_boxes": _safe_round(total_boxes, 2),
            "monthly_boxes": _safe_round(monthly_boxes, 2),
            "latest_day_boxes": _safe_round(latest_day_boxes, 2),
            "store_count": store_count,
            "inventory_boxes": _safe_round(kpi_inventory_boxes, 2),
            "q1_avg_boxes": _safe_round(kpi_q1_avg_boxes, 2),
            "sellable_months": _safe_round(kpi_sellable_months, 1),
            "current_new_customers": _safe_round(kpi_current_newcust, 2),
            "prev3_new_customers": _safe_round(kpi_prev3_newcust, 2),
            "cumulative_new_customers": _safe_round(kpi_cumulative_newcust, 2),
            "mom_rate": _safe_round(mom_rate * 100.0, 2) if mom_rate is not None else None,
            "selected_month_label": selected_month_label,
        },
        "monthly_trend": {
            "labels": monthly_df["month_label"].tolist() if not monthly_df.empty else [],
            "values": [_safe_round(v, 2) for v in monthly_df["boxes"].tolist()] if not monthly_df.empty else [],
        },
        "daily_trend": {
            "labels": [str(int(x)) for x in daily_df["day"].tolist()] if not daily_df.empty else [],
            "values": [_safe_round(v, 2) for v in daily_df["boxes"].tolist()] if not daily_df.empty else [],
        },
        "newcust_monthly_trend": {
            "labels": [f"{int(x) // 100}-{int(x) % 100:02d}" for x in newcust_monthly_df["ym"].tolist()]
            if not newcust_monthly_df.empty
            else [],
            "values": [_safe_round(v, 2) for v in newcust_monthly_df["new_customers"].tolist()]
            if not newcust_monthly_df.empty
            else [],
        },
        "drill": {
            "level": drill_level,
            "path": drill_path,
            "name_label": name_label,
            "month_columns": month_columns,
            "metrics_columns": metrics_columns,
            "rows": drill_rows,
        },
    }


def _snapshot_state() -> dict[str, Any]:
    with _state_lock:
        if not APP_STATE["loaded"]:
            return {"loaded": False}
        return {
            "loaded": True,
            "file_name": APP_STATE["file_name"],
            "updated_at": APP_STATE["updated_at"],
            "main": APP_STATE["main"].copy(),
            "stock": APP_STATE["stock"].copy(),
            "scan": APP_STATE["scan"].copy(),
            "outbound": APP_STATE["outbound"].copy(),
            "newcust": APP_STATE["newcust"].copy(),
            "meta": dict(APP_STATE["meta"]),
        }


@app.get("/")
def index():
    with _state_lock:
        loaded = bool(APP_STATE["loaded"])
        file_name = APP_STATE["file_name"]
    return render_template("index.html", loaded=loaded, file_name=file_name)


@app.post("/api/upload")
def api_upload():
    file_obj = request.files.get("file")
    if file_obj is None:
        return jsonify({"ok": False, "error": "未检测到上传文件。"}), 400

    file_name = file_obj.filename or "uploaded.xlsx"
    file_bytes = file_obj.read()
    if not file_bytes:
        return jsonify({"ok": False, "error": "上传文件为空。"}), 400

    try:
        parsed = _parse_uploaded_dataset(file_bytes, file_name)
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    with _state_lock:
        APP_STATE["loaded"] = True
        APP_STATE["file_name"] = file_name
        APP_STATE["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        APP_STATE["main"] = parsed["main"]
        APP_STATE["stock"] = parsed["stock"]
        APP_STATE["scan"] = parsed["scan"]
        APP_STATE["outbound"] = parsed["outbound"]
        APP_STATE["newcust"] = parsed["newcust"]
        APP_STATE["meta"] = parsed["meta"]

    dataset = _snapshot_state()
    dashboard_payload = _build_dashboard_payload(dataset)
    outbound_payload = _build_outbound_payload(dataset)
    return jsonify(
        {
            "ok": True,
            "message": "文件解析完成。",
            "data": dashboard_payload,
            "dashboard": dashboard_payload,
            "outbound": outbound_payload,
        }
    )


@app.get("/api/dashboard")
def api_dashboard():
    dataset = _snapshot_state()
    if not dataset.get("loaded"):
        return jsonify({"ok": False, "error": "当前没有加载数据，请先上传文件。"}), 400

    province = request.args.get("province", "全部")
    distributor = request.args.get("distributor", "全部")
    category = request.args.get("category", "全部")
    payload = _build_dashboard_payload(dataset, province=province, distributor=distributor, category=category)
    return jsonify({"ok": True, "data": payload})


@app.get("/api/outbound/trend")
def api_outbound_trend():
    dataset = _snapshot_state()
    if not dataset.get("loaded"):
        return jsonify({"ok": False, "error": "当前没有加载数据，请先上传文件。"}), 400

    province = request.args.get("province", "全部")
    distributor = request.args.get("distributor", "全部")
    category = request.args.get("category", "全部")
    weight = request.args.get("weight", "全部")
    year = request.args.get("year", "全部")
    month = request.args.get("month", "全部")

    payload = _build_outbound_payload(
        dataset,
        province=province,
        distributor=distributor,
        category=category,
        weight=weight,
        year=year,
        month=month,
    )
    return jsonify({"ok": True, "data": payload})


@app.get("/api/health")
def api_health():
    return jsonify({"ok": True, "service": "data-screen"})


if __name__ == "__main__":
    # Keep LAN service stable: disable debug reloader to avoid transient 503 during reloads.
    # Default to port 5050 to avoid collisions with other local services on 5000.
    run_host = os.environ.get("HOST", "0.0.0.0")
    run_port = int(os.environ.get("PORT", "5050"))
    app.run(host=run_host, port=run_port, debug=False, use_reloader=False)
