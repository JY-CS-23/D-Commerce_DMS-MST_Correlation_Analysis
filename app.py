from __future__ import annotations

import re
import unicodedata
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
SEARCH_DIRS = [APP_DIR / "data", APP_DIR]
ALL_LABEL = "전체(합계)"
AMOUNT = "금액"
GROWTH = "전년 대비 증가율"
INDEX = "지수(기준연도=100)"
MEDIA = ["PC", "모바일", "합계"]
MEDIA_BOTH = "PC·모바일 각각"
MEDIA_CHOICES = MEDIA + [MEDIA_BOTH]
ESTIMATE_KEY = "estimate_choice"
YEAR_RE = re.compile(r"(\d{4})")
MONTH_RE = re.compile(r"(\d{4})\.(\d{1,2})")
FILE_RULES = {
    "pc": lambda n: n.startswith("PC") and "광고비" in n,
    "mobile": lambda n: n.startswith("모바일") and "광고비" in n,
    "shop": lambda n: "거래액" in n,
}
FILE_HINT = {
    "pc": "PC_세부광고유형별_광고비_데이터분석연구발표회.xlsx",
    "mobile": "모바일_세부광고유형별_광고비_데이터분석연구발표회.xlsx",
    "shop": "온라인쇼핑몰_판매매체별_상품군별거래액_데이터분석연구발표회.xlsx",
}
AD_COLORS = {"PC": "#7aa7ff", "모바일": "#2548c9", "합계": "#2563eb"}
SHOP_COLORS = px.colors.qualitative.Safe


def nfc(text: str) -> str:
    return unicodedata.normalize("NFC", text)


def find_data_files() -> dict[str, Path]:
    found: dict[str, Path] = {}
    for folder in SEARCH_DIRS:
        if not folder.is_dir():
            continue
        for path in sorted(folder.glob("*.xlsx")):
            name = nfc(path.name)
            if name.startswith("~$"):
                continue
            for key, rule in FILE_RULES.items():
                if key not in found and rule(name):
                    found[key] = path
    return found


def read_sheet(path: str) -> pd.DataFrame:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        excel = pd.ExcelFile(path, engine="openpyxl")
        sheet = "데이터" if "데이터" in excel.sheet_names else excel.sheet_names[0]
        return excel.parse(sheet, header=0, dtype=object)


@st.cache_data(show_spinner="광고비 데이터를 불러오는 중...")
def load_ad(path: str, mtime: float) -> pd.DataFrame:
    raw = read_sheet(path)
    if len(raw.columns) < 4:
        raise ValueError(f"광고비 파일의 열 구조가 예상과 다릅니다: {Path(path).name}")
    c0, c1, c2 = raw.columns[:3]
    mask = (
        raw[c0].astype(str).str.strip().eq("총계")
        & raw[c1].astype(str).str.strip().eq("총계")
        & raw[c2].astype(str).str.strip().eq("소계")
    )
    if not mask.any():
        raise ValueError("'총계-총계-소계' 행을 찾지 못했습니다.")
    row = raw[mask].iloc[0]
    records = []
    for col in raw.columns[3:]:
        match = YEAR_RE.search(str(col))
        if not match:
            continue
        val = pd.to_numeric(row[col], errors="coerce")
        if pd.isna(val):
            continue
        suffix = str(col)[match.end():].strip()
        records.append({
            "year": int(match.group(1)),
            "value": float(val) / 1_000_000,
            "estimate": bool(suffix),
        })
    out = pd.DataFrame(records)
    if out.empty:
