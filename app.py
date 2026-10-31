"""
디지털 광고 매출액과 모바일 쇼핑 거래액 대시보드 (Streamlit)

세 부분으로 나뉘어 있습니다. (팀원 모두 이 세 부분을 설명할 수 있어야 합니다.)
  1) 데이터 불러오기 : 엑셀 3개를 읽어 연도별 표로 만든다.   → find_data_files / load_ad / load_shop
  2) 사이드바 만들기 : 연도 범위, 광고 매체, 상품군, 값 기준, 예측값 포함 여부  → main() 의 with st.sidebar
  3) 차트 그리기    : 광고 매출액과 모바일 쇼핑 거래액을 한 그래프에 겹쳐 그린다. → make_figure

데이터 (모두 KOSIS)
  - PC 세부광고유형별 광고비       : 방송통신광고비조사, 한국방송광고진흥공사·방송미디어통신위원회 (단위: 백만원)
  - 모바일 세부광고유형별 광고비   : 위와 같음 (단위: 백만원)
  - 온라인쇼핑몰 판매매체별/상품군별 거래액 : 온라인쇼핑동향조사, 국가데이터처 (단위: 백만원, 월별)
"""

from __future__ import annotations

import re
import unicodedata
import warnings
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

# ──────────────────────────────────────────────────────────────
# 설정값
# ──────────────────────────────────────────────────────────────
APP_DIR = Path(__file__).resolve().parent
SEARCH_DIRS = [APP_DIR / "data", APP_DIR]  # 엑셀 파일을 찾는 폴더 (data 폴더 → app.py 와 같은 폴더)

AMOUNT = "금액"
GROWTH = "전년 대비 증가율"
INDEX = "지수(기준연도=100)"
ALL_LABEL = "전체(합계)"
MEDIA = ["PC", "모바일", "합계"]
INCLUDE = "포함"
EXCLUDE = "미포함"

AD_COLOR = "#2563eb"
SHOP_COLORS = [
    "#d95f02", "#1b9e77", "#7570b3", "#e7298a",
    "#66a61e", "#e6ab02", "#a6761d", "#555555",
]

# 파일 이름 규칙 (macOS에서 만든 파일 이름은 한글 조합 방식이 달라서 NFC로 맞춰 비교한다)
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

YEAR_RE = re.compile(r"(\d{4})")
MONTH_RE = re.compile(r"(\d{4})\.(\d{1,2})")


# ──────────────────────────────────────────────────────────────
# 1) 데이터 불러오기
# ──────────────────────────────────────────────────────────────
def _nfc(text: str) -> str:
    return unicodedata.normalize("NFC", text)


def find_data_files() -> dict[str, Path]:
    """data 폴더(또는 app.py 옆)에서 엑셀 3개를 찾아 {'pc','mobile','shop': 경로}로 돌려준다."""
    found: dict[str, Path] = {}
    for folder in SEARCH_DIRS:
        if not folder.is_dir():
            continue
        for path in sorted(folder.glob("*.xlsx")):
            name = _nfc(path.name)
            if name.startswith("~$"):  # 엑셀 임시 파일
                continue
            for key, rule in FILE_RULES.items():
                if key not in found and rule(name):
                    found[key] = path
    return found


def _read_sheet(path: str) -> pd.DataFrame:
    """KOSIS 엑셀의 '데이터' 시트를 모두 문자열/원본 그대로(object) 읽는다."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")  # KOSIS 파일의 '기본 스타일 없음' 경고 숨김
        excel = pd.ExcelFile(path, engine="openpyxl")
        sheet = "데이터" if "데이터" in excel.sheet_names else excel.sheet_names[0]
        return excel.parse(sheet, header=0, dtype=object)


@st.cache_data(show_spinner="광고비 데이터를 불러오는 중...")
def load_ad(path: str, mtime: float) -> pd.DataFrame:
    """광고비 엑셀 → 연도별 표(year, value[조 원], estimate).

    - 파일에는 '총계/총계/소계' 행 외에 권역별·광고매출액 규모별·광고유형별 행이 함께 있다.
      같은 금액을 쪼갠 행이라 모두 더하면 중복되므로 '총계-총계-소계' 한 행만 쓴다.
    - 열 이름이 '2025 e)'처럼 붙은 연도는 추정치(estimate=True)이다.
    - 단위: 백만원 → 조 원 (÷ 1,000,000)
    """
    raw = _read_sheet(path)
    c0, c1, c2 = raw.columns[:3]
    is_total = (
        (raw[c0].astype(str).str.strip() == "총계")
        & (raw[c1].astype(str).str.strip() == "총계")
        & (raw[c2].astype(str).str.strip() == "소계")
    )
    if not is_total.any():
        raise ValueError("'총계-총계-소계' 행을 찾지 못했습니다. 파일 구조를 확인하세요.")
    row = raw[is_total].iloc[0]

    records = []
    for col in raw.columns[3:]:
        match = YEAR_RE.search(str(col))
        if not match:
            continue
        value = pd.to_numeric(row[col], errors="coerce")
        suffix = str(col)[match.end():].strip()  # 'e)' 같은 표시
        records.append({"year": int(match.group(1)), "value": value / 1_000_000, "estimate": bool(suffix)})
    out = pd.DataFrame(records).dropna(subset=["value"])
    return out.sort_values("year").reset_index(drop=True)


@st.cache_data(show_spinner="쇼핑 거래액 데이터를 불러오는 중...")
def load_shop(path: str, mtime: float):
    """온라인쇼핑 엑셀 → (연도별 표, 제외된 연도 {연도: 월 수}, 상품군 목록).

    - 상품군별(2)가 '소계'이고 판매매체가 '모바일쇼핑'인 행만 쓴다.
      ('가전·전자', '통신기기' 같은 하위 행은 상위 '가전·전자·통신기기'와 겹치고 2017~2019년 값이 비어 있어 쓰지 않는다.)
    - 월별 값을 연도별로 합친다. 12개월이 다 없는 해(예: 2026년 1~8월)는 연 합계를 비교할 수 없어 뺀다.
    - 열 이름이 '2026.07 p)'처럼 붙은 달이 있는 해는 잠정치(estimate=True)이다.
    - 단위: 백만원 → 조 원
    """
    raw = _read_sheet(path)
    c0, c1, c2 = raw.columns[:3]

    # 상품군 이름은 첫 행에만 적혀 있으므로 아래로 채운다.
    raw[[c0, c1]] = raw[[c0, c1]].replace(r"^\s*$", pd.NA, regex=True).ffill()

    is_mobile = (raw[c1].astype(str).str.strip() == "소계") & (raw[c2].astype(str).str.strip() == "모바일쇼핑")
    sub = raw[is_mobile].copy()
    sub["category"] = sub[c0].astype(str).str.strip().replace({"합계": ALL_LABEL})

    month_cols: dict = {}
    for col in raw.columns[3:]:
        match = MONTH_RE.search(str(col))
        if match:
            provisional = bool(str(col)[match.end():].strip())
            month_cols[col] = (int(match.group(1)), int(match.group(2)), provisional)

    long = sub.melt(id_vars="category", value_vars=list(month_cols), var_name="col", value_name="raw")
    long["value"] = pd.to_numeric(long["raw"], errors="coerce")
    long["year"] = long["col"].map(lambda c: month_cols[c][0])

    meta = (
        pd.DataFrame(list(month_cols.values()), columns=["year", "month", "prov"])
        .groupby("year")
        .agg(months=("month", "nunique"), provisional=("prov", "any"))
        .reset_index()
    )
    annual = (
        long.groupby(["category", "year"])
        .agg(total=("value", "sum"), n=("value", "count"))
        .reset_index()
        .merge(meta, on="year")
    )
    complete = (annual["months"] == 12) & (annual["n"] == 12)
    annual = annual[complete].copy()
    annual["value"] = annual["total"] / 1_000_000
    annual["estimate"] = annual["provisional"].astype(bool)

    skipped = {int(r.year): int(r.months) for r in meta[meta["months"] < 12].itertuples()}
    categories = list(dict.fromkeys(long["category"]))
    table = annual[["category", "year", "value", "estimate"]].sort_values(["category", "year"]).reset_index(drop=True)
    return table, skipped, categories


def build_ad_tables(pc: pd.DataFrame, mobile: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """PC·모바일 표와 둘을 더한 '합계' 표를 만든다. (합계는 두 매체 값이 모두 있는 해만)"""
    merged = pc.merge(mobile, on="year", suffixes=("_pc", "_mo"))
    total = pd.DataFrame({
        "year": merged["year"],
        "value": merged["value_pc"] + merged["value_mo"],
        "estimate": merged["estimate_pc"] | merged["estimate_mo"],
    })
    return {"PC": pc, "모바일": mobile, "합계": total}


# ──────────────────────────────────────────────────────────────
# 계산 도우미
# ──────────────────────────────────────────────────────────────
def to_yoy(df: pd.DataFrame) -> pd.DataFrame:
    """연도별 값 → 전년 대비 증가율(%). 바로 앞 연도가 있는 해만 계산한다."""
    if df.empty:
        return df.copy()
    values = df.set_index("year")["value"]
    flags = df.set_index("year")["estimate"]
    years = range(int(values.index.min()), int(values.index.max()) + 1)
    full = values.reindex(years)
    growth = (full / full.shift(1) - 1) * 100
    est = flags.reindex(years, fill_value=False).astype(bool)
    out = pd.DataFrame({"year": list(years), "value": growth.to_numpy(), "estimate": est.to_numpy()})
    return out.dropna(subset=["value"]).reset_index(drop=True)


def prepare(df: pd.DataFrame, mode: str, include_estimates: bool, year_range: tuple[int, int]) -> pd.DataFrame:
    """예측값 포함 여부 → (증가율이면 변환) → 연도 범위 순서로 걸러낸다.
    증가율은 범위를 자르기 전에 계산해서, 범위 첫 해에도 전년 값이 있으면 증가율이 나오게 한다."""
    out = df if include_estimates else df[~df["estimate"]]
    if mode == GROWTH:
        out = to_yoy(out)
    low, high = year_range
    return out[(out["year"] >= low) & (out["year"] <= high)].reset_index(drop=True)


def to_index(ad: pd.DataFrame, shops: dict[str, pd.DataFrame]):
    """금액 → 지수(기준연도=100). 그려질 모든 선이 값을 가진 가장 이른 해를 기준연도로 삼아,
    모든 선이 같은 해에 100에서 출발하게 한다. (금액이 달라도 '몇 배로 늘었는지'를 같은 축에서 비교)
    돌려주는 값: (광고 표, {이름: 쇼핑 표}, 기준연도)"""
    frames = [df for df in [ad, *shops.values()] if not df.empty]
    if not frames:
        return ad, shops, None
    base_year = max(int(df["year"].min()) for df in frames)

    def rebase(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df
        base = df.loc[df["year"] == base_year, "value"]
        if base.empty or base.iloc[0] == 0:
            return df.iloc[0:0]
        out = df[df["year"] >= base_year].copy()
        out["value"] = out["value"] / base.iloc[0] * 100
        return out.reset_index(drop=True)

    return rebase(ad), {name: rebase(df) for name, df in shops.items()}, base_year


def correlation(a: pd.DataFrame, b: pd.DataFrame) -> tuple[float, int]:
    """두 연도별 표의 피어슨 상관계수와 겹치는 연도 수. 연도가 3개 미만이면 계산하지 않는다."""
    merged = a[["year", "value"]].merge(b[["year", "value"]], on="year", suffixes=("_a", "_b")).dropna()
    n = len(merged)
    if n < 3 or merged["value_a"].nunique() < 2 or merged["value_b"].nunique() < 2:
        return float("nan"), n
    return float(merged["value_a"].corr(merged["value_b"])), n


# ──────────────────────────────────────────────────────────────
# 3) 차트 그리기
# ──────────────────────────────────────────────────────────────
def _add_line(fig, df: pd.DataFrame, name: str, color: str, unit: str, dual: bool, right_axis: bool) -> None:
    """꺾은선 하나를 그린다. 추정치 구간은 점선·속이 빈 점으로 구분한다."""
    kwargs = {"secondary_y": right_axis} if dual else {}
    actual = df[~df["estimate"]]
    estimated = df[df["estimate"]]

    if not actual.empty:
        fig.add_trace(
            go.Scatter(
                x=actual["year"], y=actual["value"], mode="lines+markers",
                name=name, legendgroup=name,
                line=dict(color=color, width=2.5), marker=dict(size=7),
                hovertemplate=f"<b>{name}</b><br>%{{x}}년 · %{{y:,.2f}}{unit}<extra></extra>",
            ),
            **kwargs,
        )
    if not estimated.empty:
        link = pd.concat([actual.tail(1), estimated])  # 마지막 실제값과 이어 붙여 선이 끊기지 않게 한다
        symbols = ["circle"] * (len(link) - len(estimated)) + ["circle-open"] * len(estimated)
        fig.add_trace(
            go.Scatter(
                x=link["year"], y=link["value"], mode="lines+markers",
                name=name, legendgroup=name, showlegend=actual.empty,
                line=dict(color=color, width=2.5, dash="dash"),
                marker=dict(size=8, symbol=symbols),
                hovertemplate=f"<b>{name}</b> (추정·잠정)<br>%{{x}}년 · %{{y:,.2f}}{unit}<extra></extra>",
            ),
            **kwargs,
        )


def make_figure(ad_label: str, ad_series: pd.DataFrame, shop_series: dict[str, pd.DataFrame], mode: str):
    """금액 기준이면 왼쪽 축=광고 매출액, 오른쪽 축=모바일 쇼핑 거래액(이중축).
    증가율·지수 기준이면 단위가 같으므로 축 하나에 겹쳐 그린다."""
    dual = mode == AMOUNT and len(shop_series) > 0
    fig = make_subplots(specs=[[{"secondary_y": True}]]) if dual else go.Figure()
    unit = {AMOUNT: " 조 원", GROWTH: "%", INDEX: ""}[mode]

    if not ad_series.empty:
        _add_line(fig, ad_series, ad_label, AD_COLOR, unit, dual, right_axis=False)
    for i, (name, df) in enumerate(shop_series.items()):
        if not df.empty:
            _add_line(fig, df, name, SHOP_COLORS[i % len(SHOP_COLORS)], unit, dual, right_axis=True)

    fig.update_xaxes(title_text="연도", dtick=1, tickformat="d")
    if dual:
        # 두 축 모두 0에서 시작해 눈금 차이로 인한 착시를 줄인다.
        fig.update_yaxes(title_text="디지털 광고 매출액 (조 원)", rangemode="tozero", secondary_y=False)
        fig.update_yaxes(title_text="모바일 쇼핑 거래액 (조 원)", rangemode="tozero", showgrid=False, secondary_y=True)
    elif mode == AMOUNT:
        fig.update_yaxes(title_text="디지털 광고 매출액 (조 원)", rangemode="tozero")
    elif mode == INDEX:
        fig.update_yaxes(title_text="지수 (기준연도 = 100)")
        fig.add_hline(y=100, line_width=1, line_dash="dot", opacity=0.5)
    else:
        fig.update_yaxes(title_text="전년 대비 증가율 (%)")
        fig.add_hline(y=0, line_width=1, line_dash="dot", opacity=0.5)
    fig.update_layout(
        height=520,
        hovermode="x unified",
        margin=dict(l=10, r=10, t=60, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    return fig


def build_value_table(ad_label: str, ad_series: pd.DataFrame, shop_series: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """그래프에 쓰인 값을 연도 × 계열 표로 만든다 (원본 엑셀과 대조할 때 사용)."""
    columns = {}
    if not ad_series.empty:
        columns[ad_label] = ad_series.set_index("year")["value"]
    for name, df in shop_series.items():
        if not df.empty:
            columns[name] = df.set_index("year")["value"]
    table = pd.DataFrame(columns)
    table.index.name = "연도"
    return table


# ──────────────────────────────────────────────────────────────
# 화면 구성
# ──────────────────────────────────────────────────────────────
def main() -> None:
    st.set_page_config(page_title="디지털 광고 매출액과 모바일 쇼핑 거래액", page_icon="📈", layout="wide")
    st.title("디지털 광고 매출액과 모바일 쇼핑 거래액")
    st.caption("연도별 · 단위: 조 원 · 자료: KOSIS 방송통신광고비조사(한국방송광고진흥공사), 온라인쇼핑동향조사(국가데이터처)")

    # ── 1) 데이터 불러오기 ──
    files = find_data_files()
    missing = [FILE_HINT[k] for k in FILE_RULES if k not in files]
    if missing:
        st.error("엑셀 파일을 찾지 못했습니다. 아래 파일을 `data` 폴더(또는 app.py와 같은 폴더)에 넣어 주세요.")
        for name in missing:
            st.code(name, language=None)
        st.stop()

    try:
        pc = load_ad(str(files["pc"]), files["pc"].stat().st_mtime)
        mobile = load_ad(str(files["mobile"]), files["mobile"].stat().st_mtime)
        shop, skipped_years, categories = load_shop(str(files["shop"]), files["shop"].stat().st_mtime)
    except Exception as error:  # 파일 구조가 달라졌을 때 원인을 화면에 보여 준다
        st.error(f"데이터를 읽는 중 오류가 났습니다: {error}")
        st.stop()
    ad_tables = build_ad_tables(pc, mobile)

    all_years = list(ad_tables["합계"]["year"]) + list(shop["year"])
    year_min, year_max = int(min(all_years)), int(max(all_years))
    estimate_ad_years = sorted(int(y) for y in ad_tables["합계"].loc[ad_tables["합계"]["estimate"], "year"])
    estimate_shop_years = sorted(int(y) for y in shop.loc[shop["estimate"], "year"].unique())

    # ── 2) 사이드바 ──
    with st.sidebar:
        st.header("분석 조건")
        year_range = st.slider(
            "연도 범위",
            min_value=year_min, max_value=year_max,
            value=(max(year_min, 2018), min(year_max, 2024)), step=1,
        )
        medium = st.radio("광고 매체", MEDIA, index=2, help="합계는 PC 광고비와 모바일 광고비를 더한 값입니다.")
        chosen = st.multiselect(
            "상품군 (모바일 쇼핑 거래액)", options=categories, default=[ALL_LABEL],
            help="여러 개를 고르면 상품군마다 선이 하나씩 그려집니다.",
        )
        mode = st.radio(
            "값 기준", [AMOUNT, GROWTH, INDEX],
            help="금액: 광고(왼쪽 축)와 쇼핑(오른쪽 축)의 눈금이 따로입니다. "
                 "증가율: 바로 앞 연도 대비 증감률(%)입니다. "
                 "지수: 기준연도(선택 범위에서 모든 선이 값을 가진 첫 해)를 100으로 맞춰 같은 축에서 비교합니다.",
        )
        estimate_choice = st.radio(
            "예측값(추정치)", [EXCLUDE, INCLUDE], horizontal=True,
            help="광고비의 추정치(e)와 쇼핑 거래액의 잠정치(p)를 그래프에 넣을지 정합니다.",
        )
    include_estimates = estimate_choice == INCLUDE

    # ── 3) 값 준비와 차트 ──
    ad_label = "디지털 광고 매출액 (PC+모바일)" if medium == "합계" else f"디지털 광고 매출액 ({medium})"

    def shop_label(category: str) -> str:
        return f"모바일 쇼핑 거래액 · {category}"

    def series_for(chart_mode: str):
        ad = prepare(ad_tables[medium], chart_mode, include_estimates, year_range)
        shops = {
            shop_label(c): prepare(shop[shop["category"] == c][["year", "value", "estimate"]], chart_mode, include_estimates, year_range)
            for c in chosen
        }
        return ad, shops

    base_year = None
    if mode == INDEX:
        ad_amount_raw, shop_amount_raw = series_for(AMOUNT)
        ad_series, shop_series, base_year = to_index(ad_amount_raw, shop_amount_raw)
    else:
        ad_series, shop_series = series_for(mode)

    if not chosen:
        st.info("사이드바에서 상품군을 고르면 모바일 쇼핑 거래액이 함께 그려집니다. 지금은 광고 매출액만 보입니다.")
    if ad_series.empty and all(df.empty for df in shop_series.values()):
        st.warning("선택한 조건에 해당하는 데이터가 없습니다. 연도 범위를 넓히거나 예측값을 '포함'으로 바꿔 보세요.")
        st.stop()

    st.plotly_chart(make_figure(ad_label, ad_series, shop_series, mode))

    if mode == AMOUNT and shop_series:
        st.caption("※ 왼쪽 축(광고)과 오른쪽 축(쇼핑)의 눈금이 다릅니다. 선의 높이를 서로 비교하지 말고 변화하는 모양을 보세요.")
    if mode == INDEX and base_year is not None:
        st.caption(
            f"※ {base_year}년을 100으로 맞춘 지수입니다. 예를 들어 200이면 {base_year}년의 2배라는 뜻이며, "
            "금액이 아니라 증가한 배수를 비교합니다. 기준연도를 바꾸면(연도 범위를 바꾸면) 선의 위치가 달라집니다."
        )
    notes = []
    if not include_estimates and estimate_ad_years:
        notes.append(f"광고비 {', '.join(map(str, estimate_ad_years))}년은 추정치라 '미포함'에서는 표시되지 않습니다.")
    if include_estimates and (estimate_ad_years or estimate_shop_years):
        notes.append("점선과 속이 빈 점은 추정치(광고비) 또는 잠정치(쇼핑 거래액)입니다.")
    if skipped_years:
        text = ", ".join(f"{y}년({m}개월)" for y, m in skipped_years.items())
        notes.append(f"쇼핑 거래액은 12개월이 모두 없는 해({text})를 연 합계로 비교할 수 없어 제외했습니다.")
    for note in notes:
        st.caption("※ " + note)

    # ── 상관계수 (결론 확인용) ──
    if chosen:
        st.subheader("광고 매출액과의 상관계수")
        ad_amount, shop_amount = series_for(AMOUNT)
        ad_growth, shop_growth = series_for(GROWTH)
        rows = []
        for category in chosen:
            label = shop_label(category)
            r_amount, n_amount = correlation(ad_amount, shop_amount[label])
            r_growth, n_growth = correlation(ad_growth, shop_growth[label])
            rows.append({
                "상품군": category,
                "금액 상관계수": f"{r_amount:.2f}" if pd.notna(r_amount) else "계산 불가",
                "금액 연도 수": n_amount,
                "증가율 상관계수": f"{r_growth:.2f}" if pd.notna(r_growth) else "계산 불가",
                "증가율 연도 수": n_growth,
            })
        st.dataframe(pd.DataFrame(rows), hide_index=True)
        st.caption(
            "※ 두 값이 모두 해마다 늘어난 추세라서 금액의 상관계수는 높게 나오기 쉽습니다. "
            "연도 수가 적어 값이 불안정하며, 상관관계가 있어도 광고가 거래액을 늘렸다는 뜻은 아닙니다(상관과 인과는 다릅니다). "
            "연도 수가 3개 미만이면 계산하지 않습니다."
        )

    # ── 그래프에 쓰인 값 ──
    with st.expander("그래프에 쓰인 값 보기 / 내려받기"):
        table = build_value_table(ad_label, ad_series, shop_series)
        st.dataframe(table.round(3))
        st.download_button(
            "CSV로 내려받기", data=table.to_csv().encode("utf-8-sig"),
            file_name="그래프_값.csv", mime="text/csv",
        )


if __name__ == "__main__":
    main()
