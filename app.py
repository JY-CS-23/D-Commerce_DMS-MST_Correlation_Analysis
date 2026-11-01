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
        raise ValueError(f"연도별 광고비 값을 읽지 못했습니다: {Path(path).name}")
    return out.sort_values("year").drop_duplicates("year", keep="last").reset_index(drop=True)


@st.cache_data(show_spinner="모바일 쇼핑 데이터를 불러오는 중...")
def load_shop(path: str, mtime: float):
    raw = read_sheet(path)
    if len(raw.columns) < 4:
        raise ValueError(f"쇼핑 거래액 파일의 열 구조가 예상과 다릅니다: {Path(path).name}")
    c0, c1, c2 = raw.columns[:3]
    raw[[c0, c1]] = raw[[c0, c1]].replace(r"^\s*$", pd.NA, regex=True).ffill()
    mask = raw[c1].astype(str).str.strip().eq("소계") & raw[c2].astype(str).str.strip().eq("모바일쇼핑")
    sub = raw[mask].copy()
    sub["category"] = sub[c0].astype(str).str.strip().replace({"합계": ALL_LABEL})
    month_cols = {}
    for col in raw.columns[3:]:
        match = MONTH_RE.search(str(col))
        if match:
            month_cols[col] = (int(match.group(1)), int(match.group(2)), bool(str(col)[match.end():].strip()))
    if not month_cols:
        raise ValueError("쇼핑 파일에서 월별 열(예: 2020.01)을 찾지 못했습니다.")
    long = sub.melt(id_vars="category", value_vars=list(month_cols), var_name="col", value_name="raw")
    long["value"] = pd.to_numeric(long["raw"], errors="coerce")
    long["year"] = long["col"].map(lambda c: month_cols[c][0])
    meta = (
        pd.DataFrame(list(month_cols.values()), columns=["year", "month", "provisional"])
        .groupby("year")
        .agg(months=("month", "nunique"), provisional=("provisional", "any"))
        .reset_index()
    )
    annual = (
        long.groupby(["category", "year"])
        .agg(total=("value", "sum"), n=("value", "count"))
        .reset_index()
        .merge(meta, on="year", how="left")
    )
    annual = annual[(annual["months"] == 12) & (annual["n"] == 12)].copy()
    annual["value"] = annual["total"] / 1_000_000
    annual["estimate"] = annual["provisional"].astype(bool)
    skipped = {int(r.year): int(r.months) for r in meta[meta["months"] < 12].itertuples()}
    categories = list(dict.fromkeys(long["category"]))
    table = annual[["category", "year", "value", "estimate"]].sort_values(["category", "year"]).reset_index(drop=True)
    return table, skipped, categories


def build_ad_tables(pc: pd.DataFrame, mobile: pd.DataFrame) -> dict[str, pd.DataFrame]:
    merged = pc.merge(mobile, on="year", suffixes=("_pc", "_mo"))
    total = pd.DataFrame({
        "year": merged["year"],
        "value": merged["value_pc"] + merged["value_mo"],
        "estimate": merged["estimate_pc"] | merged["estimate_mo"],
    })
    return {"PC": pc.copy(), "모바일": mobile.copy(), "합계": total}


def to_yoy(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    values = df.set_index("year")["value"]
    flags = df.set_index("year")["estimate"]
    years = range(int(values.index.min()), int(values.index.max()) + 1)
    full = values.reindex(years)
    growth = (full / full.shift(1) - 1) * 100
    est = flags.reindex(years, fill_value=False).astype(bool)
    return pd.DataFrame({"year": list(years), "value": growth.to_numpy(), "estimate": est.to_numpy()}).dropna(subset=["value"]).reset_index(drop=True)


def prepare(df: pd.DataFrame, mode: str, include_estimates: bool, year_range: tuple[int, int]) -> pd.DataFrame:
    out = df.copy() if include_estimates else df[~df["estimate"]].copy()
    if mode == GROWTH:
        out = to_yoy(out)
    low, high = year_range
    return out[(out["year"] >= low) & (out["year"] <= high)].reset_index(drop=True)


def to_index(ads: dict[str, pd.DataFrame], shops: dict[str, pd.DataFrame]):
    frames = [df for df in [*ads.values(), *shops.values()] if not df.empty]
    if not frames:
        return ads, shops, None
    base_year = max(int(df["year"].min()) for df in frames)

    def rebase(df):
        if df.empty:
            return df
        base = df.loc[df["year"] == base_year, "value"]
        if base.empty or pd.isna(base.iloc[0]) or base.iloc[0] == 0:
            return df.iloc[0:0]
        out = df[df["year"] >= base_year].copy()
        out["value"] = out["value"] / base.iloc[0] * 100
        return out.reset_index(drop=True)

    return ({k: rebase(v) for k, v in ads.items()},
            {k: rebase(v) for k, v in shops.items()}, base_year)


def correlation(a: pd.DataFrame, b: pd.DataFrame) -> tuple[float, int, pd.DataFrame]:
    merged = a[["year", "value"]].merge(b[["year", "value"]], on="year", suffixes=("_ad", "_shop")).dropna()
    n = len(merged)
    if n < 3 or merged["value_ad"].nunique() < 2 or merged["value_shop"].nunique() < 2:
        return float("nan"), n, merged
    return float(merged["value_ad"].corr(merged["value_shop"])), n, merged


def add_estimate_style(fig, df: pd.DataFrame, name: str, color: str, unit: str, secondary_y=None):
    actual = df[~df["estimate"]]
    estimated = df[df["estimate"]]
    kwargs = {"secondary_y": secondary_y} if secondary_y is not None else {}
    if not actual.empty:
        fig.add_trace(go.Scatter(
            x=actual["year"], y=actual["value"], mode="lines+markers", name=name,
            line=dict(color=color, width=2.5), marker=dict(size=7),
            hovertemplate=f"<b>{name}</b><br>%{{x}}년 · %{{y:,.2f}}{unit}<extra></extra>",
        ), **kwargs)
    if not estimated.empty:
        link = pd.concat([actual.tail(1), estimated]).drop_duplicates("year").sort_values("year")
        symbols = ["circle-open" if bool(e) else "circle" for e in link["estimate"]]
        fig.add_trace(go.Scatter(
            x=link["year"], y=link["value"], mode="lines+markers", name=f"{name} (추정·잠정)",
            line=dict(color=color, width=2.5, dash="dash"), marker=dict(size=8, symbol=symbols),
            hovertemplate=f"<b>{name} (추정·잠정)</b><br>%{{x}}년 · %{{y:,.2f}}{unit}<extra></extra>",
        ), **kwargs)


def make_main_figure(ad_series, shop_series, mode):
    dual = mode == AMOUNT and bool(shop_series)
    fig = go.Figure()
    if dual:
        from plotly.subplots import make_subplots
        fig = make_subplots(specs=[[{"secondary_y": True}]])
    unit = {AMOUNT: " 조 원", GROWTH: "%", INDEX: ""}[mode]
    for medium, df in ad_series.items():
        if not df.empty:
            add_estimate_style(fig, df, f"디지털 광고 매출액 ({medium})", AD_COLORS[medium], unit,
                               secondary_y=False if dual else None)
    for i, (name, df) in enumerate(shop_series.items()):
        if not df.empty:
            add_estimate_style(fig, df, name, SHOP_COLORS[i % len(SHOP_COLORS)], unit,
                               secondary_y=True if dual else None)
    fig.update_xaxes(title_text="연도", dtick=1, tickformat="d")
    if dual:
        fig.update_yaxes(title_text="디지털 광고 매출액 (조 원)", rangemode="tozero", secondary_y=False)
        fig.update_yaxes(title_text="모바일 쇼핑 거래액 (조 원)", rangemode="tozero", showgrid=False, secondary_y=True)
    elif mode == AMOUNT:
        fig.update_yaxes(title_text="금액 (조 원)", rangemode="tozero")
    elif mode == INDEX:
        fig.update_yaxes(title_text="지수 (기준연도=100)")
        fig.add_hline(y=100, line_dash="dot", opacity=0.5)
    else:
        fig.update_yaxes(title_text="전년 대비 증가율 (%)")
        fig.add_hline(y=0, line_dash="dot", opacity=0.5)
    fig.update_layout(height=500, hovermode="x unified", margin=dict(l=10, r=10, t=55, b=10),
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0))
    return fig


def describe_series(df: pd.DataFrame) -> dict:
    vals = df["value"].dropna()
    if vals.empty:
        return {}
    return {
        "관측 연도 수": int(vals.size),
        "평균": float(vals.mean()),
        "표준편차(흩어진 정도)": float(vals.std(ddof=1)) if len(vals) > 1 else np.nan,
        "최솟값": float(vals.min()),
        "최솟값 연도": int(df.loc[df["value"].idxmin(), "year"]),
        "최댓값": float(vals.max()),
        "최댓값 연도": int(df.loc[df["value"].idxmax(), "year"]),
        "범위(최대-최소)": float(vals.max() - vals.min()),
    }


def cagr_for_category(df: pd.DataFrame, start_year: int, end_year: int) -> float:
    selected = df[(df["year"] >= start_year) & (df["year"] <= end_year)].sort_values("year")
    if selected.empty:
        return np.nan
    # CAGR은 시작·종료 연도 값이 모두 있고, 기간이 1년 이상이며, 양수일 때 계산한다.
    start = selected[selected["year"] == start_year]["value"]
    end = selected[selected["year"] == end_year]["value"]
    if start.empty or end.empty or start.iloc[0] <= 0 or end.iloc[0] <= 0 or end_year <= start_year:
        return np.nan
    return (float(end.iloc[0]) / float(start.iloc[0])) ** (1 / (end_year - start_year)) * 100 - 100


def main():
    st.set_page_config(page_title="디지털 광고와 모바일 쇼핑 분석", page_icon="📈", layout="wide")
    st.title("디지털 광고 매출액과 모바일 쇼핑 거래액")
    st.caption("2026 데이터분석 연구발표회 · 자료: KOSIS · 금액 단위: 조 원")
    st.info("이 대시보드는 두 지표의 동반 변화를 탐색합니다. 상관관계만으로 광고가 쇼핑 거래액을 증가시켰다고 해석할 수는 없습니다.")

    files = find_data_files()
    missing = [FILE_HINT[k] for k in FILE_RULES if k not in files]
    if missing:
        st.error("엑셀 파일을 찾지 못했습니다. 파일을 `data` 폴더 또는 `app.py`와 같은 폴더에 넣어 주세요.")
        for name in missing:
            st.code(name)
        st.stop()

    try:
        pc = load_ad(str(files["pc"]), files["pc"].stat().st_mtime)
        mobile = load_ad(str(files["mobile"]), files["mobile"].stat().st_mtime)
        shop, skipped_years, categories = load_shop(str(files["shop"]), files["shop"].stat().st_mtime)
    except Exception as e:
        st.error(f"데이터를 읽는 중 오류가 발생했습니다: {e}")
        st.stop()

    ad_tables = build_ad_tables(pc, mobile)
    all_years = list(ad_tables["합계"]["year"]) + list(shop["year"])
    year_min, year_max_all = int(min(all_years)), int(max(all_years))
    actual_years = ad_tables["합계"].loc[~ad_tables["합계"]["estimate"], "year"]
    year_max_actual = int(actual_years.max()) if not actual_years.empty else year_max_all
    with st.sidebar:
        st.header("분석 조건")
        include_now = st.session_state.get(ESTIMATE_KEY, "미포함") == "포함"
        year_max = year_max_all if include_now else year_max_actual
        default_start = max(year_min, 2018)
        year_range = st.slider("연도 범위", min_value=year_min, max_value=year_max,
                               value=(min(default_start, year_max), year_max), step=1)
        st.caption(f"미포함: 최대 {year_max_actual}년 / 포함: 최대 {year_max_all}년")
        medium = st.radio("광고 매체", MEDIA_CHOICES, index=2)
        chosen = st.multiselect("상품군(모바일 쇼핑 거래액)", options=categories, default=[ALL_LABEL] if ALL_LABEL in categories else categories[:1])
        mode = st.radio("값 기준", [AMOUNT, GROWTH, INDEX])
        estimate_choice = st.radio("예측값(추정치)", ["미포함", "포함"], horizontal=True, key=ESTIMATE_KEY)
    include_estimates = estimate_choice == "포함"
    shown_media = ["PC", "모바일"] if medium == MEDIA_BOTH else [medium]

    def series_for(chart_mode):
        ads = {m: prepare(ad_tables[m], chart_mode, include_estimates, year_range) for m in shown_media}
        shops = {}
        for c in chosen:
            frame = shop[shop["category"] == c][["year", "value", "estimate"]]
            shops[f"모바일 쇼핑 거래액 · {c}"] = prepare(frame, chart_mode, include_estimates, year_range)
        return ads, shops

    base_year = None
    if mode == INDEX:
        raw_ads, raw_shops = series_for(AMOUNT)
        ad_series, shop_series, base_year = to_index(raw_ads, raw_shops)
    else:
        ad_series, shop_series = series_for(mode)

    st.subheader("1. 연도별 추이")
    if not chosen:
        st.caption("상품군을 선택하지 않아 광고 매출액만 표시합니다.")
    if all(df.empty for df in ad_series.values()) and all(df.empty for df in shop_series.values()):
        st.warning("선택한 조건에 맞는 데이터가 없습니다. 연도 범위나 예측값 설정을 바꿔 보세요.")
        st.stop()
    st.plotly_chart(make_main_figure(ad_series, shop_series, mode), use_container_width=True)
    if mode == AMOUNT and shop_series:
        st.caption("금액 보기에서는 왼쪽 축이 광고 매출액, 오른쪽 축이 쇼핑 거래액입니다. 두 축의 선 높이를 직접 비교하지 마세요.")
    if mode == INDEX and base_year is not None:
        st.caption(f"{base_year}년을 100으로 환산한 지수입니다. 200은 기준연도의 2배를 뜻합니다.")
    if skipped_years:
        st.caption("쇼핑 거래액에서 연간 12개월이 모두 없는 연도는 제외했습니다: " +
                   ", ".join(f"{y}년({m}개월)" for y, m in skipped_years.items()))
    if include_estimates:
        st.caption("점선과 속이 빈 점은 추정치 또는 잠정치입니다.")

    # 선택한 범위 안의 광고비 총계 분석
    st.subheader("2. 디지털 광고비 총계와 증가액 추세")
    total = prepare(ad_tables["합계"], AMOUNT, include_estimates, year_range).sort_values("year")
    if len(total) >= 2:
        total = total.copy()
        total["증가액(조 원)"] = total["value"].diff()
        fig_ad = go.Figure()
        fig_ad.add_trace(go.Scatter(x=total["year"], y=total["value"], mode="lines+markers",
                                    name="디지털 광고비 합계", line=dict(width=3)))
        if len(total) >= 2:
            coef = np.polyfit(total["year"].astype(float), total["value"].astype(float), 1)
            fig_ad.add_trace(go.Scatter(x=total["year"], y=np.polyval(coef, total["year"]),
                                        mode="lines", name=f"광고비 금액 추세선 (기울기 {coef[0]:.3f}조 원/년)",
                                        line=dict(dash="dash")))
        fig_ad.update_layout(xaxis_title="연도", yaxis_title="광고비 합계(조 원)", height=400,
                             legend=dict(orientation="h", y=1.12))
        st.plotly_chart(fig_ad, use_container_width=True)
        st.caption("추세선 기울기는 연도 1년 증가에 따른 광고비 합계의 선형 추세 변화량입니다. 인과효과를 뜻하지 않습니다.")

        increases = total.dropna(subset=["증가액(조 원)"])
        if not increases.empty:
            inc_slope = np.polyfit(increases["year"].astype(float), increases["증가액(조 원)"].astype(float), 1)[0] if len(increases) >= 2 else np.nan
            c1, c2, c3 = st.columns(3)
            c1.metric("평균 연간 증가액", f"{increases['증가액(조 원)'].mean():.3f}조 원")
            c2.metric("최소 연간 증가액", f"{increases['증가액(조 원)'].min():.3f}조 원", f"{int(increases.loc[increases['증가액(조 원)'].idxmin(), 'year'])}년")
            c3.metric("최대 연간 증가액", f"{increases['증가액(조 원)'].max():.3f}조 원", f"{int(increases.loc[increases['증가액(조 원)'].idxmax(), 'year'])}년")
            if pd.notna(inc_slope):
                st.write(f"연간 증가액 추세선 기울기: **{inc_slope:.4f}조 원/년²** (연간 증가액이 매년 얼마나 커지거나 작아지는지 나타내는 선형 추세)")
    else:
        st.info("광고비 추세선과 연간 증가액을 계산하려면 선택 범위에 연도 2개 이상이 필요합니다.")

    # 상관관계 산점도
    st.subheader("3. 광고비 합계와 모바일 쇼핑 거래액의 상관관계")
    shop_total = shop[shop["category"] == ALL_LABEL][["year", "value", "estimate"]]
    ad_for_corr = prepare(ad_tables["합계"], AMOUNT, include_estimates, year_range)
    shop_for_corr = prepare(shop_total, AMOUNT, include_estimates, year_range)
    r, n, paired = correlation(ad_for_corr, shop_for_corr)
    if not paired.empty:
        scatter = px.scatter(paired, x="value_ad", y="value_shop", hover_data=["year"],
                             labels={"value_ad": "디지털 광고비 합계(조 원)", "value_shop": "모바일 쇼핑 거래액(조 원)"},
                             title="연도별 광고비와 모바일 쇼핑 거래액")
        if len(paired) >= 2:
            slope, intercept = np.polyfit(paired["value_ad"], paired["value_shop"], 1)
            xline = np.linspace(paired["value_ad"].min(), paired["value_ad"].max(), 100)
            scatter.add_trace(go.Scatter(x=xline, y=slope*xline+intercept, mode="lines", name="선형 추세선",
                                         line=dict(dash="dash")))
        st.plotly_chart(scatter, use_container_width=True)
        if pd.notna(r):
            st.metric("피어슨 상관계수 r", f"{r:.3f}", f"겹치는 연도 {n}개")
        else:
            st.warning(f"상관계수를 계산하기에 유효한 겹치는 연도가 부족하거나 값이 일정합니다. 겹치는 연도: {n}개 (최소 3개 필요)")
        st.caption("금액끼리의 상관계수는 두 지표가 장기적으로 함께 증가하는 추세의 영향을 받을 수 있습니다. 산점도와 증가율 분석을 함께 보고, 상관관계를 인과관계로 해석하지 마세요.")
    else:
        st.info("선택한 연도 범위에 두 자료가 겹치는 연도가 없습니다.")

    # 상품군별 CAGR 비교
    st.subheader("4. 상품군별 연평균 거래액 증가율 비교")
    shop_start = st.number_input("CAGR 시작 연도", min_value=year_min, max_value=year_max_all,
                                 value=max(year_min, year_range[0]), step=1, key="cagr_start")
    shop_end = st.number_input("CAGR 종료 연도", min_value=year_min, max_value=year_max_all,
                               value=max(int(shop_start) + 1, year_range[1]), step=1, key="cagr_end")
    if int(shop_end) <= int(shop_start):
        st.warning("CAGR 종료 연도는 시작 연도보다 커야 합니다.")
    else:
        cagr_rows = []
        for cat in categories:
            catdf = shop[(shop["category"] == cat) & (include_estimates | ~shop["estimate"])]
            cagr = cagr_for_category(catdf, int(shop_start), int(shop_end))
            cagr_rows.append({"상품군": cat, "연평균 증가율(CAGR, %)": cagr})
        cagr_table = pd.DataFrame(cagr_rows).dropna().sort_values("연평균 증가율(CAGR, %)", ascending=False)
        if cagr_table.empty:
            st.info("선택한 시작·종료 연도에 유효한 거래액이 없어 CAGR을 계산할 수 없습니다. 두 연도 모두 자료가 있는지 확인하세요.")
        else:
            fig_cagr = px.bar(cagr_table, x="상품군", y="연평균 증가율(CAGR, %)", title="상품군별 연평균 거래액 증가율",
                              labels={"연평균 증가율(CAGR, %)": "연평균 증가율(%)", "상품군": "상품군"},
                              color="상품군")
            fig_cagr.update_layout(showlegend=False, height=450, xaxis_tickangle=-35)
            st.plotly_chart(fig_cagr, use_container_width=True)
            st.dataframe(cagr_table.round(2), use_container_width=True, hide_index=True)
            st.caption("CAGR은 시작 연도와 종료 연도의 거래액을 바탕으로 계산한 연평균 복합성장률입니다. 매년 실제 증가율의 산술평균과는 다릅니다.")

    # 분산/극값/증가율 범위
    st.subheader("5. 흩어진 정도·극값·증가율의 최솟값과 최댓값")
    st.caption("아래 요약은 선택한 연도 범위와 예측값 설정을 적용합니다. 표준편차와 범위가 클수록 해당 기간 값의 변동 폭이 큽니다.")
    summary_rows = []
    series_to_summarize = {}
    for m in MEDIA:
        series_to_summarize[f"광고비 {m}"] = prepare(ad_tables[m], AMOUNT, include_estimates, year_range)
    series_to_summarize["모바일 쇼핑 거래액 · 전체 합계"] = prepare(shop_total, AMOUNT, include_estimates, year_range)
    for label, df in series_to_summarize.items():
        desc = describe_series(df)
        if desc:
            summary_rows.append({"지표": label, **desc})
    if summary_rows:
        summary = pd.DataFrame(summary_rows)
        for col in ["평균", "표준편차(흩어진 정도)", "최솟값", "최댓값", "범위(최대-최소)"]:
            summary[col] = summary[col].round(3)
        st.dataframe(summary, use_container_width=True, hide_index=True)
    growth_rows = []
    for label, df in series_to_summarize.items():
        growth = prepare(df.assign(estimate=df["estimate"]), GROWTH, include_estimates, year_range) if not df.empty else pd.DataFrame()
        # 증가율은 원자료에서 계산해야 범위 첫 해의 전년 값을 유지하므로 원자료를 다시 사용한다.
        source = ad_tables[label.replace("광고비 ", "")] if label.startswith("광고비 ") else shop_total
        if label.startswith("광고비 "):
            growth = prepare(source, GROWTH, include_estimates, year_range)
        else:
            growth = prepare(source, GROWTH, include_estimates, year_range)
        if not growth.empty:
            imin, imax = growth["value"].idxmin(), growth["value"].idxmax()
            growth_rows.append({
                "지표": label,
                "증가율 최솟값(%)": growth.loc[imin, "value"],
                "최솟값 연도": int(growth.loc[imin, "year"]),
                "증가율 최댓값(%)": growth.loc[imax, "value"],
                "최댓값 연도": int(growth.loc[imax, "year"]),
                "증가율 표준편차(퍼센트포인트)": growth["value"].std(ddof=1) if len(growth) > 1 else np.nan,
            })
    if growth_rows:
        growth_summary = pd.DataFrame(growth_rows)
        numeric = ["증가율 최솟값(%)", "증가율 최댓값(%)", "증가율 표준편차(퍼센트포인트)"]
        growth_summary[numeric] = growth_summary[numeric].round(2)
        st.dataframe(growth_summary, use_container_width=True, hide_index=True)
    st.caption("극값은 선택한 기간 안의 최솟값·최댓값입니다. 표준편차는 평균에서 얼마나 흩어져 있는지를 요약하지만, 이상치의 원인을 설명하지는 않습니다.")

    # 원자료/그래프 값 다운로드
    st.subheader("6. 그래프에 사용된 데이터")
    with st.expander("연도별 광고비·쇼핑 거래액 표 보기 / CSV 내려받기"):
        table_parts = []
        for m in MEDIA:
            frame = prepare(ad_tables[m], mode if mode != INDEX else AMOUNT, include_estimates, year_range).copy()
            frame["계열"] = f"광고비 {m}"
            table_parts.append(frame)
        for cat in chosen:
            frame = prepare(shop[shop["category"] == cat][["year", "value", "estimate"]], mode if mode != INDEX else AMOUNT,
                            include_estimates, year_range).copy()
            frame["계열"] = f"모바일 쇼핑 · {cat}"
            table_parts.append(frame)
        if table_parts:
            values = pd.concat(table_parts, ignore_index=True)[["year", "계열", "value", "estimate"]]
            st.dataframe(values, use_container_width=True, hide_index=True)
            st.download_button("CSV로 내려받기", values.to_csv(index=False).encode("utf-8-sig"),
                               file_name="디지털광고_모바일쇼핑_분석값.csv", mime="text/csv")


if __name__ == "__main__":
    main()
