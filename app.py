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
