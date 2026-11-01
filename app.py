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
