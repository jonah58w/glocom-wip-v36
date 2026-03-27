def render_sales_detail_from_teable(source_df: pd.DataFrame):
    """
    業績明細表 - 修正版（参照附件截图格式）
    """
    import datetime as _dt
    st.subheader("業績明細表")

    if source_df is None or source_df.empty:
        st.warning("Teable 主表目前沒有資料。")
        return

    st.caption("資料來源：Teable 主表即時欄位（全客戶）")

    FX_NTD_PER_USD = 31.5

    def _pick_col(*candidate_groups):
        merged = []
        for group in candidate_groups:
            merged.extend(list(group))
        return first_existing_column(source_df, merged)

    def _parse_date_from_col(col_name):
        if not col_name:
            return pd.Series(pd.NaT, index=source_df.index)
        return parse_mixed_date_series(get_series_by_col(source_df, col_name))

    def _clean_factory(v):
        txt = str(v).strip()
        if txt.lower() in {"nan", "none"}:
            return " "
        return re.sub(r"\s+", " ", txt)

    def _clean_customer(v):
        txt = str(v).strip()
        if txt.lower() in {"nan", "none"}:
            return " "
        m = re.match(r"^([A-Za-z0-9_\-]+)", txt)
        return m.group(1) if m else (txt.split()[0] if txt.split() else txt)

    def _fmt_usd(v):
        return f"${float(v):,.2f}"

    def _usd_to_ntd_10k(v):
        return round((float(v) * FX_NTD_PER_USD) / 10000.0, 2)

    # 欄位偵測
    customer_col_local = _pick_col(CUSTOMER_CANDIDATES, ["Customer"])
    factory_col_local = _pick_col(FACTORY_CANDIDATES)
    po_col_local = _pick_col(PO_CANDIDATES)
    pn_col_local = _pick_col(PART_CANDIDATES)
    qty_col_local = _pick_col(QTY_CANDIDATES, ["Order QTY (PCS)", "Order QTY", "QTY"])
    wip_col_local = _pick_col(WIP_CANDIDATES)

    order_date_col = _pick_col(ORDER_DATE_CANDIDATES, ["客戶下單日期", "工廠下單日期", "下單日期", "接單日期"])
    actual_ship_col = _pick_col(["出貨日期", "出貨日期_排序"], SHIP_DATE_CANDIDATES)
    planned_ship_col = _pick_col(["Ship date", "Ship Date", "預計出貨日", "客戶交期"], SHIP_DATE_CANDIDATES)
    effective_ship_col = actual_ship_col or planned_ship_col

    sales_amt_col = find_amount_column(
        source_df,
        ["銷貨金額", "出貨金額", "出貨發票金額", "Invoice Amount", "Invoice Total", "Invoice", "INVOICE", "發票"]
        + AMOUNT_SHIP_CANDIDATES
    )
    order_amt_col = find_amount_column(
        source_df,
        ["接單金額", "接單總金額", "Order Amount", "Order Total", "客戶金額", "Total Amount", "Amount"]
        + AMOUNT_ORDER_CANDIDATES
    )

    # 日期解析
    order_dates = _parse_date_from_col(order_date_col)
    actual_ship_dates = _parse_date_from_col(actual_ship_col)
    planned_ship_dates = _parse_date_from_col(planned_ship_col)

    effective_ship_dates = actual_ship_dates.copy()
    missing_ship_date = effective_ship_dates.isna()
    effective_ship_dates.loc[missing_ship_date] = planned_ship_dates.loc[missing_ship_date]

    # 金額解析
    ship_amt_series = parse_numeric_series(get_series_by_col(source_df, sales_amt_col) if sales_amt_col else None)
    order_amt_series = parse_numeric_series(get_series_by_col(source_df, order_amt_col) if order_amt_col else None)

    if len(ship_amt_series) != len(source_df):
        ship_amt_series = pd.Series(0.0, index=source_df.index)
    else:
        ship_amt_series.index = source_df.index
    if len(order_amt_series) != len(source_df):
        order_amt_series = pd.Series(0.0, index=source_df.index)
    else:
        order_amt_series.index = source_df.index

    sales_value_series = ship_amt_series.where(ship_amt_series.ne(0), order_amt_series).fillna(0.0)
    order_value_series = order_amt_series.where(order_amt_series.ne(0), ship_amt_series).fillna(0.0)

    # WIP 狀態
    if wip_col_local:
        wip_series = get_series_by_col(source_df, wip_col_local).fillna("").astype(str).str.strip().str.upper()
    else:
        wip_series = pd.Series("", index=source_df.index)

    is_shipment = wip_series.eq("SHIPMENT")
    is_cancelled = wip_series.str.contains(r"CANCEL|取消", na=False)
    is_hold = wip_series.str.contains(r"HOLD|異常", na=False)

    # 月份統計
    ship_periods = effective_ship_dates.dt.to_period("M")
    order_periods = order_dates.dt.to_period("M")
    all_periods = sorted(set(ship_periods.dropna().tolist()) | set(order_periods.dropna().tolist()), reverse=True)

    if not all_periods:
        st.warning("找不到有效的日期資料。")
        return

    current_period = pd.Period(_dt.datetime.now().strftime("%Y-%m"), freq="M")
    default_idx = all_periods.index(current_period) if current_period in all_periods else 0

    selected_period = st.selectbox(
        "📅 選擇統計月份",
        all_periods,
        index=default_idx,
        format_func=lambda p: f"{p.year}年{p.month}月",
        key="sales_detail_month_teable",
    )

    # 篩選條件
    ship_month_mask = (ship_periods == selected_period).fillna(False)
    order_month_mask = (order_periods == selected_period).fillna(False)

    shipped_mask = is_shipment & ship_month_mask & (~is_cancelled)
    forecast_mask = (~is_shipment) & ship_month_mask & (~is_cancelled)
    order_mask = order_month_mask & (~is_cancelled)
    hold_mask = is_hold & (~is_cancelled)
    month_total_mask = shipped_mask | forecast_mask

    # 計算金額
    shipped_usd = float(sales_value_series[shipped_mask].sum())
    forecast_usd = float(sales_value_series[forecast_mask].sum())
    order_usd = float(order_value_series[order_mask].sum())
    total_usd = shipped_usd + forecast_usd
    hold_usd = float(order_value_series[hold_mask].sum())

    # NTD 計算
    shipped_ntd_10k = _usd_to_ntd_10k(shipped_usd)
    order_ntd_10k = _usd_to_ntd_10k(order_usd)

    # 工廠統計
    factory_order_ntd_10k = 0.0
    if factory_col_local and month_total_mask.any():
        factory_df = pd.DataFrame({
            "工廠": get_series_by_col(source_df, factory_col_local)[month_total_mask].apply(_clean_factory),
            "金額": sales_value_series[month_total_mask],
        })
        factory_total_usd = float(factory_df["金額"].sum())
        factory_order_ntd_10k = _usd_to_ntd_10k(factory_total_usd)

    # 顯示標題
    today_str = pd.Timestamp.now().strftime("%Y/%m/%d")
    st.markdown(
        f"<h3 style='text-align:center;margin-bottom:4px;'>{selected_period.month}月 業績明細表</h3>"
        f"<p style='text-align:right;color:gray;margin-top:0;'>{today_str}</p>",
        unsafe_allow_html=True,
    )

    # 指標卡片 - 按照截图格式（6列）
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("接單金額 (USD)", f"{order_usd:,.2f}" if order_usd else "—")
    c2.metric("接單台幣(NTD:萬)", f"{order_ntd_10k:,.2f}" if order_ntd_10k else "—")
    c3.metric("已出貨 (NTD:萬)", f"{shipped_ntd_10k:,.2f}" if shipped_ntd_10k else "—")
    c4.metric("HOLD (USD)", f"{hold_usd:,.2f}" if hold_usd else "0")
    c5.metric("下單工廠(NTD:萬)", f"{factory_order_ntd_10k:,.2f}" if factory_order_ntd_10k else "—")
    c6.metric("美金出貨 (USD)", f"{total_usd:,.2f}" if total_usd else "—")

    st.markdown("---")

    # HOLD 訂單明細
    if hold_mask.any():
        st.markdown("**HOLD:**")
        hold_df = pd.DataFrame(index=source_df.index[hold_mask])
        if customer_col_local:
            hold_df["客戶"] = get_series_by_col(source_df, customer_col_local)[hold_mask]
        if po_col_local:
            hold_df["PO#"] = get_series_by_col(source_df, po_col_local)[hold_mask]
        if pn_col_local:
            hold_df["P/N"] = get_series_by_col(source_df, pn_col_local)[hold_mask]
        if wip_col_local:
            hold_df["WIP"] = wip_series[hold_mask]
        hold_df["金額(USD)"] = order_value_series[hold_mask].map(_fmt_usd)
        st.dataframe(hold_df, use_container_width=True, hide_index=True, height=150)

    st.markdown("---")

    # 圖表區域
    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("**預出貨(NTD:萬) 依據 月**")
        if month_total_mask.any():
            monthly_df = pd.DataFrame({
                "月份": ship_periods[month_total_mask],
                "金額(NTD:萬)": sales_value_series[month_total_mask].apply(lambda x: _usd_to_ntd_10k(x)),
            }).dropna()
            if not monthly_df.empty:
                monthly_grp = monthly_df.groupby("月份", as_index=False)["金額(NTD:萬)"].sum()
                monthly_grp["月份"] = monthly_grp["月份"].astype(str)
                st.bar_chart(monthly_grp.set_index("月份"), height=200)

    with col_right:
        st.markdown("**預出貨(USD) 依據 客戶分類**")
        if month_total_mask.any() and customer_col_local:
            customer_df = pd.DataFrame({
                "客戶": get_series_by_col(source_df, customer_col_local)[month_total_mask].apply(_clean_customer),
                "金額(USD)": sales_value_series[month_total_mask],
            })
            customer_grp = customer_df.groupby("客戶", as_index=False)["金額(USD)"].sum()
            st.bar_chart(customer_grp.set_index("客戶"), height=200)

    # 台币出货依据日 - 折线图
    st.markdown("**台幣出貨(NTD:萬) 依據 日**")
    if month_total_mask.any():
        daily_df = pd.DataFrame({
            "日期": effective_ship_dates[month_total_mask],
            "金額(NTD:萬)": sales_value_series[month_total_mask].apply(lambda x: _usd_to_ntd_10k(x)),
        }).dropna(subset=["日期"])
        if not daily_df.empty:
            daily_grp = daily_df.groupby(daily_df["日期"].dt.normalize(), as_index=False)["金額(NTD:萬)"].sum()
            daily_grp = daily_grp.sort_values("日期")
            daily_grp["日期"] = pd.to_datetime(daily_grp["日期"]).dt.strftime("%m/%d")
            st.line_chart(daily_grp.set_index("日期"), height=250)

    # 簽名欄
    st.markdown("---")
    sig_col1, sig_col2, sig_col3, sig_col4 = st.columns(4)
    with sig_col1:
        st.text_input("核准:", key="approve_sig")
    with sig_col2:
        st.text_input("覆核2:", key="review2_sig")
    with sig_col3:
        st.text_input("覆核1:", key="review1_sig")
    with sig_col4:
        st.text_input("初核:", key="check_sig", value="Jacky")
        st.text_input("製表:", key="prepare_sig", value="Shirley")
