"""
app.py 修改說明 — 只需改兩處，各加一行
════════════════════════════════════════════

① 找到這段（約第 20 行，import 區塊末尾）：

    from reports import (
        show_new_orders_wip_report,
        ...
    )

   在它「之後」加入：

    from signflow_approval import render_approval_page


② 找到這段（約 app.py 末尾附近）：

    menu = st.sidebar.radio(
        "功能選單",
        [
            "Dashboard",
            "Factory Load",
            "Delayed Orders",
            "Shipment Forecast",
            "Orders",
            "新訂單 WIP",
            "Sandy 內部 WIP",
            "Sandy 銷貨底",
            "業績明細表",
            "Customer Preview",
            "Import / Update",
        ]
    )

   改成（只加最後那一行）：

    menu = st.sidebar.radio(
        "功能選單",
        [
            "Dashboard",
            "Factory Load",
            "Delayed Orders",
            "Shipment Forecast",
            "Orders",
            "新訂單 WIP",
            "Sandy 內部 WIP",
            "Sandy 銷貨底",
            "業績明細表",
            "Customer Preview",
            "Import / Update",
            "✍ 簽核平台",        # ← 新增這行
        ]
    )


③ 找到 app.py 最末尾：

    # Excel Quote Export removed from menu.

   在它「之前」加入：

    elif menu == "✍ 簽核平台":
        render_approval_page()


完成！共改 3 處，不動其他任何程式碼。
"""
