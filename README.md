# GLOCOM Factory PO Generator v3

> Phase 1A+1B 完整整合版,接 GLOCOM Control Tower (`glocom-wip-v36`)
> 
> 重大改變(v1/v2 → v3):
> - 不再有 customers.json 與 SQLite,所有資料來源 = Teable 主表
> - 從「西拓訂單編號」選擇開單,單號類型自動推導(ET/G → GLOCOM,EW → EUSWAY)
> - 支援多品項(主表一個訂單編號可能有 N 個 P/N)
> - 產出後可一鍵回寫 Teable 主表的 `Factory PO PDF` 欄位

## 已驗證

用您 2852 row 的真實主表 CSV 跑端到端測試:

- ✓ 列出 2003 個西拓訂單編號可選
- ✓ G1150014-01 撈出 3 個 P/N (VORNE: 90-0341-00, 90-0329-02, 90-0333-01)
- ✓ 產出 docx + PDF,3 個品項都正確
- ✓ 合計 NT$ 5,200.00 跟主表「銷貨金額」加總一致
- ✓ 訂單號自動解析:ET/EW/G 各對應正確的 issuing_company

## 整合說明

請看 `docs/INTEGRATION.md` — 包含:
- 檔案複製清單
- app.py 修改的 3 處(逐字標明改在哪)
- Teable 加新欄位的步驟
- 第一次執行步驟
- 常見問題排查

## 目錄結構

```
glocom_factory_po_v3/
├── factory_po_page.py             ← Streamlit 頁面入口(整合到 app.py 用)
├── core/
│   ├── factory_master.py          ← 工廠主檔載入(factories.json)
│   ├── teable_query.py            ← 從主表 DataFrame 撈 PO 資料
│   ├── teable_writeback.py        ← 回寫 Factory PO PDF 欄位到 Teable
│   └── pdf_generator.py           ← docx + PDF 產生(支援多品項)
├── data/
│   └── factories.json             ← 8 家工廠主檔(4 家完整,4+ 家空殼待補)
├── templates/
│   ├── PO_GLOCOM.docx             ← Word 母版
│   └── build_po_glocom_template.py ← 母版生成腳本
├── output/                        ← 產出 PDF 放這裡
├── docs/
│   └── INTEGRATION.md             ← 整合到 Control Tower 的具體步驟
├── tests/
│   └── test_end_to_end_v3.py      ← 端到端測試
├── requirements.txt
└── README.md (this file)
```

## Phase 2 待辦

跑順之後,以下是可以回頭優化的項目(都不阻塞當前使用):

1. **空殼工廠補完整資料** — 全興、優技、祥竑、柏承、台豐、FCF
2. **EUSWAY template** — `templates/PO_EUSWAY.docx`,給 EW 三角貿易單用
3. **Logo 圖替換** — 目前 template 上是「GLOCOM」文字佔位,可換成實際 logo
4. **PDF URL 上雲** — 目前回寫到 Teable 的是本機路徑,可改成上傳 pCloud / S3 後的 URL
5. **多工廠拆單** — 目前假設一個西拓編號 = 一張 PDF,若實際會出現
   「同編號拆多家工廠」的情況再改
```
