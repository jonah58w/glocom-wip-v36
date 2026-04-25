# GLOCOM Control Tower 整合說明 — Factory PO Generator

> **目標**:把 Factory PO Generator 整合進 `glocom-wip-v36` 的 `app.py`,
> 新增「Factory PO」選單項目。
> 
> **不影響現有功能**:Import / Update、Dashboard、新訂單 WIP、業績明細表 等全部保留。

## 一、檔案複製清單

把這幾個資料夾 / 檔案,從 zip 解壓後**整個放到** `glocom-wip-v36` 的根目錄(跟 `app.py` 同層):

```
glocom-wip-v36/
├── app.py                     ← 你的現有檔(下面要改 3 處)
├── teable_api.py              ← 現有
├── factory_progress_updater.py ← 現有
├── ... (其他現有檔)
│
├── factory_po_page.py         ← 新增(從 zip 帶過來)
├── core/                      ← 新增資料夾
│   ├── __init__.py
│   ├── factory_master.py
│   ├── teable_query.py
│   ├── teable_writeback.py
│   └── pdf_generator.py
├── data/
│   └── factories.json         ← 新增(8 家工廠主檔,空殼待補)
├── templates/
│   ├── PO_GLOCOM.docx         ← 新增
│   └── build_po_glocom_template.py
└── output/                    ← 新增(空目錄,產出 PDF 用)
```

> **注意**:如果 `glocom-wip-v36` 已經有 `data/` 或 `templates/` 資料夾,**合併進去**,不要覆蓋既有檔案。

## 二、套件需求

`requirements.txt` 加這幾行(如果還沒有):

```
docxtpl>=0.16
python-docx>=1.0
docx2pdf>=0.1.8 ; sys_platform == "win32"
```

Linux/Mac(部署在 Streamlit Cloud 等):需要 `libreoffice`,在 `packages.txt` 加:

```
libreoffice
```

(您的 `packages.txt` 已經有 `tesseract-ocr` 等,加一行 `libreoffice` 即可)

## 三、app.py 修改 — 共 3 處

### 修改 ①:menu 加項目(在 `menu = st.sidebar.radio(...)` 區塊)

**位置**:`app.py` 大約 line 870 附近,`menu = st.sidebar.radio(...)` 那段。

**找到這段**:

```python
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
        "SignFlow",
    ]
)
```

**改成**(在 `Import / Update` 下面加一行 `Factory PO`):

```python
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
        "Factory PO",          # ← 新增這一行
        "SignFlow",
    ]
)
```

### 修改 ②:menu dispatch 加 elif(在頁面分派區塊)

**位置**:`app.py` 最尾段,`elif menu == "SignFlow":` 之前。

**找到這段**:

```python
elif menu == "SignFlow":
    render_approval_page()
```

**改成**(在前面加一段):

```python
elif menu == "Factory PO":
    from factory_po_page import render_factory_po_page
    render_factory_po_page(orders, TABLE_URL, HEADERS)

elif menu == "SignFlow":
    render_approval_page()
```

### 修改 ③:**(可選)** 如果想在 Debug 區塊看 Factory PO PDF 欄位

**位置**:`app.py` 的 Debug expander。

**找到這段**(line 約 750):

```python
with st.expander("Debug"):
    st.write("API Status:", api_status)
    st.write("TABLE_URL:", TABLE_URL)
    ...
```

**可選加一行**(確認新欄位有讀進來):

```python
    st.write("Factory PO PDF 欄位存在:", "Factory PO PDF" in orders.columns)
```

## 四、Teable 主表新增欄位

**欄位名稱**:`Factory PO PDF`
**類型**:URL(若無此選項則用 Single Line Text)
**位置**:加在「西拓訂單編號」右邊
**作用**:產生 PDF 時,系統會把 PDF 路徑回寫到這個欄位

## 五、第一次執行步驟

```bash
# 1. 安裝依賴
pip install -r requirements.txt

# 2. 第一次跑要先生成 template(這個只需做一次,除非要改版面)
python templates/build_po_glocom_template.py

# 3. 啟動(跟原本一樣)
streamlit run app.py
```

## 六、驗證可不可以用

1. 啟動 Streamlit
2. 左側選單應該多了「Factory PO」項目
3. 點進去,應該看到下拉選單,可以選任何「西拓訂單編號」(預設顯示主表共 2003 個編號可選)
4. 選一個有實際資料的(例如 `G1150014-01`)→ 會顯示 3 個品項(VORNE 那筆)
5. 點「產生 PDF (不回寫)」→ 應該下載到 docx + PDF
6. PDF 打開應該長得像鉅盛訂購單格式

## 七、第一次跑可能遇到的問題

### 問題 1:工廠資料是空殼

**症狀**:PDF 上廠商地址、電話顯示 `[請補]`
**解法**:編輯 `data/factories.json`,把對應工廠的欄位填完整。改完不用重啟 Streamlit,下一次點「產生」就會用新資料。

### 問題 2:`Factory PO PDF` 欄位回寫失敗

**症狀**:點「產生 + 回寫」後出現 `404 / Field not found`
**可能原因**:
- Teable 上欄位名稱拼寫不一致(我預設用 `Factory PO PDF`,空格要對)
- 您加的欄位類型不接受字串

**解法**:打開 `core/teable_writeback.py`,改 `PDF_FIELD_NAME = "Factory PO PDF"` 成您 Teable 上實際的名稱。

### 問題 3:LibreOffice 沒裝(Streamlit Cloud)

**症狀**:docx 產出但 PDF 失敗,訊息「LibreOffice 轉換失敗」
**解法**:`packages.txt` 加 `libreoffice` 後重新部署。docx 仍可下載,只是少了 PDF。

## 八、之後可能要回頭調整的地方

1. **factories.json 補完整**:7 家工廠目前是空殼,逐步補
2. **PDF URL**:第一版是寫本機路徑,之後若要 https URL,改 `factory_po_page.py` 中
   `pdf_url_to_write = ...` 那行,改成上傳到 pCloud / S3 後的 URL
3. **Logo 圖**:目前 template 只有「GLOCOM」文字佔位,可在 Word 開
   `templates/PO_GLOCOM.docx` 換成實際 Logo 圖
4. **EUSWAY 抬頭 template**:目前只有 GLOCOM 抬頭,EW 號(三角貿易)也會用 GLOCOM template。
   之後要做 `templates/PO_EUSWAY.docx`(版面同 GLOCOM,但抬頭換成 EUSWAY)
