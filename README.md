# 🚢 船期整理系統 - Shipping Schedule Organizer

> POL: **HAIPHONG** → POD: **HONG KONG / SHEKOU / KAOHSIUNG / TAICHUNG**

---

## 支援船公司

| 船公司 | 資料來源 | 格式 |
|--------|---------|------|
| **CNC** | 上傳檔案 | CSV / PDF / Excel |
| **TSL** | 上傳檔案 | Excel / PDF |
| **IAL** | 網頁爬取 | https://www.interasia.cc |
| **KMTC** | 網頁爬取 | https://www.ekmtc.com |
| **YML** | 網頁爬取 | https://www.yangming.com |

---

## 輸出欄位

| 欄位 | 說明 |
|------|------|
| POL | 起運港（HAIPHONG） |
| POD | 目的港 |
| Vessel | 船名 |
| Voyage | 航次 |
| ETD | 出發日期 (YYYY/MM/DD) |
| ETA | 到達日期 (YYYY/MM/DD) |
| T/T Time | 運輸天數 |
| CY Cut-off | 貨物截關日期時間 |
| SI Cut-off | 文件截止日期時間 |

---

## 本地安裝與執行

```bash
# 1. 建立虛擬環境（建議）
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate

# 2. 安裝套件
pip install -r requirements.txt

# 3. 安裝 Chromium（僅首次需要）
# macOS: brew install chromium
# Ubuntu: sudo apt-get install chromium-browser chromium-chromedriver
# Windows: 下載 ChromeDriver https://chromedriver.chromium.org/

# 4. 執行 APP
streamlit run app.py
```

---

## 部署到 Streamlit Cloud

1. 將專案上傳到 GitHub（包含 `app.py`, `requirements.txt`, `packages.txt`）
2. 前往 https://streamlit.io/cloud → New App
3. 選擇您的 repo 和 `app.py`
4. 點擊 Deploy

> **注意**: `packages.txt` 會自動安裝 `chromium` 和 `chromium-driver`，無需額外設定。

---

## 使用流程

### 方式一：上傳檔案（CNC / TSL）
1. 在側邊欄選擇**年份**和**月份**
2. 勾選所需的**目的港 (POD)**
3. 前往「📂 上傳檔案」標籤
4. 上傳對應船公司的檔案（檔名需含 CNC 或 TSL）
5. 點擊「解析」按鈕

### 方式二：網頁爬取（IAL / KMTC / YML）
1. 在側邊欄選擇**年份**和**月份**
2. 勾選所需的**目的港 (POD)**
3. 前往「🌐 網頁爬取」標籤
4. 個別點擊各船公司的「爬取」按鈕，或使用「一鍵爬取全部」

### 匯出
1. 前往「📋 資料預覽」確認資料正確
2. 前往「📥 匯出 Excel」
3. 點擊「生成 Excel 報表」
4. 下載 Excel 檔案

---

## TSL CY/SI Cut-off 計算邏輯

TSL 的截關日期透過「服務時刻表」的星期推算：

- **CY Cut-off**：從 ETD 日期往前推算至指定星期幾，加上 CY 欄位的時間
  - 例：ETD = 週一 2026/02/02，CY 欄位 = `24:00 SAT` → CY Cut-off = **2026/01/31 24:00**
- **SI Cut-off**：同理，從 ETD 往前推算至 SI/VGM 欄位的星期幾
  - 例：ETD = 週一 2026/02/02，SI 欄位 = `09:00 FRI` → SI Cut-off = **2026/01/30 09:00**

---

## 常見問題

**Q: 爬取時顯示「無法啟動瀏覽器驅動」**  
A: 請確認已安裝 `chromium` 和對應版本的 `chromedriver`

**Q: CNC CSV 解析不到資料**  
A: 確認 CSV 欄位名稱包含：`Origin`, `Destination`, `Vessel name`, `Departure Date` 等

**Q: TSL CY Cut-off 顯示空白**  
A: Excel 中需要有「SERVICE NAME」參考表，包含各服務的 ETD 星期和截關時間

---

## 版本資訊

- v1.0 — 支援 CNC, TSL, IAL, KMTC, YML
- POL: HAIPHONG | POD: HONG KONG, SHEKOU, KAOHSIUNG, TAICHUNG
