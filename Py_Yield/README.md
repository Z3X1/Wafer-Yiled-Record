# Wafer Yield Analyzer

一個用於分析晶圓良率的 Python 程式，自動讀取包含 "Wafer_Summary" 的 Excel 檔案，提取良率數據並生成視覺化報告。

## 功能特點

- 🔍 自動搜尋指定目錄下所有包含 "Wafer_Summary" 的 Excel 檔案
- 📊 從每個檔案提取 Wafer ID（B4儲存格）和良率（D11儲存格）
- 📈 生成美觀的折線圖，展示各晶圓的良率趨勢
- 💾 輸出包含數據和圖表的 Excel 報告

## 安裝依賴

使用 pip 安裝所需套件：

```bash
pip install -r requirements.txt
```

或使用 Poetry：

```bash
poetry install
```

## 使用方法

### 1. 直接執行

```bash
python wafer_yield_analyzer.py
```

### 2. 自訂路徑

修改程式中的 `source_directory` 變數：

```python
source_directory = r"C:\Users\andrel52"  # 修改為您的路徑
```

## 輸出檔案

執行後會生成兩個檔案：

1. **wafer_yield_report.xlsx** - 包含：
   - "Yield Data" 工作表：完整的良率數據表格
   - "Yield Chart" 工作表：嵌入的視覺化圖表

2. **wafer_yield_chart.png** - 高解析度的折線圖圖片（300 DPI）

## Excel 檔案格式要求

程式假設您的 Excel 檔案具有以下格式：
- **B4 儲存格**：Wafer ID（晶圓識別碼）
- **D11 儲存格**：Yield（良率百分比）

## 視覺化特點

- 清晰的折線圖，帶有標記點
- 每個數據點顯示精確的良率值
- 自動調整 Y 軸範圍以最佳化顯示
- 專業的配色方案和網格線
- 支援多個晶圓數據的對比分析

## 錯誤處理

程式包含完善的錯誤處理機制：
- 自動跳過無法讀取的檔案
- 驗證必要儲存格的數據
- 詳細的日誌記錄

## 系統需求

- Python 3.10+
- Windows/Linux/macOS
- 足夠的磁碟空間存放輸出檔案

## 授權

MIT License

