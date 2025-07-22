# 米斯特 門市月考核查詢平台

📊 本專案為米斯特內部「門市月考核查詢系統」之 Streamlit App，支援多欄位查詢、考核分布展示、結果匯出等功能。

## 🔧 使用說明

### 1️⃣ 功能包含：
- 區主管、部門編號、員工編號、人員姓名查詢
- 本月考核等級分布（分頁：等級分布）
- 四大資料顯示：
  - 門店考核總表
  - 人效分析
  - 店長/副店 考核明細
  - 店員/儲備 考核明細
- 📥 匯出查詢結果（ZIP壓縮檔內含4張CSV）

### 2️⃣ 資料來源
App 使用公開的 Google Sheet Excel 檔：
[2025.06_MST-PA.xlsx](https://raw.githubusercontent.com/ainstaccc/kpi-checker/main/2025.06_MST-PA.xlsx)

## 🚀 如何執行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 🛠 Hugging Face Space 部署須知

- 本 App 為 Streamlit 類型
- `runtime.txt` 指定 Python 版本為 3.10
  
