
# 台灣50正2再平衡模擬器 v1.3

## v1.3 修正：2009以前的市場資料

v1.2 主要依賴 Yahoo Finance 建立歷史資料庫；部分環境可能只抓得到
2009 年以後的 `^TWII`，因此 2000 / 2008 股災壓力測試會缺資料。

v1.3 改為：

- **TAIEX（大盤）1999年至今：證交所官方歷史資料優先補齊**
- 0050：Yahoo Finance（主要保留 Adj Close / 還原價格）
- 00631L：Yahoo Finance（主要保留 Adj Close / 還原價格）
- 一般回測仍完全使用 GitHub 內的本地 `data/market_daily.csv`
- 只有資料管理頁更新時才連網

證交所官方 TAIEX 來源：
`https://www.twse.com.tw/rwd/zh/TAIEX/MI_5MINS_HIST`

## 壓力測試資料分界

### 2000 網路泡沫
使用 **證交所官方 TAIEX 2000–2002**，
再套用 2014 年以來 00631L 實際偏差模型產生「歷史模擬正2」。

### 2008 金融海嘯
改用 **證交所官方 TAIEX 2007–2009**。
不再依賴 0050 的 2008 年 Yahoo 歷史資料，避免資料源不完整。

### 2014/10/31 以後
使用實際：
- 0050
- 00631L

因此「實際資料回測」與「歷史股災壓力測試」會清楚分開。

## 更新方式

把下列檔案覆蓋到 GitHub：

```text
app.py
requirements.txt
README.md
data/market_daily.csv
```

部署後進入：

`⑤ 資料管理`

如果目前 TAIEX 最早日期顯示 2009：

按：

`補齊1999年至今官方TAIEX`

程式只下載本地資料庫缺少的月份，不會重複下載已存在的月份。

補齊完成後，下載最新版：

`market_daily.csv`

並覆蓋回 GitHub 的：

`data/market_daily.csv`

Commit 後即可永久保存。

## 注意

證交所 TAIEX 歷史端點是逐月資料，因此程式採：
- 缺月份才下載
- 失敗自動重試
- 429 限流自動等待
- 已存在月份不重複抓

Streamlit Cloud 本機檔案仍不是永久儲存空間；
補齊完成後記得把最新版 CSV Commit 回 GitHub。
