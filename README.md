# 台灣50正2再平衡模擬器 v2.0

## 新增：Fear & Greed 五區間「轉換路徑」研究

除了原本單純研究：
- Extreme Fear
- Fear
- Neutral
- Greed
- Extreme Greed

現在會另外研究「從哪一區進入哪一區」。

例如：
- Fear → Neutral
- Greed → Neutral
- Extreme Fear → Fear
- Neutral → Greed
- Greed → Extreme Greed

### 統計規則
只在 F&G 真正跨區間的那一天記一筆。
連續待在同一區間不會重複計算，因此不會把 Neutral → Neutral 當成事件。

若數值一次直接跨兩級，例如 Extreme Fear → Neutral，
會照實記成 Extreme Fear → Neutral，不強迫拆成兩筆。

## 5×5轉換矩陣
列 = 從哪個區間來（From）
欄 = 進入哪個區間（To）

可切換顯示：
- 樣本數
- 後5日平均報酬
- 後20日平均報酬
- 後60日平均報酬
- 後120日平均報酬
- 後20日上漲率
- 後60日上漲率
- 後120日上漲率

對角線空白，因為同區間停留不算轉換。

## 詳細統計
每一種 From → To 顯示：
- 樣本數
- 進入後 F&G 平均值
- 後5/20/60/120日平均報酬
- 後5/20/60/120日中位數
- 後5/20/60/120日上漲率

可以設定最低樣本數，避免只看一兩次偶然事件。

## 指定到達區間比較
可指定例如 Neutral，
直接比較：
- Extreme Fear → Neutral
- Fear → Neutral
- Greed → Neutral
- Extreme Greed → Neutral

這正是用來回答「同樣進入Neutral，但它是從恐懼方向上來，還是從貪婪方向下來，後續報酬是否不同」。

## 資料來源可分開研究
- 全部（2011年至今）
- 僅第三方重建（至2021/01）
- 僅CNN端點尾端（2021/02後）

## 事件明細
可以指定某一種轉換，查看每次實際發生日，以及：
- 使用的 F&G 來源日
- 當時 F&G 值
- 台股當日值
- 後5/20/60/120日報酬

## 判讀提醒
5區間有20種非同區轉換，再乘上多個持有期間，容易出現多重比較問題。
不要只看平均報酬；至少同時檢視：
- 樣本數
- 中位數
- 上漲率
- reconstructed / CNN端點資料是否一致
