
from pathlib import Path
from datetime import date
import io

import numpy as np
import pandas as pd
import streamlit as st
import yfinance as yf

st.set_page_config(page_title="台灣50正2再平衡模擬器 v1.2", layout="wide")

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DB_FILE = DATA_DIR / "market_daily.csv"

SYMBOLS = {
    "TAIEX": {
        "ticker": "^TWII",
        "start": "1999-01-01",
        "label": "台灣加權指數",
    },
    "0050": {
        "ticker": "0050.TW",
        "start": "2003-06-30",
        "label": "元大台灣50",
    },
    "00631L": {
        "ticker": "00631L.TW",
        "start": "2014-10-31",
        "label": "元大台灣50正2",
    },
}

DB_COLUMNS = ["Date", "Symbol", "Close", "AdjClose"]


# ============================================================
# 本地市場資料庫
# ============================================================
def empty_database():
    return pd.DataFrame(columns=DB_COLUMNS)


@st.cache_data(show_spinner=False)
def read_local_database(path_str: str, modified_ns: int = 0):
    path = Path(path_str)
    if not path.exists() or path.stat().st_size == 0:
        return empty_database()

    try:
        df = pd.read_csv(path)
    except pd.errors.EmptyDataError:
        return empty_database()

    if df.empty:
        return empty_database()

    missing = set(DB_COLUMNS) - set(df.columns)
    if missing:
        raise RuntimeError(f"本地資料庫缺少欄位：{', '.join(sorted(missing))}")

    df = df[DB_COLUMNS].copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Symbol"] = df["Symbol"].astype(str)
    df["Close"] = pd.to_numeric(df["Close"], errors="coerce")
    df["AdjClose"] = pd.to_numeric(df["AdjClose"], errors="coerce")
    df = (
        df.dropna(subset=["Date", "Symbol"])
          .drop_duplicates(["Date", "Symbol"], keep="last")
          .sort_values(["Symbol", "Date"])
          .reset_index(drop=True)
    )
    return df


def get_local_database():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    modified = DB_FILE.stat().st_mtime_ns if DB_FILE.exists() else 0
    return read_local_database(str(DB_FILE), modified)


def save_local_database(df: pd.DataFrame):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    out = df[DB_COLUMNS].copy()
    out["Date"] = pd.to_datetime(out["Date"]).dt.strftime("%Y-%m-%d")
    out = (
        out.drop_duplicates(["Date", "Symbol"], keep="last")
           .sort_values(["Symbol", "Date"])
    )
    out.to_csv(DB_FILE, index=False, encoding="utf-8-sig")
    st.cache_data.clear()


def merge_database(old: pd.DataFrame, new: pd.DataFrame):
    if old is None or old.empty:
        merged = new.copy()
    elif new is None or new.empty:
        merged = old.copy()
    else:
        merged = pd.concat([old, new], ignore_index=True)

    if merged.empty:
        return empty_database()

    merged["Date"] = pd.to_datetime(merged["Date"], errors="coerce")
    merged["Close"] = pd.to_numeric(merged["Close"], errors="coerce")
    merged["AdjClose"] = pd.to_numeric(merged["AdjClose"], errors="coerce")
    return (
        merged.dropna(subset=["Date", "Symbol"])
              .drop_duplicates(["Date", "Symbol"], keep="last")
              .sort_values(["Symbol", "Date"])
              .reset_index(drop=True)
    )


@st.cache_data(ttl=1800, show_spinner=False)
def download_yahoo_long(symbol_key: str, start: str, end: str):
    info = SYMBOLS[symbol_key]
    end_dt = pd.Timestamp(end) + pd.Timedelta(days=1)

    raw = yf.download(
        info["ticker"],
        start=start,
        end=end_dt.strftime("%Y-%m-%d"),
        auto_adjust=False,
        progress=False,
        actions=False,
        threads=False,
    )
    if raw is None or raw.empty:
        raise RuntimeError(f"{info['ticker']} 無法取得資料")

    def get_col(name):
        if isinstance(raw.columns, pd.MultiIndex):
            if name not in raw.columns.get_level_values(0):
                return None
            x = raw[name]
            if isinstance(x, pd.DataFrame):
                x = x.iloc[:, 0]
            return x
        if name not in raw.columns:
            return None
        return raw[name]

    close = get_col("Close")
    adj = get_col("Adj Close")
    if close is None:
        raise RuntimeError(f"{info['ticker']} 缺少 Close 欄位")
    if adj is None:
        adj = close.copy()

    out = pd.DataFrame({
        "Date": pd.to_datetime(close.index).tz_localize(None),
        "Symbol": symbol_key,
        "Close": pd.to_numeric(close.values, errors="coerce"),
        "AdjClose": pd.to_numeric(adj.values, errors="coerce"),
    })
    return out.dropna(subset=["Date", "Close"]).reset_index(drop=True)


def database_status(db: pd.DataFrame):
    rows = []
    for key, info in SYMBOLS.items():
        x = db.loc[db["Symbol"] == key].copy() if not db.empty else empty_database()
        rows.append({
            "資料": info["label"],
            "代號": key,
            "起始日": x["Date"].min().date().isoformat() if not x.empty else "尚未建立",
            "最後日期": x["Date"].max().date().isoformat() if not x.empty else "尚未建立",
            "交易日筆數": int(x["Date"].nunique()) if not x.empty else 0,
        })
    return pd.DataFrame(rows)


def update_database(db: pd.DataFrame, full_rebuild=False, end=None):
    end = end or date.today().isoformat()
    all_new = []
    progress = st.progress(0)

    for i, (key, info) in enumerate(SYMBOLS.items(), start=1):
        if full_rebuild or db.empty:
            start = info["start"]
        else:
            x = db.loc[db["Symbol"] == key]
            if x.empty:
                start = info["start"]
            else:
                # 重抓最後14日，修正晚到資料或資料源微調
                start = (x["Date"].max() - pd.Timedelta(days=14)).strftime("%Y-%m-%d")

        new = download_yahoo_long(key, start, end)
        all_new.append(new)
        progress.progress(i / len(SYMBOLS))

    progress.empty()
    fresh = pd.concat(all_new, ignore_index=True) if all_new else empty_database()

    if full_rebuild:
        out = fresh
    else:
        out = merge_database(db, fresh)

    save_local_database(out)
    return out


# ============================================================
# 資料處理
# ============================================================
def symbol_frame(db: pd.DataFrame, symbol: str):
    x = db.loc[db["Symbol"] == symbol, ["Date", "Close", "AdjClose"]].copy()
    if x.empty:
        return pd.DataFrame(columns=["Close", "AdjClose"])
    x = x.drop_duplicates("Date", keep="last").set_index("Date").sort_index()
    return x


def repair_split_jumps(price: pd.Series, threshold=0.30):
    """
    防止資料源未調整 ETF 分割時，把分割誤判成股災。
    若單日大跌超過 threshold 且前價/後價接近整數倍，
    將分割日前的歷史價格按該倍數調整。
    """
    s = price.astype(float).copy().dropna().sort_index()
    events = []

    for _ in range(10):
        r = s.pct_change()
        bad = r[r < -threshold]
        fixed = False

        for dt, ret in bad.items():
            i = s.index.get_loc(dt)
            if i <= 0:
                continue
            prev = float(s.iloc[i - 1])
            curr = float(s.iloc[i])
            if curr <= 0:
                continue

            raw_ratio = prev / curr
            ratio = int(round(raw_ratio))
            if ratio >= 2:
                gap = curr * ratio / prev - 1
                if abs(gap) <= 0.20:
                    s.iloc[:i] = s.iloc[:i] / ratio
                    events.append({
                        "日期": pd.Timestamp(dt),
                        "推定分割比率": f"{ratio}:1",
                        "原始單日變動": float(ret),
                    })
                    fixed = True
                    break
        if not fixed:
            break

    return s, events


def prepare_actual_data(db: pd.DataFrame):
    d0050 = symbol_frame(db, "0050")
    d2x = symbol_frame(db, "00631L")
    twii = symbol_frame(db, "TAIEX")

    if d0050.empty or d2x.empty:
        raise RuntimeError("資料庫必須至少包含0050與00631L，才能執行實際回測。")

    start = max(pd.Timestamp("2014-10-31"), d0050.index.min(), d2x.index.min())
    common = d0050.loc[start:].index.intersection(d2x.loc[start:].index)

    raw0050 = d0050["Close"].reindex(common).dropna()
    raw2x = d2x["Close"].reindex(common).dropna()

    p0050, split0050 = repair_split_jumps(raw0050)
    p2x, split2x = repair_split_jumps(raw2x)

    tr0050 = d0050["AdjClose"].reindex(common).dropna()
    common = p0050.index.intersection(p2x.index).intersection(tr0050.index)
    p0050 = p0050.reindex(common)
    p2x = p2x.reindex(common)
    tr0050 = tr0050.reindex(common)

    r0050_price = p0050.pct_change().fillna(0)
    r0050_total = tr0050.pct_change().fillna(0)
    r2x = p2x.pct_change().fillna(0)

    warnings = []
    if (r0050_price.abs() > 0.30).any():
        warnings.append("0050分割修復後仍有單日價格變動超過30%。")
    if (r2x.abs() > 0.30).any():
        warnings.append("00631L分割修復後仍有單日價格變動超過30%。")

    return {
        "d0050": d0050,
        "d2x": d2x,
        "twii": twii,
        "p0050": p0050,
        "p2x": p2x,
        "tr0050": tr0050,
        "r0050_price": r0050_price,
        "r0050_total": r0050_total,
        "r2x": r2x,
        "split0050": split0050,
        "split2x": split2x,
        "warnings": warnings,
    }


# ============================================================
# 正2偏差模型
# ============================================================
def build_features(r):
    x = pd.DataFrame(index=r.index)
    x["r"] = r
    x["abs_r"] = r.abs()
    x["r2"] = r.pow(2)
    x["vol20"] = r.rolling(20, min_periods=10).std() * np.sqrt(252)
    x["mom20"] = (1 + r).rolling(20, min_periods=10).apply(np.prod, raw=True) - 1
    return x.replace([np.inf, -np.inf], np.nan)


def fit_gap_model(under_ret, lev_ret):
    z = pd.concat([under_ret.rename("u"), lev_ret.rename("l")], axis=1).dropna()
    f = build_features(z["u"]).dropna()
    z = z.loc[f.index]

    y = z["l"] - 2 * z["u"]
    X = np.column_stack([
        np.ones(len(f)),
        f["r"].values,
        f["abs_r"].values,
        f["r2"].values,
        f["vol20"].values,
        f["mom20"].values,
    ])

    coef, *_ = np.linalg.lstsq(X, y.values, rcond=None)
    pred = X @ coef
    eps = y.values - pred

    cal = f.copy()
    cal["gap"] = y.values
    cal["pred_gap"] = pred
    cal["eps"] = eps
    cal["sign"] = np.where(cal["r"] >= 0, 1, -1)
    q = cal["vol20"].quantile([0.25, 0.50, 0.75]).values
    cal["vol_bucket"] = np.digitize(cal["vol20"].values, q, right=True)
    return coef, cal, q


def predict_gap(r, coef):
    f = build_features(r).copy()
    for c in ["vol20", "mom20"]:
        fill = f[c].median()
        if pd.isna(fill):
            fill = 0.0
        f[c] = f[c].fillna(fill)

    X = np.column_stack([
        np.ones(len(f)),
        f["r"].fillna(0).values,
        f["abs_r"].fillna(0).values,
        f["r2"].fillna(0).values,
        f["vol20"].fillna(0).values,
        f["mom20"].fillna(0).values,
    ])
    return pd.Series(X @ coef, index=r.index)


def synthetic_2x_returns(under_ret, coef, calibration, vol_q, mode="校準平均", seed=42):
    base = 2 * under_ret.fillna(0)
    if mode == "理論2倍":
        return base.clip(lower=-0.95, upper=1.50)

    out = base + predict_gap(under_ret.fillna(0), coef)

    if mode == "波動狀態抽樣":
        rng = np.random.default_rng(seed)
        f = build_features(under_ret.fillna(0))
        f["vol20"] = f["vol20"].fillna(calibration["vol20"].median())
        signs = np.where(under_ret.fillna(0).values >= 0, 1, -1)
        buckets = np.digitize(f["vol20"].values, vol_q, right=True)
        all_eps = calibration["eps"].dropna().values
        sampled = np.zeros(len(f))

        for i, (s, b) in enumerate(zip(signs, buckets)):
            pool = calibration.loc[
                (calibration["sign"] == s) &
                (calibration["vol_bucket"] == b), "eps"
            ].dropna().values
            if len(pool) < 20:
                pool = all_eps
            sampled[i] = rng.choice(pool)

        out = out + sampled

    return out.clip(lower=-0.95, upper=1.50)


def make_level(ret, start=100):
    return start * (1 + ret.fillna(0)).cumprod()


# ============================================================
# 策略與績效
# ============================================================
def rebalance_to_target(lev, cash, target, buy_cost, sell_cost):
    total = lev + cash
    desired = total * target

    if desired > lev:
        x = max(0.0, (target * total - lev) / (1 + target * buy_cost))
        x = min(x, cash / (1 + buy_cost))
        cost = x * buy_cost
        lev += x
        cash -= x + cost
    else:
        denom = max(1e-12, 1 - target * sell_cost)
        x = max(0.0, (lev - target * total) / denom)
        x = min(x, lev)
        cost = x * sell_cost
        lev -= x
        cash += x - cost

    return lev, cash, x


def simulate_rebalance(
    lev_ret, lev_level, capital, target,
    trigger_mode, down_trigger, up_trigger, band,
    cash_yield, buy_cost, sell_cost, min_days=0
):
    idx = lev_ret.index.intersection(lev_level.index)
    r = lev_ret.reindex(idx).fillna(0)
    level = lev_level.reindex(idx).ffill()

    lev = capital * target
    cash = capital * (1 - target)
    anchor = float(level.iloc[0])
    last_rebal_i = -10**9
    n_rebal = 0
    turnover_sum = 0.0
    rows = []

    daily_cash = (1 + cash_yield) ** (1 / 252) - 1

    # 第0筆只記錄起始本金，不先套用報酬
    first_dt = idx[0]
    rows.append((first_dt, capital, lev, cash, target, "起始"))

    for i, dt in enumerate(idx[1:], start=1):
        lev *= 1 + float(r.loc[dt])
        cash *= 1 + daily_cash
        total = max(lev + cash, 1e-12)
        weight = lev / total
        rel_price = float(level.loc[dt]) / anchor - 1 if anchor > 0 else 0

        trigger = False
        reason = ""

        if i - last_rebal_i >= min_days:
            if trigger_mode == "價格漲跌門檻":
                if rel_price <= -down_trigger:
                    trigger, reason = True, "下跌門檻"
                elif rel_price >= up_trigger:
                    trigger, reason = True, "上漲門檻"
            else:
                if weight <= target - band:
                    trigger, reason = True, "低於比例帶"
                elif weight >= target + band:
                    trigger, reason = True, "高於比例帶"

        if trigger:
            lev, cash, turnover = rebalance_to_target(
                lev, cash, target, buy_cost, sell_cost
            )
            turnover_sum += turnover
            n_rebal += 1
            last_rebal_i = i
            anchor = float(level.loc[dt])
            total = lev + cash
            weight = lev / total if total > 0 else 0

        rows.append((dt, lev + cash, lev, cash, weight, reason))

    df = pd.DataFrame(
        rows,
        columns=["Date", "Portfolio", "Leveraged", "Cash", "LevWeight", "Reason"]
    ).set_index("Date")
    df["Peak"] = df["Portfolio"].cummax()
    df["Drawdown"] = df["Portfolio"] / df["Peak"] - 1
    return df, n_rebal, turnover_sum


def simulate_benchmark(ret, capital):
    idx = ret.index
    vals = [capital]
    for x in ret.iloc[1:]:
        vals.append(vals[-1] * (1 + float(x)))
    s = pd.Series(vals, index=idx, name="Portfolio")
    df = pd.DataFrame({"Portfolio": s})
    df["Peak"] = s.cummax()
    df["Drawdown"] = s / df["Peak"] - 1
    return df


def metrics(eq, initial_capital, n_rebal=0, turnover=0):
    eq = eq.dropna()
    if len(eq) < 2:
        return {}

    years = max((eq.index[-1] - eq.index[0]).days / 365.25, 1/252)
    total_ret = eq.iloc[-1] / initial_capital - 1
    cagr = (eq.iloc[-1] / initial_capital) ** (1 / years) - 1

    dr = eq.pct_change().dropna()
    vol = dr.std() * np.sqrt(252) if len(dr) else np.nan
    dd = eq / eq.cummax() - 1
    mdd = dd.min()
    sharpe = dr.mean() / dr.std() * np.sqrt(252) if len(dr) and dr.std() > 0 else np.nan
    calmar = cagr / abs(mdd) if mdd < 0 else np.nan

    return {
        "期末資產": eq.iloc[-1],
        "累積報酬": total_ret,
        "CAGR": cagr,
        "年化波動": vol,
        "最大回撤": mdd,
        "Sharpe(無風險=0)": sharpe,
        "Calmar": calmar,
        "再平衡次數": n_rebal,
        "累計換手金額": turnover,
    }


def pareto_frontier(df):
    vals = df[["CAGR", "最大回撤"]].values
    keep = np.ones(len(df), dtype=bool)
    for i in range(len(df)):
        dominated = (
            (vals[:, 0] >= vals[i, 0]) &
            (vals[:, 1] >= vals[i, 1]) &
            ((vals[:, 0] > vals[i, 0]) | (vals[:, 1] > vals[i, 1]))
        )
        if dominated.any():
            keep[i] = False
    return df.loc[keep].copy()


# ============================================================
# UI
# ============================================================
st.title("台灣50正2 × 現金：再平衡與股災壓力測試模擬器 v1.2")
st.caption(
    "v1.2 改為『本地歷史資料庫優先』：一般回測不連Yahoo Finance；"
    "只有在資料管理頁按更新時，才下載缺少的最新交易日。"
)

db = get_local_database()

with st.sidebar:
    st.header("共通設定")
    capital = st.number_input("初始資金（元）", min_value=10_000, value=1_000_000, step=100_000)
    cash_yield_pct = st.number_input("現金年化收益率（%）", value=0.0, step=0.1)
    buy_cost_pct = st.number_input("買進單邊成本（%）", value=0.05, step=0.01)
    sell_cost_pct = st.number_input("賣出單邊成本（%）", value=0.15, step=0.01)
    min_days = st.number_input("兩次再平衡最少間隔交易日", min_value=0, value=0, step=1)

    st.divider()
    st.header("手動策略參數")
    target_pct = st.slider("正2目標比例（%）", 10, 80, 50, 5)
    trigger_mode = st.radio("再平衡觸發方式", ["價格漲跌門檻", "部位比例偏離"])

    if trigger_mode == "價格漲跌門檻":
        down_pct = st.slider("自上次再平衡後，下跌幾%買回目標比例", 5, 60, 30, 5)
        up_pct = st.slider("自上次再平衡後，上漲幾%賣回目標比例", 5, 100, 40, 5)
        band_pct = 7.5
    else:
        band_pct = st.slider("正2占比偏離目標幾個百分點觸發", 2.5, 25.0, 7.5, 2.5)
        down_pct, up_pct = 30, 40

cash_yield = cash_yield_pct / 100
buy_cost = buy_cost_pct / 100
sell_cost = sell_cost_pct / 100
target = target_pct / 100
down_trigger = down_pct / 100
up_trigger = up_pct / 100
band = band_pct / 100

status = database_status(db)
has_required_data = (
    not db.empty and
    all((db["Symbol"] == s).any() for s in ["TAIEX", "0050", "00631L"])
)

tabs = st.tabs([
    "① 實際資料回測",
    "② 參數最佳化",
    "③ 2000/2008壓力測試",
    "④ 正2偏差模型",
    "⑤ 資料管理",
])

actual = None
coef = calibration = vol_q = None

if has_required_data:
    try:
        actual = prepare_actual_data(db)
        coef, calibration, vol_q = fit_gap_model(
            actual["r0050_price"], actual["r2x"]
        )
    except Exception as e:
        has_required_data = False
        prep_error = str(e)
else:
    prep_error = "本地市場資料庫尚未建立完整。"


# ----------------------------
# ① 實際資料回測
# ----------------------------
with tabs[0]:
    if not has_required_data:
        st.warning(f"{prep_error} 請先到「⑤ 資料管理」建立或匯入資料庫。")
    else:
        p2x = actual["p2x"]
        lev_level = 100 * p2x / float(p2x.iloc[0])

        strat, nreb, turnover = simulate_rebalance(
            actual["r2x"], lev_level, capital, target,
            trigger_mode, down_trigger, up_trigger, band,
            cash_yield, buy_cost, sell_cost, min_days
        )
        bench = simulate_benchmark(
            actual["r0050_total"].reindex(strat.index).fillna(0), capital
        )

        c0, c1, c2, c3 = st.columns(4)
        c0.metric("初始總資產", f"{capital:,.0f} 元")
        c1.metric("初始正2", f"{capital * target:,.0f} 元")
        c2.metric("初始現金", f"{capital * (1-target):,.0f} 元")
        c3.metric("初始約當曝險", f"{target * 2:.0%}")

        summary = pd.DataFrame([
            metrics(bench["Portfolio"], capital),
            metrics(strat["Portfolio"], capital, nreb, turnover),
        ], index=["100% 0050", f"{target_pct}% 正2＋現金"])

        st.dataframe(
            summary.style.format({
                "期末資產": "{:,.0f}",
                "累積報酬": "{:.1%}",
                "CAGR": "{:.2%}",
                "年化波動": "{:.2%}",
                "最大回撤": "{:.2%}",
                "Sharpe(無風險=0)": "{:.2f}",
                "Calmar": "{:.2f}",
                "再平衡次數": "{:.0f}",
                "累計換手金額": "{:,.0f}",
            }),
            use_container_width=True
        )

        curve = pd.concat([
            bench["Portfolio"].rename("100% 0050"),
            strat["Portfolio"].rename("正2＋現金"),
        ], axis=1)
        st.line_chart(curve)

        st.caption("回撤")
        dd = pd.concat([
            bench["Drawdown"].rename("100% 0050"),
            strat["Drawdown"].rename("正2＋現金"),
        ], axis=1)
        st.line_chart(dd)

        with st.expander("前10個交易日資產檢核"):
            audit = strat[["Portfolio", "Leveraged", "Cash", "LevWeight"]].head(10)
            st.dataframe(
                audit.style.format({
                    "Portfolio": "{:,.0f}",
                    "Leveraged": "{:,.0f}",
                    "Cash": "{:,.0f}",
                    "LevWeight": "{:.2%}",
                }),
                use_container_width=True
            )

        with st.expander("資料完整性／ETF分割修復"):
            if actual["split0050"]:
                st.write("0050偵測到的分割跳空")
                st.dataframe(pd.DataFrame(actual["split0050"]), use_container_width=True)
            if actual["split2x"]:
                st.write("00631L偵測到的分割跳空")
                st.dataframe(pd.DataFrame(actual["split2x"]), use_container_width=True)
            if not actual["split0050"] and not actual["split2x"]:
                st.success("未偵測到需要額外修復的分割跳空。")
            for w in actual["warnings"]:
                st.warning(w)

        trades = strat.loc[strat["Reason"].isin(["下跌門檻", "上漲門檻", "低於比例帶", "高於比例帶"])]
        with st.expander(f"再平衡紀錄（{len(trades)}次）"):
            st.dataframe(
                trades[["Portfolio", "Leveraged", "Cash", "LevWeight", "Reason"]],
                use_container_width=True
            )


# ----------------------------
# ② 參數最佳化
# ----------------------------
with tabs[1]:
    if not has_required_data:
        st.warning("請先建立本地市場資料庫。")
    else:
        st.subheader("搜尋『穩健區域』而非單一最高報酬參數")

        c1, c2, c3 = st.columns(3)
        with c1:
            t_min, t_max = st.slider("正2比例搜尋範圍（%）", 10, 80, (30, 70), 5)
        with c2:
            d_min, d_max = st.slider("下跌門檻搜尋範圍（%）", 5, 60, (15, 45), 5)
        with c3:
            u_min, u_max = st.slider("上漲門檻搜尋範圍（%）", 5, 100, (20, 60), 5)

        dd_limit = st.slider("最大回撤限制", -70, -10, -35, 1) / 100

        if st.button("開始參數掃描", type="primary"):
            lev_level = 100 * actual["p2x"] / float(actual["p2x"].iloc[0])
            targets = np.arange(t_min, t_max + 0.1, 5) / 100
            downs = np.arange(d_min, d_max + 0.1, 5) / 100
            ups = np.arange(u_min, u_max + 0.1, 5) / 100
            total_n = len(targets) * len(downs) * len(ups)

            rows = []
            prog = st.progress(0)
            k = 0

            for t in targets:
                for d in downs:
                    for u in ups:
                        sim, nr, to = simulate_rebalance(
                            actual["r2x"], lev_level, capital, t,
                            "價格漲跌門檻", d, u, 0.075,
                            cash_yield, buy_cost, sell_cost, min_days
                        )
                        m = metrics(sim["Portfolio"], capital, nr, to)
                        rows.append({
                            "正2比例": t,
                            "下跌門檻": d,
                            "上漲門檻": u,
                            **m,
                        })
                        k += 1
                        if k % max(1, total_n // 100) == 0:
                            prog.progress(k / total_n)

            prog.empty()
            res = pd.DataFrame(rows).sort_values(
                ["Calmar", "CAGR"], ascending=False
            )
            st.session_state["grid_result_v12"] = res

        if "grid_result_v12" in st.session_state:
            res = st.session_state["grid_result_v12"]
            feasible = res[res["最大回撤"] >= dd_limit].copy()
            front = pareto_frontier(res)

            st.markdown("**Calmar 前20名**")
            st.dataframe(
                res.head(20).style.format({
                    "正2比例": "{:.0%}",
                    "下跌門檻": "{:.0%}",
                    "上漲門檻": "{:.0%}",
                    "期末資產": "{:,.0f}",
                    "累積報酬": "{:.1%}",
                    "CAGR": "{:.2%}",
                    "年化波動": "{:.2%}",
                    "最大回撤": "{:.2%}",
                    "Sharpe(無風險=0)": "{:.2f}",
                    "Calmar": "{:.2f}",
                }),
                use_container_width=True
            )

            st.markdown(f"**符合最大回撤 ≥ {dd_limit:.0%} 的最高CAGR**")
            if feasible.empty:
                st.warning("目前搜尋範圍沒有符合回撤限制的組合。")
            else:
                st.dataframe(
                    feasible.sort_values("CAGR", ascending=False).head(20)
                    .style.format({
                        "正2比例": "{:.0%}",
                        "下跌門檻": "{:.0%}",
                        "上漲門檻": "{:.0%}",
                        "期末資產": "{:,.0f}",
                        "CAGR": "{:.2%}",
                        "最大回撤": "{:.2%}",
                        "Calmar": "{:.2f}",
                    }),
                    use_container_width=True
                )

            st.markdown("**報酬—回撤 Pareto frontier**")
            st.dataframe(
                front.sort_values("最大回撤", ascending=False)
                .style.format({
                    "正2比例": "{:.0%}",
                    "下跌門檻": "{:.0%}",
                    "上漲門檻": "{:.0%}",
                    "期末資產": "{:,.0f}",
                    "CAGR": "{:.2%}",
                    "最大回撤": "{:.2%}",
                    "Calmar": "{:.2f}",
                }),
                use_container_width=True
            )


# ----------------------------
# ③ 壓力測試
# ----------------------------
with tabs[2]:
    if not has_required_data:
        st.warning("請先建立本地市場資料庫。")
    else:
        st.subheader("2000 / 2008 股災壓力測試")
        mode = st.selectbox(
            "歷史期間",
            ["2000網路泡沫（TAIEX代理）", "2008金融海嘯（0050代理）"]
        )
        model_mode = st.radio(
            "正2生成方式",
            ["校準平均", "波動狀態抽樣", "理論2倍"],
            horizontal=True
        )

        if mode.startswith("2000"):
            base = actual["twii"]["Close"].loc["2000-01-01":"2002-12-31"].dropna()
            ur = base.pct_change().dropna()
            br = ur.copy()
            st.info("2000年沒有0050與00631L，這裡使用TAIEX作市場代理，因此是壓力測試，不是真實00631L回測。")
        else:
            close0050, _ = repair_split_jumps(
                actual["d0050"]["Close"].loc["2007-01-01":"2009-12-31"].dropna()
            )
            ur = close0050.pct_change().dropna()
            adj = actual["d0050"]["AdjClose"].reindex(ur.index).dropna()
            br = adj.pct_change().reindex(ur.index).fillna(0)
            st.info("2008年00631L尚未上市；使用0050當時實際行情生成模擬正2。")

        if model_mode == "波動狀態抽樣":
            n_mc = st.slider("蒙地卡羅路徑數", 50, 1000, 300, 50)
            seed = st.number_input("亂數種子", value=42, step=1)
            rows = []

            for j in range(n_mc):
                sr = synthetic_2x_returns(
                    ur, coef, calibration, vol_q,
                    "波動狀態抽樣", int(seed + j)
                )
                lvl = make_level(sr)
                sim, nr, to = simulate_rebalance(
                    sr, lvl, capital, target,
                    trigger_mode, down_trigger, up_trigger, band,
                    cash_yield, buy_cost, sell_cost, min_days
                )
                rows.append(metrics(sim["Portfolio"], capital, nr, to))

            dist = pd.DataFrame(rows)
            st.dataframe(
                dist[["期末資產", "CAGR", "最大回撤"]]
                .quantile([0.05, 0.50, 0.95])
                .style.format({
                    "期末資產": "{:,.0f}",
                    "CAGR": "{:.2%}",
                    "最大回撤": "{:.2%}",
                }),
                use_container_width=True
            )
        else:
            sr = synthetic_2x_returns(
                ur, coef, calibration, vol_q, model_mode, 42
            )
            lvl = make_level(sr)
            sim, nr, to = simulate_rebalance(
                sr, lvl, capital, target,
                trigger_mode, down_trigger, up_trigger, band,
                cash_yield, buy_cost, sell_cost, min_days
            )
            bm = simulate_benchmark(br.reindex(sim.index).fillna(0), capital)

            result = pd.DataFrame([
                metrics(bm["Portfolio"], capital),
                metrics(sim["Portfolio"], capital, nr, to),
            ], index=["100%市場／0050代理", "模擬正2＋現金"])

            st.dataframe(
                result.style.format({
                    "期末資產": "{:,.0f}",
                    "累積報酬": "{:.1%}",
                    "CAGR": "{:.2%}",
                    "年化波動": "{:.2%}",
                    "最大回撤": "{:.2%}",
                    "Sharpe(無風險=0)": "{:.2f}",
                    "Calmar": "{:.2f}",
                }),
                use_container_width=True
            )

            st.line_chart(pd.concat([
                bm["Portfolio"].rename("100%市場／0050代理"),
                sim["Portfolio"].rename("模擬正2＋現金"),
            ], axis=1))


# ----------------------------
# ④ 偏差模型
# ----------------------------
with tabs[3]:
    if not has_required_data:
        st.warning("請先建立本地市場資料庫。")
    else:
        st.subheader("00631L實際報酬 vs 理論每日2倍")
        gap = calibration["gap"]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("平均日偏差", f"{gap.mean():.4%}")
        c2.metric("偏差標準差", f"{gap.std():.4%}")
        c3.metric("5%分位", f"{gap.quantile(.05):.4%}")
        c4.metric("95%分位", f"{gap.quantile(.95):.4%}")

        coef_table = pd.DataFrame({
            "變數": ["常數", "日報酬", "|日報酬|", "日報酬²", "20日年化波動", "20日動能"],
            "係數": coef,
        })
        st.dataframe(coef_table, use_container_width=True)

        st.line_chart(pd.DataFrame({
            "實際偏差": calibration["gap"],
            "模型預測偏差": calibration["pred_gap"],
        }))


# ----------------------------
# ⑤ 資料管理
# ----------------------------
with tabs[4]:
    st.subheader("本地市場資料庫")
    st.write(
        "一般回測只讀取 GitHub 專案內的 `data/market_daily.csv`。"
        "只有你在這個頁面按下更新按鈕時，才會連 Yahoo Finance。"
    )

    st.dataframe(status, use_container_width=True)

    if db.empty:
        st.warning("目前資料庫是空的。第一次部署後，請按「建立完整歷史資料庫」。")
    else:
        latest = db["Date"].max()
        st.success(f"本地資料庫目前最新日期：{latest.date().isoformat()}")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("建立完整歷史資料庫", type="primary"):
            try:
                with st.spinner("第一次建立資料庫：下載1999年至今資料…"):
                    db = update_database(db, full_rebuild=True)
                st.success("完整歷史資料庫已建立。")
                st.rerun()
            except Exception as e:
                st.error(f"建立資料庫失敗：{e}")

    with c2:
        if st.button("只更新最新缺漏資料"):
            try:
                with st.spinner("只下載各資料最後日期附近至今天的資料…"):
                    db = update_database(db, full_rebuild=False)
                st.success("資料庫已更新。")
                st.rerun()
            except Exception as e:
                st.error(f"更新失敗：{e}")

    st.divider()
    st.markdown("#### 匯入／匯出資料庫")

    uploaded = st.file_uploader(
        "上傳 market_daily.csv 以覆蓋目前本機資料庫",
        type=["csv"]
    )
    if uploaded is not None:
        try:
            incoming = pd.read_csv(uploaded)
            missing = set(DB_COLUMNS) - set(incoming.columns)
            if missing:
                st.error(f"上傳檔缺少欄位：{', '.join(sorted(missing))}")
            elif st.button("確認匯入這份資料庫"):
                incoming["Date"] = pd.to_datetime(incoming["Date"])
                save_local_database(incoming)
                st.success("已匯入。")
                st.rerun()
        except Exception as e:
            st.error(f"匯入失敗：{e}")

    if not db.empty:
        export = db.copy()
        export["Date"] = pd.to_datetime(export["Date"]).dt.strftime("%Y-%m-%d")
        csv_bytes = export.to_csv(index=False).encode("utf-8-sig")

        st.download_button(
            "下載最新版 market_daily.csv",
            data=csv_bytes,
            file_name="market_daily.csv",
            mime="text/csv",
        )

        st.info(
            "Streamlit Community Cloud 的執行環境不是永久儲存空間。"
            "如果你在網站上更新了資料，請下載最新版 market_daily.csv，"
            "再把它覆蓋回 GitHub 的 data/market_daily.csv。"
            "這樣下一次重新部署或休眠喚醒後，仍會使用最新資料。"
        )

    st.divider()
    st.markdown("#### 資料安全檢查")
    if not db.empty:
        dup = db.duplicated(["Date", "Symbol"]).sum()
        invalid = db["Close"].isna().sum()
        c1, c2, c3 = st.columns(3)
        c1.metric("總資料列", f"{len(db):,}")
        c2.metric("重複日期/代號", f"{dup:,}")
        c3.metric("Close缺值", f"{invalid:,}")


st.divider()
st.caption(
    "v1.2：歷史資料以本地資料庫為主；2000年使用TAIEX代理。"
    "歷史回測與模型最佳化均不代表未來報酬。"
)
