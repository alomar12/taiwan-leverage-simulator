
from pathlib import Path
from datetime import date
import io
import time
import requests

import numpy as np
import pandas as pd
import streamlit as st
import yfinance as yf

st.set_page_config(page_title="台灣50正2再平衡模擬器 v1.5", layout="wide")

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



@st.cache_data(ttl=86400, show_spinner=False)
def download_twse_taiex_month(year: int, month: int):
    """證交所官方 TAIEX 每月 OHLC。"""
    yyyymm01 = f"{year:04d}{month:02d}01"
    url = "https://www.twse.com.tw/rwd/zh/TAIEX/MI_5MINS_HIST"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json,text/plain,*/*",
        "Referer": "https://www.twse.com.tw/zh/indices/taiex/mi-5min-hist.html",
    }

    last_error = None
    for attempt in range(5):
        try:
            resp = requests.get(
                url,
                params={"date": yyyymm01, "response": "json"},
                headers=headers,
                timeout=30,
            )
            if resp.status_code == 429:
                time.sleep(2 + attempt * 2)
                continue
            resp.raise_for_status()
            obj = resp.json()

            if obj.get("stat") != "OK":
                return empty_database()

            rows = obj.get("data", [])
            if not rows:
                return empty_database()

            out = []
            for row in rows:
                roc = str(row[0]).strip().split("/")
                if len(roc) != 3:
                    continue
                gy = int(roc[0]) + 1911
                dt = pd.Timestamp(gy, int(roc[1]), int(roc[2]))
                close = float(str(row[4]).replace(",", "").strip())
                out.append({
                    "Date": dt,
                    "Symbol": "TAIEX",
                    "Close": close,
                    "AdjClose": close,
                })
            return pd.DataFrame(out, columns=DB_COLUMNS)

        except Exception as e:
            last_error = e
            time.sleep(1.5 + attempt * 1.5)

    raise RuntimeError(
        f"證交所官方TAIEX {year}/{month:02d} 下載失敗：{last_error}"
    )


def month_starts(start, end):
    p = pd.Timestamp(start).replace(day=1)
    e = pd.Timestamp(end).replace(day=1)
    while p <= e:
        yield p
        p = p + pd.offsets.MonthBegin(1)


def download_twse_taiex_range(start: str, end: str, progress_label=True):
    """
    逐月下載證交所官方 TAIEX。
    只用來建立/補齊固定歷史資料庫，不在一般回測時執行。
    """
    months = list(month_starts(start, end))
    pieces = []
    prog = st.progress(0) if progress_label and months else None

    for i, m in enumerate(months, start=1):
        piece = download_twse_taiex_month(m.year, m.month)
        if piece is not None and not piece.empty:
            pieces.append(piece)
        if prog is not None:
            prog.progress(i / len(months))

        # 證交所逐月端點避免過度頻繁請求
        if i < len(months):
            time.sleep(0.45)

    if prog is not None:
        prog.empty()

    if not pieces:
        return empty_database()
    return pd.concat(pieces, ignore_index=True)


def ensure_official_taiex_history(db: pd.DataFrame, start="1999-01-01", end=None):
    """
    將 TAIEX 1999 起缺少的月份，以證交所官方資料補入。
    已存在的月份不再下載。
    """
    end = pd.Timestamp(end or date.today().isoformat())
    start = pd.Timestamp(start)

    existing = db.loc[db["Symbol"] == "TAIEX", ["Date"]].copy()
    existing_months = set()
    if not existing.empty:
        existing_months = set(
            pd.to_datetime(existing["Date"]).dt.to_period("M").astype(str)
        )

    needed = [
        m for m in month_starts(start, end)
        if str(m.to_period("M")) not in existing_months
    ]

    if not needed:
        return db, 0

    pieces = []
    prog = st.progress(0)
    status_box = st.empty()

    for i, m in enumerate(needed, start=1):
        status_box.caption(
            f"補齊證交所官方 TAIEX：{m.year}/{m.month:02d} "
            f"（{i}/{len(needed)}）"
        )
        piece = download_twse_taiex_month(m.year, m.month)
        if piece is not None and not piece.empty:
            pieces.append(piece)

        prog.progress(i / len(needed))
        if i < len(needed):
            time.sleep(0.45)

    prog.empty()
    status_box.empty()

    if pieces:
        official = pd.concat(pieces, ignore_index=True)
        db = merge_database(db, official)
        return db, len(official)

    return db, 0


def update_database(db: pd.DataFrame, full_rebuild=False, end=None):
    """
    混合資料源：
    - TAIEX：Yahoo 只作快速補近期；1999起完整歷史由 TWSE 官方逐月補齊。
    - 0050 / 00631L：Yahoo Finance，主要用途是保留 Adj Close。
    """
    end = end or date.today().isoformat()

    if full_rebuild:
        base = empty_database()
    else:
        base = db.copy()

    all_new = []
    keys = ["0050", "00631L", "TAIEX"]
    prog = st.progress(0)

    for i, key in enumerate(keys, start=1):
        info = SYMBOLS[key]
        if full_rebuild or base.empty:
            # TAIEX 的早期歷史由 TWSE 官方負責；Yahoo只抓較近期資料，
            # 可顯著降低 Yahoo 長區間失敗造成的影響。
            if key == "TAIEX":
                start = "2009-01-01"
            else:
                start = info["start"]
        else:
            x = base.loc[base["Symbol"] == key]
            if x.empty:
                start = "2009-01-01" if key == "TAIEX" else info["start"]
            else:
                start = (x["Date"].max() - pd.Timedelta(days=14)).strftime("%Y-%m-%d")

        try:
            new = download_yahoo_long(key, start, end)
            all_new.append(new)
        except Exception:
            # Yahoo掛掉不阻止TAIEX官方歷史補齊；
            # 0050/00631L若已有本地資料也可繼續使用。
            pass

        prog.progress(i / len(keys))

    prog.empty()

    if all_new:
        fresh = pd.concat(all_new, ignore_index=True)
        base = merge_database(base, fresh)

    # 最重要：TAIEX 1999年至今缺漏月份，以證交所官方資料補齊。
    base, _ = ensure_official_taiex_history(
        base, start="1999-01-01", end=end
    )

    save_local_database(base)
    return base


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



def build_long_history_series(actual, coef, calibration, vol_q):
    """
    1999年至今長期研究序列。
    市場基準使用官方TAIEX；模擬正2為每日 2x TAIEX 報酬
    加上2014年至今實際00631L校準所得的條件偏差。

    1999-2013不是00631L真實歷史，只能視為歷史重建／壓力測試。
    """
    twii = actual["twii"]["Close"].dropna().sort_index()
    if twii.empty:
        raise RuntimeError("TAIEX歷史資料為空。")

    market_ret = twii.pct_change().fillna(0)
    sim_2x_ret = synthetic_2x_returns(
        market_ret, coef, calibration, vol_q, "校準平均", 42
    )
    market_level = 100.0 * twii / float(twii.iloc[0])
    sim_2x_level = make_level(sim_2x_ret, start=100.0)

    chart = pd.DataFrame({
        "TAIEX大盤": market_level,
        "模擬正2": sim_2x_level.reindex(market_level.index),
    })

    if not actual["p2x"].empty:
        p2x = actual["p2x"].dropna()
        actual_2x = 100.0 * p2x / float(p2x.iloc[0])
        chart["實際00631L（上市後）"] = actual_2x.reindex(chart.index)

    return {
        "market_level": twii,
        "market_ret": market_ret,
        "sim_2x_ret": sim_2x_ret,
        "sim_2x_level": sim_2x_level,
        "chart": chart,
    }


def slice_returns_and_level(ret, level, start_date):
    """從使用者指定日期起，取第一個可用交易日作投資起點。"""
    start = pd.Timestamp(start_date)
    idx = ret.index.intersection(level.index)
    idx = idx[idx >= start]
    if len(idx) < 2:
        raise RuntimeError("指定投資起日之後沒有足夠交易資料。")
    r = ret.reindex(idx).copy()
    lvl = level.reindex(idx).ffill().copy()
    # 第一個交易日固定視為投入日，不先承受當日報酬
    r.iloc[0] = 0.0
    return r, lvl, idx[0]

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



def buy_leveraged_with_cash(lev, cash, amount, buy_cost):
    """只用現金買正2；絕不允許負現金或融資。"""
    max_buy = cash / (1 + buy_cost) if buy_cost >= 0 else cash
    x = max(0.0, min(float(amount), max_buy))
    cost = x * buy_cost
    lev += x
    cash -= x + cost
    if cash < 1e-8:
        cash = 0.0
    return lev, cash, x, cost


def sell_leveraged_to_cash(lev, cash, amount, sell_cost):
    """賣出正2轉回現金。"""
    x = max(0.0, min(float(amount), lev))
    cost = x * sell_cost
    lev -= x
    cash += x - cost
    return lev, cash, x, cost


def rebalance_to_initial_target(lev, cash, target, buy_cost, sell_cost):
    """回復原始正2／現金配置。"""
    return rebalance_to_target(lev, cash, target, buy_cost, sell_cost)


def simulate_ladder_buy_strategy(
    lev_ret,
    market_level,
    capital,
    initial_lev_target=0.50,
    start_drawdown=0.20,
    drawdown_step=0.05,
    buy_step=0.05,
    buy_basis="當下總資產",
    exit_style="回到前高一次再平衡",
    rebound_start=0.10,
    rebound_step=0.10,
    sell_step=0.05,
    cash_yield=0.0,
    buy_cost=0.0,
    sell_cost=0.0,
):
    """
    越跌越買策略：
    - 起始：initial_lev_target 正2，其餘現金。
    - 市場自前次歷史高點回撤 start_drawdown，買第一階。
    - 之後每再跌 drawdown_step，再買一階。
    - 每階買 buy_step × 當下總資產（或初始本金）。
    - 可以把現金買到 0，但禁止融資。
    - 反彈端三種：
      1) 回到前高一次再平衡。
      2) 低點反彈指定幅度一次再平衡。
      3) 低點反彈分批再平衡，前高時強制完成。
    """
    idx = lev_ret.index.intersection(market_level.index)
    if len(idx) < 2:
        raise ValueError("可用交易日不足。")

    r = lev_ret.reindex(idx).fillna(0)
    market = market_level.reindex(idx).ffill()

    lev = capital * initial_lev_target
    cash = capital * (1 - initial_lev_target)
    daily_cash = (1 + cash_yield) ** (1 / 252) - 1

    # 以尚未進入本輪下跌前的歷史高點為參考高點
    reference_peak = float(market.iloc[0])
    cycle_active = False
    cycle_peak = reference_peak
    cycle_low = reference_peak

    next_buy_trigger = start_drawdown
    next_sell_trigger = rebound_start
    buy_stage = 0
    sell_stage = 0
    n_buys = 0
    n_sells = 0
    n_resets = 0
    turnover = 0.0

    rows = []
    events = []
    cash_exhausted_dates = []

    first_dt = idx[0]
    rows.append((
        first_dt, capital, lev, cash,
        lev / capital if capital else 0.0,
        0.0, reference_peak, reference_peak,
        0.0, "", buy_stage, sell_stage
    ))

    for dt in idx[1:]:
        lev *= 1 + float(r.loc[dt])
        cash *= 1 + daily_cash
        mkt = float(market.loc[dt])

        if (not cycle_active) and mkt > reference_peak:
            reference_peak = mkt

        drawdown = mkt / reference_peak - 1.0 if reference_peak > 0 else 0.0
        reason_parts = []

        # 啟動一輪下跌加碼
        if (not cycle_active) and drawdown <= -start_drawdown:
            cycle_active = True
            cycle_peak = reference_peak
            cycle_low = mkt
            next_buy_trigger = start_drawdown
            next_sell_trigger = rebound_start
            buy_stage = 0
            sell_stage = 0
            reason_parts.append("啟動加碼循環")
            events.append({
                "Date": dt,
                "Event": "啟動加碼循環",
                "Market": mkt,
                "Drawdown": drawdown,
                "ReboundFromLow": 0.0,
                "Amount": 0.0,
                "CashAfter": cash,
                "LevAfter": lev,
            })

        if cycle_active:
            # 新低時更新本輪低點。分批反彈門檻重新從最新低點計算，
            # 但已經執行過的賣出階數不回復，避免反覆買賣同一階。
            if mkt < cycle_low:
                cycle_low = mkt

            cycle_dd_abs = max(0.0, 1.0 - mkt / cycle_peak)

            # ----------------------------------------------------
            # 下跌：跨過幾個門檻，就補幾階，直到現金用完
            # ----------------------------------------------------
            while (
                cash > 1e-8
                and next_buy_trigger < 1.0
                and cycle_dd_abs + 1e-12 >= next_buy_trigger
            ):
                total_before = lev + cash
                if buy_basis == "初始本金":
                    desired_buy = capital * buy_step
                else:
                    desired_buy = total_before * buy_step

                lev, cash, bought, cost = buy_leveraged_with_cash(
                    lev, cash, desired_buy, buy_cost
                )
                turnover += bought
                n_buys += 1
                buy_stage += 1

                reason_parts.append(f"第{buy_stage}階加碼")
                events.append({
                    "Date": dt,
                    "Event": f"第{buy_stage}階加碼",
                    "Market": mkt,
                    "Drawdown": -cycle_dd_abs,
                    "ReboundFromLow": mkt / cycle_low - 1 if cycle_low > 0 else 0.0,
                    "Trigger": next_buy_trigger,
                    "Amount": bought,
                    "Cost": cost,
                    "CashAfter": cash,
                    "LevAfter": lev,
                })

                next_buy_trigger += drawdown_step

                if cash <= 1e-8:
                    cash = 0.0
                    cash_exhausted_dates.append(dt)
                    reason_parts.append("現金用完")
                    events.append({
                        "Date": dt,
                        "Event": "現金用完",
                        "Market": mkt,
                        "Drawdown": -cycle_dd_abs,
                        "ReboundFromLow": mkt / cycle_low - 1 if cycle_low > 0 else 0.0,
                        "Amount": 0.0,
                        "CashAfter": cash,
                        "LevAfter": lev,
                    })
                    break

            rebound = mkt / cycle_low - 1.0 if cycle_low > 0 else 0.0

            # ----------------------------------------------------
            # 反彈：一次再平衡
            # ----------------------------------------------------
            do_full_reset = False
            reset_reason = ""

            if exit_style == "回到前高一次再平衡":
                if mkt >= cycle_peak:
                    do_full_reset = True
                    reset_reason = "回到前高"

            elif exit_style == "低點反彈指定幅度一次再平衡":
                if rebound >= rebound_start:
                    do_full_reset = True
                    reset_reason = f"低點反彈{rebound_start:.0%}"

            # ----------------------------------------------------
            # 反彈：分批賣正2
            # ----------------------------------------------------
            elif exit_style == "低點反彈分批再平衡":
                # 每跨過一個反彈門檻賣一階，但不賣到低於原始配置。
                while rebound + 1e-12 >= next_sell_trigger:
                    total_before = lev + cash
                    min_lev = total_before * initial_lev_target
                    excess_lev = max(0.0, lev - min_lev)

                    if excess_lev <= 1e-8:
                        break

                    desired_sell = total_before * sell_step
                    sell_amount = min(desired_sell, excess_lev)

                    lev, cash, sold, cost = sell_leveraged_to_cash(
                        lev, cash, sell_amount, sell_cost
                    )
                    turnover += sold
                    n_sells += 1
                    sell_stage += 1

                    reason_parts.append(f"第{sell_stage}階反彈減碼")
                    events.append({
                        "Date": dt,
                        "Event": f"第{sell_stage}階反彈減碼",
                        "Market": mkt,
                        "Drawdown": mkt / cycle_peak - 1 if cycle_peak > 0 else 0.0,
                        "ReboundFromLow": rebound,
                        "Trigger": next_sell_trigger,
                        "Amount": sold,
                        "Cost": cost,
                        "CashAfter": cash,
                        "LevAfter": lev,
                    })

                    next_sell_trigger += rebound_step

                # 回到前高，不論前面分批完成多少，都強制回原始配置
                if mkt >= cycle_peak:
                    do_full_reset = True
                    reset_reason = "回到前高強制完成"

            if do_full_reset:
                lev, cash, traded = rebalance_to_initial_target(
                    lev, cash, initial_lev_target, buy_cost, sell_cost
                )
                turnover += traded
                n_resets += 1
                reason_parts.append(f"{reset_reason}→恢復原始配置")

                events.append({
                    "Date": dt,
                    "Event": f"{reset_reason}→恢復原始配置",
                    "Market": mkt,
                    "Drawdown": mkt / cycle_peak - 1 if cycle_peak > 0 else 0.0,
                    "ReboundFromLow": rebound,
                    "Amount": traded,
                    "CashAfter": cash,
                    "LevAfter": lev,
                })

                # 完成本輪，重新建立下一輪歷史高點基準
                cycle_active = False
                reference_peak = max(reference_peak, mkt)
                cycle_peak = reference_peak
                cycle_low = reference_peak
                next_buy_trigger = start_drawdown
                next_sell_trigger = rebound_start
                buy_stage = 0
                sell_stage = 0

        total = lev + cash
        lev_weight = lev / total if total > 0 else 0.0
        current_rebound = (
            mkt / cycle_low - 1.0
            if cycle_active and cycle_low > 0
            else 0.0
        )

        rows.append((
            dt, total, lev, cash, lev_weight,
            drawdown, reference_peak, cycle_low,
            current_rebound, "；".join(reason_parts),
            buy_stage, sell_stage
        ))

    df = pd.DataFrame(
        rows,
        columns=[
            "Date", "Portfolio", "Leveraged", "Cash", "LevWeight",
            "MarketDrawdown", "ReferencePeak", "CycleLow",
            "ReboundFromLow", "Reason", "BuyStage", "SellStage"
        ],
    ).set_index("Date")

    df["Peak"] = df["Portfolio"].cummax()
    df["Drawdown"] = df["Portfolio"] / df["Peak"] - 1

    event_df = pd.DataFrame(events)
    if not event_df.empty:
        event_df = event_df.set_index("Date").sort_index()

    stats = {
        "加碼次數": n_buys,
        "反彈分批減碼次數": n_sells,
        "完整重置次數": n_resets,
        "現金用完次數": len(cash_exhausted_dates),
        "首次現金用完日期": cash_exhausted_dates[0] if cash_exhausted_dates else None,
        "累計換手金額": turnover,
    }
    return df, event_df, stats


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
st.title("台灣50正2 × 現金：再平衡與股災壓力測試模擬器 v1.5")
st.caption(
    "v1.5 改為『本地歷史資料庫優先』：一般回測不連Yahoo Finance；"
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
    "⑤ 越跌越買策略",
    "⑥ 資料管理",
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
        st.warning(f"{prep_error} 請先到「⑥ 資料管理」建立或匯入資料庫。")
    else:
        st.subheader("回測資料模式與投資起點")
        backtest_mode = st.radio(
            "資料模式",
            [
                "2014/10/31至今｜實際0050＋實際00631L",
                "1999年至今｜官方TAIEX＋校準模擬正2",
            ],
            horizontal=True,
            key="main_backtest_mode",
        )

        if backtest_mode.startswith("2014"):
            min_start = max(pd.Timestamp("2014-10-31"), actual["r2x"].index.min())
            max_start = actual["r2x"].index.max() - pd.Timedelta(days=1)
            default_start = min_start
        else:
            long_data = build_long_history_series(actual, coef, calibration, vol_q)
            min_start = long_data["market_ret"].index.min()
            max_start = long_data["market_ret"].index.max() - pd.Timedelta(days=1)
            default_start = min_start

        invest_start = st.date_input(
            "投資開始日期",
            value=default_start.date(),
            min_value=min_start.date(),
            max_value=max_start.date(),
            key="main_invest_start",
        )

        if backtest_mode.startswith("2014"):
            raw_ret = actual["r2x"]
            raw_level = 100 * actual["p2x"] / float(actual["p2x"].iloc[0])
            lev_ret_bt, lev_level_bt, actual_start = slice_returns_and_level(
                raw_ret, raw_level, invest_start
            )
            bench_ret = actual["r0050_total"].reindex(lev_ret_bt.index).fillna(0).copy()
            bench_ret.iloc[0] = 0.0
            bench_name = "100% 0050"
            lev_name = f"{target_pct}% 實際00631L＋現金"
            data_note = (
                "本模式完全使用00631L上市後的實際日報酬；"
                "0050使用還原價格計算總報酬。"
            )
        else:
            raw_ret = long_data["sim_2x_ret"]
            raw_level = long_data["sim_2x_level"]
            lev_ret_bt, lev_level_bt, actual_start = slice_returns_and_level(
                raw_ret, raw_level, invest_start
            )
            bench_ret = long_data["market_ret"].reindex(lev_ret_bt.index).fillna(0).copy()
            bench_ret.iloc[0] = 0.0
            bench_name = "100% TAIEX大盤"
            lev_name = f"{target_pct}% 模擬正2＋現金"
            data_note = (
                "1999–2013的正2為模型重建值，不是00631L實際歷史。"
                "模型以TAIEX逐日報酬×2，再加上2014年至今正2實際追蹤偏差的校準結果。"
            )

        st.info(f"實際投入起始交易日：{actual_start.date()}。{data_note}")

        strat, nreb, turnover = simulate_rebalance(
            lev_ret_bt, lev_level_bt, capital, target,
            trigger_mode, down_trigger, up_trigger, band,
            cash_yield, buy_cost, sell_cost, min_days
        )
        bench = simulate_benchmark(bench_ret.reindex(strat.index).fillna(0), capital)

        c0, c1, c2, c3, c4 = st.columns(5)
        c0.metric("初始總資產", f"{capital:,.0f} 元")
        c1.metric("初始正2", f"{capital * target:,.0f} 元")
        c2.metric("初始現金", f"{capital * (1-target):,.0f} 元")
        c3.metric("初始約當曝險", f"{target * 2:.0%}")
        c4.metric("投資起日", actual_start.strftime("%Y/%m/%d"))

        summary = pd.DataFrame([
            metrics(bench["Portfolio"], capital),
            metrics(strat["Portfolio"], capital, nreb, turnover),
        ], index=[bench_name, lev_name])

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
            bench["Portfolio"].rename(bench_name),
            strat["Portfolio"].rename(lev_name),
        ], axis=1)
        st.line_chart(curve)

        st.caption("回撤")
        dd = pd.concat([
            bench["Drawdown"].rename(bench_name),
            strat["Drawdown"].rename(lev_name),
        ], axis=1)
        st.line_chart(dd)

        if not backtest_mode.startswith("2014"):
            st.divider()
            st.subheader("1999年至今｜大盤、模擬正2與實際00631L走勢")
            st.caption(
                "三條線都正規化為各自起始值100。實際00631L只會從2014/10/31之後出現；"
                "1999–2013的模擬正2用來研究若當時存在每日2倍產品可能出現的路徑。"
            )
            long_chart = long_data["chart"].loc[pd.Timestamp(invest_start):].copy()
            # 讓比較更直觀：從選定起日重新正規化TAIEX與模擬正2為100
            for col in ["TAIEX大盤", "模擬正2"]:
                valid = long_chart[col].dropna()
                if not valid.empty:
                    long_chart[col] = 100 * long_chart[col] / float(valid.iloc[0])
            st.line_chart(long_chart)

            # 顯示2014後模型與實際00631L的校驗差距
            compare = long_data["chart"][["模擬正2", "實際00631L（上市後）"]].dropna()
            if not compare.empty:
                compare = compare / compare.iloc[0] * 100
                st.caption("2014/10/31後：模擬正2 vs 實際00631L（同日起點=100）")
                st.line_chart(compare)

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
            ["2000網路泡沫（官方TAIEX）", "2008金融海嘯（官方TAIEX）"]
        )
        model_mode = st.radio(
            "正2生成方式",
            ["校準平均", "波動狀態抽樣", "理論2倍"],
            horizontal=True
        )

        if mode.startswith("2000"):
            base = actual["twii"]["Close"].loc["2000-01-01":"2002-12-31"].dropna()
            if base.empty or base.index.min() > pd.Timestamp("2000-01-31"):
                st.error("本地資料庫缺少2000年TAIEX。請到「⑥ 資料管理」按『補齊1999年至今官方TAIEX』。")
                st.stop()
            ur = base.pct_change().dropna()
            br = ur.copy()
            st.info("2000年沒有0050與00631L；使用證交所官方TAIEX重建歷史正2，因此屬壓力測試。")
        else:
            base = actual["twii"]["Close"].loc["2007-01-01":"2009-12-31"].dropna()
            if base.empty or base.index.min() > pd.Timestamp("2007-01-31"):
                st.error("本地資料庫缺少2007–2009年TAIEX。請到「⑥ 資料管理」按『補齊1999年至今官方TAIEX』。")
                st.stop()
            ur = base.pct_change().dropna()
            br = ur.copy()
            st.info("2008年00631L尚未上市；改用證交所官方TAIEX重建模擬正2，避免0050早期資料來源不完整。")

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
# ⑤ 越跌越買策略
# ----------------------------
with tabs[4]:
    st.subheader("越跌越買到現金用完｜反彈後再平衡")

    if not has_required_data:
        st.warning("請先建立本地市場資料庫。")
    else:
        st.write(
            "這個策略允許現金一路投入到 0，但不允許融資。"
            "下跌門檻以市場指數／0050相對『本輪下跌前歷史高點』計算；"
            "加碼資產為正2。"
        )

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            ladder_initial_pct = st.slider(
                "初始正2比例（%）", 10, 80, 50, 5, key="ladder_initial"
            )
        with c2:
            ladder_start_dd_pct = st.slider(
                "回撤幾%開始第一階加碼", 5, 50, 20, 5, key="ladder_start_dd"
            )
        with c3:
            ladder_dd_step_pct = st.slider(
                "之後每再跌幾%加碼一階", 2, 20, 5, 1, key="ladder_dd_step"
            )
        with c4:
            ladder_buy_pct = st.slider(
                "每階投入比例（%）", 1, 20, 5, 1, key="ladder_buy_pct"
            )

        c5, c6 = st.columns(2)
        with c5:
            ladder_buy_basis = st.radio(
                "每階投入比例的計算基礎",
                ["當下總資產", "初始本金"],
                horizontal=True,
                key="ladder_buy_basis"
            )
        with c6:
            ladder_exit_style = st.selectbox(
                "反彈後如何再平衡",
                [
                    "回到前高一次再平衡",
                    "低點反彈指定幅度一次再平衡",
                    "低點反彈分批再平衡",
                ],
                key="ladder_exit_style"
            )

        rebound_start_pct = 20
        rebound_step_pct = 10
        sell_step_pct = 5

        if ladder_exit_style == "低點反彈指定幅度一次再平衡":
            rebound_start_pct = st.slider(
                "自本輪最低點反彈幾%後，一次恢復原始配置",
                5, 100, 20, 5, key="ladder_rebound_once"
            )

        elif ladder_exit_style == "低點反彈分批再平衡":
            c7, c8, c9 = st.columns(3)
            with c7:
                rebound_start_pct = st.slider(
                    "低點反彈幾%開始第一階減碼",
                    5, 80, 10, 5, key="ladder_rebound_start"
                )
            with c8:
                rebound_step_pct = st.slider(
                    "之後每再漲幾%減碼一階",
                    5, 50, 10, 5, key="ladder_rebound_step"
                )
            with c9:
                sell_step_pct = st.slider(
                    "每階轉回現金比例（總資產%）",
                    1, 20, 5, 1, key="ladder_sell_step"
                )

            st.caption(
                "分批減碼不會把正2降到低於原始正2比例；"
                "若市場先回到本輪前高，會把尚未完成的部分一次恢復原始配置。"
            )

        test_period = st.radio(
            "測試期間／資料",
            [
                "2014/10/31至今｜實際00631L＋0050",
                "2000–2002｜官方TAIEX＋校準模擬正2",
                "2007–2009｜官方TAIEX＋校準模擬正2",
            ],
            key="ladder_period"
        )

        if test_period.startswith("2014"):
            lev_ret_ladder = actual["r2x"]
            market_level_ladder = actual["p0050"].reindex(
                lev_ret_ladder.index
            ).ffill()
            benchmark_ret_ladder = actual["r0050_total"].reindex(
                lev_ret_ladder.index
            ).fillna(0)
            benchmark_name = "100% 0050"
            data_note = "實際00631L日報酬；0050價格作為前高／回撤訊號。"

        elif test_period.startswith("2000"):
            mkt = actual["twii"]["Close"].loc[
                "2000-01-01":"2002-12-31"
            ].dropna()
            if mkt.empty or mkt.index.min() > pd.Timestamp("2000-01-31"):
                st.error("缺少2000年官方TAIEX，請先到「⑥ 資料管理」補齊。")
                st.stop()
            uret = mkt.pct_change().fillna(0)
            lev_ret_ladder = synthetic_2x_returns(
                uret, coef, calibration, vol_q, "校準平均", 42
            )
            market_level_ladder = mkt.reindex(lev_ret_ladder.index).ffill()
            benchmark_ret_ladder = uret
            benchmark_name = "100% TAIEX"
            data_note = "2000年沒有00631L，使用官方TAIEX＋2014年至今正2偏差模型。"

        else:
            mkt = actual["twii"]["Close"].loc[
                "2007-01-01":"2009-12-31"
            ].dropna()
            if mkt.empty or mkt.index.min() > pd.Timestamp("2007-01-31"):
                st.error("缺少2007–2009官方TAIEX，請先到「⑥ 資料管理」補齊。")
                st.stop()
            uret = mkt.pct_change().fillna(0)
            lev_ret_ladder = synthetic_2x_returns(
                uret, coef, calibration, vol_q, "校準平均", 42
            )
            market_level_ladder = mkt.reindex(lev_ret_ladder.index).ffill()
            benchmark_ret_ladder = uret
            benchmark_name = "100% TAIEX"
            data_note = "2008年沒有00631L，使用官方TAIEX＋2014年至今正2偏差模型。"

        st.info(data_note)

        ladder_df, ladder_events, ladder_stats = simulate_ladder_buy_strategy(
            lev_ret=lev_ret_ladder,
            market_level=market_level_ladder,
            capital=capital,
            initial_lev_target=ladder_initial_pct / 100,
            start_drawdown=ladder_start_dd_pct / 100,
            drawdown_step=ladder_dd_step_pct / 100,
            buy_step=ladder_buy_pct / 100,
            buy_basis=ladder_buy_basis,
            exit_style=ladder_exit_style,
            rebound_start=rebound_start_pct / 100,
            rebound_step=rebound_step_pct / 100,
            sell_step=sell_step_pct / 100,
            cash_yield=cash_yield,
            buy_cost=buy_cost,
            sell_cost=sell_cost,
        )

        bench_ladder = simulate_benchmark(
            benchmark_ret_ladder.reindex(ladder_df.index).fillna(0), capital
        )

        ladder_metric = metrics(
            ladder_df["Portfolio"],
            capital,
            ladder_stats["加碼次數"] + ladder_stats["反彈分批減碼次數"] + ladder_stats["完整重置次數"],
            ladder_stats["累計換手金額"],
        )
        bench_metric = metrics(bench_ladder["Portfolio"], capital)

        summary_ladder = pd.DataFrame(
            [bench_metric, ladder_metric],
            index=[benchmark_name, "越跌越買策略"]
        )

        st.dataframe(
            summary_ladder.style.format({
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

        c10, c11, c12, c13 = st.columns(4)
        c10.metric("加碼次數", f"{ladder_stats['加碼次數']}")
        c11.metric("反彈分批減碼", f"{ladder_stats['反彈分批減碼次數']}")
        c12.metric("完整重置", f"{ladder_stats['完整重置次數']}")
        c13.metric("現金用完次數", f"{ladder_stats['現金用完次數']}")

        if ladder_stats["首次現金用完日期"] is not None:
            st.warning(
                "此參數組合曾把現金全部投入。首次發生日期："
                f"{pd.Timestamp(ladder_stats['首次現金用完日期']).date()}"
            )

        st.line_chart(pd.concat([
            bench_ladder["Portfolio"].rename(benchmark_name),
            ladder_df["Portfolio"].rename("越跌越買策略"),
        ], axis=1))

        st.caption("投資組合回撤")
        st.line_chart(pd.concat([
            bench_ladder["Drawdown"].rename(benchmark_name),
            ladder_df["Drawdown"].rename("越跌越買策略"),
        ], axis=1))

        st.caption("正2占總資產比例")
        st.line_chart(
            (ladder_df[["LevWeight"]] * 100).rename(
                columns={"LevWeight": "正2占比（%）"}
            )
        )

        st.caption("現金部位")
        st.line_chart(ladder_df[["Cash"]])

        with st.expander("查看所有加碼／反彈減碼／重置事件"):
            if ladder_events.empty:
                st.info("此期間沒有觸發事件。")
            else:
                st.dataframe(
                    ladder_events.style.format({
                        "Market": "{:,.2f}",
                        "Drawdown": "{:.2%}",
                        "ReboundFromLow": "{:.2%}",
                        "Trigger": "{:.2%}",
                        "Amount": "{:,.0f}",
                        "Cost": "{:,.0f}",
                        "CashAfter": "{:,.0f}",
                        "LevAfter": "{:,.0f}",
                    }, na_rep=""),
                    use_container_width=True
                )

        with st.expander("策略規則摘要"):
            st.write(
                f"起始：正2 {ladder_initial_pct}%＋現金 {100-ladder_initial_pct}%。"
            )
            st.write(
                f"市場自前高回撤 {ladder_start_dd_pct}% 開始第一階加碼，"
                f"之後每再跌 {ladder_dd_step_pct}% 再加碼。"
            )
            st.write(
                f"每階投入：{ladder_buy_pct}% × {ladder_buy_basis}；"
                "現金可用到 0，但不使用融資。"
            )
            if ladder_exit_style == "回到前高一次再平衡":
                st.write("反彈：回到本輪下跌前高時，一次恢復原始配置。")
            elif ladder_exit_style == "低點反彈指定幅度一次再平衡":
                st.write(
                    f"反彈：自本輪最低點上漲 {rebound_start_pct}% 時，"
                    "一次恢復原始配置。"
                )
            else:
                st.write(
                    f"反彈：低點上漲 {rebound_start_pct}% 開始減碼，"
                    f"每再漲 {rebound_step_pct}% 再減碼一階；"
                    f"每階將當下總資產 {sell_step_pct}% 從正2轉回現金。"
                )
                st.write(
                    "若先回到本輪前高，尚未完成的部分會一次恢復原始配置。"
                )


# ----------------------------
# ⑥ 資料管理
# ----------------------------
with tabs[5]:
    st.subheader("本地市場資料庫")
    st.write(
        "一般回測只讀取 GitHub 專案內的 `data/market_daily.csv`。"
        "只有你在這個頁面按下更新按鈕時，才會連 Yahoo Finance。"
    )

    st.dataframe(status, use_container_width=True)

    taiex_rows = db.loc[db["Symbol"] == "TAIEX"].copy() if not db.empty else empty_database()
    if taiex_rows.empty:
        st.warning("TAIEX 尚無資料。2000／2008壓力測試目前不能執行。")
    else:
        taiex_start = pd.to_datetime(taiex_rows["Date"]).min()
        if taiex_start > pd.Timestamp("1999-01-31"):
            st.warning(
                f"TAIEX目前最早只有 {taiex_start.date()}，尚不足以測2000年股災。"
                "請按『補齊1999年至今官方TAIEX』。"
            )
        else:
            st.success(
                f"TAIEX早期歷史已涵蓋至 {taiex_start.date()}，"
                "可進行2000與2008壓力測試。"
            )

    if db.empty:
        st.warning("目前資料庫是空的。第一次部署後，請按「建立完整歷史資料庫」。")
    else:
        latest = db["Date"].max()
        st.success(f"本地資料庫目前最新日期：{latest.date().isoformat()}")

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("建立／重建完整資料庫", type="primary"):
            try:
                with st.spinner("建立資料庫並以證交所官方資料補齊TAIEX…"):
                    db = update_database(db, full_rebuild=True)
                st.success("完整歷史資料庫已建立。")
                st.rerun()
            except Exception as e:
                st.error(f"建立資料庫失敗：{e}")

    with c2:
        if st.button("只更新最新缺漏資料"):
            try:
                with st.spinner("更新近期資料並檢查TAIEX月份缺漏…"):
                    db = update_database(db, full_rebuild=False)
                st.success("資料庫已更新。")
                st.rerun()
            except Exception as e:
                st.error(f"更新失敗：{e}")

    with c3:
        if st.button("補齊1999年至今官方TAIEX"):
            try:
                with st.spinner("從證交所官方逐月補齊缺少的TAIEX月份…"):
                    db, added = ensure_official_taiex_history(
                        db, start="1999-01-01", end=date.today().isoformat()
                    )
                    save_local_database(db)
                st.success(f"TAIEX補齊完成，本次新增／補入 {added:,} 筆交易日資料。")
                st.rerun()
            except Exception as e:
                st.error(f"官方TAIEX補齊失敗：{e}")

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
    "v1.5：新增1999年至今長期模擬回測與可調投資開始日期；越跌越買與反彈分批減碼功能仍保留。"
    "歷史回測與模型最佳化均不代表未來報酬。"
)
