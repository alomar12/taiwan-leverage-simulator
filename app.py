
from pathlib import Path
from datetime import date
import io
import time
import requests

import numpy as np
import pandas as pd
import streamlit as st
import yfinance as yf
import plotly.graph_objects as go

st.set_page_config(page_title="台灣50正2再平衡模擬器 v1.8", layout="wide")

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DB_FILE = DATA_DIR / "market_daily.csv"
FNG_FILE = DATA_DIR / "fear_greed_daily.csv"

FNG_COLUMNS = ["Date", "FearGreed", "Rating", "Source"]
FNG_COMBINED_URL = (
    "https://raw.githubusercontent.com/whit3rabbit/"
    "fear-greed-data/main/fear-greed.csv"
)
FNG_RECON_END = pd.Timestamp("2021-01-29")
FNG_CNN_START = pd.Timestamp("2021-02-01")

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
# CNN Fear & Greed 本地資料庫
# ============================================================
def empty_fng_database():
    return pd.DataFrame(columns=FNG_COLUMNS)


@st.cache_data(show_spinner=False)
def read_local_fng(path_str: str, modified_ns: int = 0):
    path = Path(path_str)
    if not path.exists() or path.stat().st_size == 0:
        return empty_fng_database()

    try:
        df = pd.read_csv(path)
    except pd.errors.EmptyDataError:
        return empty_fng_database()

    if df.empty:
        return empty_fng_database()

    # 相容舊/外部欄位名稱
    rename_map = {
        "Fear Greed": "FearGreed",
        "fear_and_greed_index": "FearGreed",
        "date": "Date",
        "rating": "Rating",
        "source": "Source",
    }
    df = df.rename(columns=rename_map)

    if "Date" not in df.columns or "FearGreed" not in df.columns:
        raise RuntimeError("Fear & Greed 資料必須至少包含 Date、FearGreed。")

    if "Rating" not in df.columns:
        df["Rating"] = ""
    if "Source" not in df.columns:
        df["Source"] = np.where(
            pd.to_datetime(df["Date"], errors="coerce") <= FNG_RECON_END,
            "reconstructed",
            "cnn_official_via_mirror",
        )

    df = df[FNG_COLUMNS].copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["FearGreed"] = pd.to_numeric(df["FearGreed"], errors="coerce")
    df["Rating"] = df["Rating"].fillna("").astype(str).str.lower()
    df["Source"] = df["Source"].fillna("").astype(str)
    df = (
        df.dropna(subset=["Date", "FearGreed"])
          .drop_duplicates("Date", keep="last")
          .sort_values("Date")
          .reset_index(drop=True)
    )
    return df


def get_local_fng():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    modified = FNG_FILE.stat().st_mtime_ns if FNG_FILE.exists() else 0
    return read_local_fng(str(FNG_FILE), modified)


def fng_rating(score):
    if pd.isna(score):
        return ""
    score = float(score)
    if score < 25:
        return "extreme fear"
    if score < 45:
        return "fear"
    if score < 55:
        return "neutral"
    if score < 75:
        return "greed"
    return "extreme greed"


def save_local_fng(df: pd.DataFrame):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    out = df.copy()
    if out.empty:
        out = empty_fng_database()

    for c in FNG_COLUMNS:
        if c not in out.columns:
            if c == "Rating":
                out[c] = out.get("FearGreed", pd.Series(dtype=float)).apply(fng_rating)
            elif c == "Source":
                out[c] = np.where(
                    pd.to_datetime(out.get("Date"), errors="coerce") <= FNG_RECON_END,
                    "reconstructed",
                    "cnn_official_via_mirror",
                )
            else:
                out[c] = np.nan

    out = out[FNG_COLUMNS].copy()
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    out["FearGreed"] = pd.to_numeric(out["FearGreed"], errors="coerce")
    out["Rating"] = out["Rating"].fillna("").astype(str).str.lower()
    out["Source"] = out["Source"].fillna("").astype(str)
    out = (
        out.dropna(subset=["Date", "FearGreed"])
           .drop_duplicates("Date", keep="last")
           .sort_values("Date")
    )
    out["Date"] = out["Date"].dt.strftime("%Y-%m-%d")
    out.to_csv(FNG_FILE, index=False, encoding="utf-8-sig")
    st.cache_data.clear()


@st.cache_data(ttl=1800, show_spinner=False)
def download_fng_combined():
    """
    下載 whit3rabbit/fear-greed-data 的 canonical combined CSV。
    2011-01-03 ~ 2021-01-29：第三方歷史重建
    2021-02-01 ~ 現在：該專案由 CNN 現行端點每日更新的尾端
    """
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 Chrome/124 Safari/537.36"
        ),
        "Accept": "text/csv,text/plain,*/*",
    }
    resp = requests.get(FNG_COMBINED_URL, headers=headers, timeout=30)
    resp.raise_for_status()

    raw = pd.read_csv(io.StringIO(resp.text))
    raw = raw.rename(columns={
        "Fear Greed": "FearGreed",
        "fear_and_greed_index": "FearGreed",
        "date": "Date",
        "rating": "Rating",
    })
    if "Date" not in raw.columns or "FearGreed" not in raw.columns:
        raise RuntimeError("下載的 Fear & Greed CSV 欄位格式不符預期。")

    if "Rating" not in raw.columns:
        raw["Rating"] = ""

    raw["Date"] = pd.to_datetime(raw["Date"], errors="coerce")
    raw["FearGreed"] = pd.to_numeric(raw["FearGreed"], errors="coerce")
    raw["Rating"] = raw["Rating"].fillna("").astype(str).str.lower()
    raw.loc[raw["Rating"].eq(""), "Rating"] = (
        raw.loc[raw["Rating"].eq(""), "FearGreed"].apply(fng_rating)
    )
    raw["Source"] = np.where(
        raw["Date"] <= FNG_RECON_END,
        "reconstructed",
        "cnn_official_via_mirror",
    )
    raw = (
        raw[FNG_COLUMNS]
        .dropna(subset=["Date", "FearGreed"])
        .drop_duplicates("Date", keep="last")
        .sort_values("Date")
        .reset_index(drop=True)
    )
    return raw


def align_fng_to_taiwan_dates(tw_dates, fng_df):
    """
    重要：CNN F&G 是美國收盤後才知道。
    台灣同一曆日交易發生在美國該日收盤之前，所以不能用同日 F&G。
    對每個台灣交易日，只使用「嚴格早於該日期」的最新 F&G。
    """
    idx = pd.DatetimeIndex(pd.to_datetime(tw_dates)).sort_values()
    left = pd.DataFrame({"TWDate": idx})
    if fng_df is None or fng_df.empty:
        out = left.copy()
        out["FNGDate"] = pd.NaT
        out["FearGreed"] = np.nan
        out["Rating"] = ""
        out["Source"] = ""
        return out.set_index("TWDate")

    right = fng_df.copy().sort_values("Date").rename(columns={"Date": "FNGDate"})
    merged = pd.merge_asof(
        left.sort_values("TWDate"),
        right.sort_values("FNGDate"),
        left_on="TWDate",
        right_on="FNGDate",
        direction="backward",
        allow_exact_matches=False,
    )
    return merged.set_index("TWDate")


def fng_multiplier_from_score(score, multipliers):
    if pd.isna(score):
        return 1.0
    s = float(score)
    if s < 25:
        return float(multipliers["extreme fear"])
    if s < 45:
        return float(multipliers["fear"])
    if s < 55:
        return float(multipliers["neutral"])
    if s < 75:
        return float(multipliers["greed"])
    return float(multipliers["extreme greed"])


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
    rows.append((
        first_dt, capital, lev, cash, target, "起始",
        0.0, "起始", lev, cash
    ))

    for i, dt in enumerate(idx[1:], start=1):
        lev *= 1 + float(r.loc[dt])
        cash *= 1 + daily_cash
        total = max(lev + cash, 1e-12)
        weight = lev / total
        rel_price = float(level.loc[dt]) / anchor - 1 if anchor > 0 else 0

        trigger = False
        reason = ""
        trade_amount_signed = 0.0
        trade_direction = ""
        lev_before_trade = lev
        cash_before_trade = cash

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
            lev_before_trade = lev
            cash_before_trade = cash
            lev, cash, turnover = rebalance_to_target(
                lev, cash, target, buy_cost, sell_cost
            )
            trade_amount_signed = lev - lev_before_trade
            if trade_amount_signed > 1e-8:
                trade_direction = "買進正2"
            elif trade_amount_signed < -1e-8:
                trade_direction = "賣出正2"
            else:
                trade_direction = "無調整"
            turnover_sum += turnover
            n_rebal += 1
            last_rebal_i = i
            anchor = float(level.loc[dt])
            total = lev + cash
            weight = lev / total if total > 0 else 0

        rows.append((
            dt, lev + cash, lev, cash, weight, reason,
            trade_amount_signed, trade_direction,
            lev_before_trade, cash_before_trade
        ))

    df = pd.DataFrame(
        rows,
        columns=[
            "Date", "Portfolio", "Leveraged", "Cash", "LevWeight", "Reason",
            "TradeAmountSigned", "TradeDirection", "LevBefore", "CashBefore"
        ]
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
    fng_signal=None,
    fng_multipliers=None,
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

    if fng_signal is None:
        fng = pd.Series(index=idx, data=np.nan, dtype=float)
    else:
        fng = pd.Series(fng_signal).reindex(idx)

    if fng_multipliers is None:
        fng_multipliers = {
            "extreme fear": 1.0,
            "fear": 1.0,
            "neutral": 1.0,
            "greed": 1.0,
            "extreme greed": 1.0,
        }

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
                "TradeDirection": "條件啟動",
                "Market": mkt,
                "Drawdown": drawdown,
                "ReboundFromLow": 0.0,
                "Amount": 0.0,
                "LevBefore": lev,
                "CashBefore": cash,
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
                    base_buy = capital * buy_step
                else:
                    base_buy = total_before * buy_step

                fng_score = fng.loc[dt] if dt in fng.index else np.nan
                fng_mult = fng_multiplier_from_score(
                    fng_score, fng_multipliers
                )
                desired_buy = base_buy * fng_mult

                lev_before_event = lev
                cash_before_event = cash
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
                    "TradeDirection": "買進正2",
                    "Market": mkt,
                    "Drawdown": -cycle_dd_abs,
                    "ReboundFromLow": mkt / cycle_low - 1 if cycle_low > 0 else 0.0,
                    "Trigger": next_buy_trigger,
                    "FNG": fng_score,
                    "FNGMultiplier": fng_mult,
                    "BaseAmount": base_buy,
                    "Amount": bought,
                    "Cost": cost,
                    "LevBefore": lev_before_event,
                    "CashBefore": cash_before_event,
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
                        "TradeDirection": "現金用完",
                        "Market": mkt,
                        "Drawdown": -cycle_dd_abs,
                        "ReboundFromLow": mkt / cycle_low - 1 if cycle_low > 0 else 0.0,
                        "Amount": 0.0,
                        "LevBefore": lev,
                        "CashBefore": cash,
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

                    lev_before_event = lev
                    cash_before_event = cash
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
                        "TradeDirection": "賣出正2",
                        "Market": mkt,
                        "Drawdown": mkt / cycle_peak - 1 if cycle_peak > 0 else 0.0,
                        "ReboundFromLow": rebound,
                        "Trigger": next_sell_trigger,
                        "Amount": sold,
                        "Cost": cost,
                        "LevBefore": lev_before_event,
                        "CashBefore": cash_before_event,
                        "CashAfter": cash,
                        "LevAfter": lev,
                    })

                    next_sell_trigger += rebound_step

                # 回到前高，不論前面分批完成多少，都強制回原始配置
                if mkt >= cycle_peak:
                    do_full_reset = True
                    reset_reason = "回到前高強制完成"

            if do_full_reset:
                lev_before_event = lev
                cash_before_event = cash
                lev, cash, traded = rebalance_to_initial_target(
                    lev, cash, initial_lev_target, buy_cost, sell_cost
                )
                reset_direction = (
                    "買進正2" if lev > lev_before_event + 1e-8
                    else "賣出正2" if lev < lev_before_event - 1e-8
                    else "無調整"
                )
                turnover += traded
                n_resets += 1
                reason_parts.append(f"{reset_reason}→恢復原始配置")

                events.append({
                    "Date": dt,
                    "Event": f"{reset_reason}→恢復原始配置",
                    "TradeDirection": reset_direction,
                    "Market": mkt,
                    "Drawdown": mkt / cycle_peak - 1 if cycle_peak > 0 else 0.0,
                    "ReboundFromLow": rebound,
                    "Amount": traded,
                    "LevBefore": lev_before_event,
                    "CashBefore": cash_before_event,
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
# 互動圖表與日期查詢
# ============================================================
def nearest_trading_date(index, requested_date):
    idx = pd.DatetimeIndex(index).sort_values()
    if len(idx) == 0:
        return None
    target = pd.Timestamp(requested_date)
    pos = idx.get_indexer([target], method="nearest")[0]
    if pos < 0:
        return None
    return idx[pos]


def _selected_x_from_plotly_event(event):
    """相容不同 Streamlit Plotly selection 回傳型態。"""
    try:
        points = event.selection.points
        if points:
            x = points[0].get("x")
            if x is not None:
                return pd.Timestamp(x)
    except Exception:
        pass

    try:
        points = event.get("selection", {}).get("points", [])
        if points:
            x = points[0].get("x")
            if x is not None:
                return pd.Timestamp(x)
    except Exception:
        pass
    return None


def portfolio_display_series(series, mode, initial_capital):
    s = series.astype(float).copy()
    if mode == "基準化（起點=100）":
        first = s.dropna().iloc[0]
        return 100 * s / first if first != 0 else s
    if mode == "累積報酬（%）":
        return (s / initial_capital - 1) * 100
    return s



def build_fixed_rebalance_events(strategy_df):
    """把固定比例再平衡結果轉成統一事件格式。"""
    if strategy_df is None or strategy_df.empty or "Reason" not in strategy_df.columns:
        return pd.DataFrame()
    x = strategy_df.loc[
        strategy_df["Reason"].isin(["下跌門檻", "上漲門檻", "低於比例帶", "高於比例帶"])
    ].copy()
    if x.empty:
        return pd.DataFrame()
    ev = pd.DataFrame(index=x.index)
    ev["Event"] = x["Reason"]
    ev["TradeDirection"] = x["TradeDirection"]
    ev["Amount"] = x["TradeAmountSigned"].abs()
    ev["LevBefore"] = x["LevBefore"]
    ev["CashBefore"] = x["CashBefore"]
    ev["LevAfter"] = x["Leveraged"]
    ev["CashAfter"] = x["Cash"]
    return ev


def _event_category(event_name, direction):
    name = str(event_name or "")
    direction = str(direction or "")
    if "現金用完" in name or direction == "現金用完":
        return "現金用完"
    if "恢復原始配置" in name or "重置" in name:
        return "完整重置"
    if direction == "買進正2" or "加碼" in name or "下跌" in name or "低於比例" in name:
        return "買進／加碼"
    if direction == "賣出正2" or "減碼" in name or "上漲" in name or "高於比例" in name:
        return "賣出／減碼"
    return "條件啟動"


def prepare_event_markers(event_df, strategy_df):
    """同一天多個階梯事件在圖上合併成一個標記，明細仍逐筆保留。"""
    if event_df is None or len(event_df) == 0:
        return pd.DataFrame()
    ev = event_df.copy()
    if "Date" in ev.columns:
        ev["Date"] = pd.to_datetime(ev["Date"])
        ev = ev.set_index("Date")
    ev.index = pd.to_datetime(ev.index)

    rows = []
    priority = {"完整重置": 5, "現金用完": 4, "買進／加碼": 3, "賣出／減碼": 2, "條件啟動": 1}
    for dt, grp in ev.groupby(ev.index):
        if dt not in strategy_df.index:
            continue
        names = grp["Event"].astype(str).tolist() if "Event" in grp else [""] * len(grp)
        dirs = grp["TradeDirection"].astype(str).tolist() if "TradeDirection" in grp else [""] * len(grp)
        cats = [_event_category(n, d) for n, d in zip(names, dirs)]
        cat = max(cats, key=lambda c: priority.get(c, 0)) if cats else "條件啟動"
        amount = pd.to_numeric(grp["Amount"], errors="coerce").fillna(0).sum() if "Amount" in grp else 0.0
        first, last = grp.iloc[0], grp.iloc[-1]
        srow = strategy_df.loc[dt]
        rows.append({
            "Date": dt,
            "Category": cat,
            "EventSummary": "；".join(names),
            "Amount": float(amount),
            "LevBefore": first.get("LevBefore", np.nan),
            "CashBefore": first.get("CashBefore", np.nan),
            "LevAfter": last.get("LevAfter", srow.get("Leveraged", np.nan)),
            "CashAfter": last.get("CashAfter", srow.get("Cash", np.nan)),
            "LevWeightAfter": srow.get("LevWeight", np.nan),
            "Count": len(grp),
        })
    return pd.DataFrame(rows).set_index("Date").sort_index() if rows else pd.DataFrame()


def render_portfolio_chart(
    strategy_df,
    benchmark_series,
    benchmark_name,
    strategy_name,
    initial_capital,
    key_prefix,
    default_query_date=None,
    event_df=None,
):
    """
    互動資產圖：
    - hover 顯示總資產、正2、現金、正2占比
    - 點擊任一點後固定顯示當日明細
    - 可手動查詢日期（自動抓最近交易日）
    - 線性 / 對數 / 起點100 / 累積報酬% 四種尺度
    """
    display_mode = st.radio(
        "資產走勢顯示方式",
        ["金額（線性）", "金額（對數）", "基準化（起點=100）", "累積報酬（%）"],
        horizontal=True,
        key=f"{key_prefix}_display_mode",
    )

    strat_raw = strategy_df["Portfolio"].astype(float)
    bench_raw = benchmark_series.astype(float).reindex(strat_raw.index)

    strat_y = portfolio_display_series(strat_raw, display_mode, initial_capital)
    bench_y = portfolio_display_series(bench_raw, display_mode, initial_capital)

    custom = np.column_stack([
        strategy_df["Portfolio"].values,
        strategy_df["Leveraged"].values,
        strategy_df["Cash"].values,
        strategy_df["LevWeight"].values * 100,
    ])

    if display_mode == "累積報酬（%）":
        y_label = "累積報酬（%）"
        y_hover = "%{y:.2f}%"
    elif display_mode == "基準化（起點=100）":
        y_label = "指數化數值"
        y_hover = "%{y:.2f}"
    else:
        y_label = "資產金額"
        y_hover = "%{y:,.0f}"

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=strat_y.index,
        y=strat_y.values,
        mode="lines",
        name=strategy_name,
        customdata=custom,
        hovertemplate=(
            "<b>%{x|%Y-%m-%d}</b><br>"
            + f"{strategy_name}：" + y_hover + "<br>"
            "總資產：%{customdata[0]:,.0f}<br>"
            "正2：%{customdata[1]:,.0f}<br>"
            "現金：%{customdata[2]:,.0f}<br>"
            "正2占比：%{customdata[3]:.2f}%"
            "<extra></extra>"
        ),
    ))

    fig.add_trace(go.Scatter(
        x=bench_y.index,
        y=bench_y.values,
        mode="lines",
        name=benchmark_name,
        customdata=bench_raw.values.reshape(-1, 1),
        hovertemplate=(
            "<b>%{x|%Y-%m-%d}</b><br>"
            + f"{benchmark_name}：" + y_hover + "<br>"
            "基準實際資產：%{customdata[0]:,.0f}"
            "<extra></extra>"
        ),
    ))

    marker_events = prepare_event_markers(event_df, strategy_df)
    show_event_markers = False
    if not marker_events.empty:
        show_event_markers = st.checkbox(
            "顯示再平衡／資金調整事件記號",
            value=True,
            key=f"{key_prefix}_show_event_markers",
        )

    if show_event_markers:
        category_style = {
            "買進／加碼": ("triangle-up", 12),
            "賣出／減碼": ("triangle-down", 12),
            "完整重置": ("diamond", 13),
            "現金用完": ("x", 13),
            "條件啟動": ("circle-open", 11),
        }
        for category, grp in marker_events.groupby("Category"):
            symbol, size = category_style.get(category, ("circle", 11))
            event_y = strat_y.reindex(grp.index)
            cdata = np.column_stack([
                grp["EventSummary"].astype(str).values,
                grp["Amount"].fillna(0).values,
                grp["LevBefore"].fillna(np.nan).values,
                grp["CashBefore"].fillna(np.nan).values,
                grp["LevAfter"].fillna(np.nan).values,
                grp["CashAfter"].fillna(np.nan).values,
                grp["LevWeightAfter"].fillna(np.nan).values * 100,
                grp["Count"].values,
            ])
            fig.add_trace(go.Scatter(
                x=grp.index,
                y=event_y.values,
                mode="markers",
                name=f"事件｜{category}",
                marker=dict(symbol=symbol, size=size, line=dict(width=1.5)),
                customdata=cdata,
                hovertemplate=(
                    "<b>%{x|%Y-%m-%d}｜" + category + "</b><br>"
                    "觸發：%{customdata[0]}<br>"
                    "當日調整金額合計：%{customdata[1]:,.0f}<br>"
                    "調整前正2：%{customdata[2]:,.0f}<br>"
                    "調整前現金：%{customdata[3]:,.0f}<br>"
                    "調整後正2：%{customdata[4]:,.0f}<br>"
                    "調整後現金：%{customdata[5]:,.0f}<br>"
                    "調整後正2占比：%{customdata[6]:.2f}%<br>"
                    "當日事件筆數：%{customdata[7]:.0f}"
                    "<extra></extra>"
                ),
            ))

    fig.update_layout(
        height=560,
        hovermode="x unified",
        margin=dict(l=10, r=10, t=30, b=10),
        legend=dict(orientation="h"),
        yaxis_title=y_label,
        xaxis_title="日期",
    )
    if display_mode == "金額（對數）":
        fig.update_yaxes(type="log")

    fig.update_xaxes(
        rangeslider_visible=True,
        rangeselector=dict(
            buttons=[
                dict(count=1, label="1年", step="year", stepmode="backward"),
                dict(count=5, label="5年", step="year", stepmode="backward"),
                dict(count=10, label="10年", step="year", stepmode="backward"),
                dict(count=20, label="20年", step="year", stepmode="backward"),
                dict(step="all", label="全部"),
            ]
        )
    )

    event = st.plotly_chart(
        fig,
        use_container_width=True,
        key=f"{key_prefix}_plot",
        on_select="rerun",
        selection_mode="points",
    )

    clicked_date = _selected_x_from_plotly_event(event)

    min_d = strategy_df.index.min().date()
    max_d = strategy_df.index.max().date()
    if default_query_date is None:
        default_query_date = max_d
    else:
        default_query_date = pd.Timestamp(default_query_date).date()
        default_query_date = min(max(default_query_date, min_d), max_d)

    qcol1, qcol2 = st.columns([1, 2])
    with qcol1:
        query_date = st.date_input(
            "查詢某一天",
            value=default_query_date,
            min_value=min_d,
            max_value=max_d,
            key=f"{key_prefix}_query_date",
        )
    with qcol2:
        if clicked_date is not None:
            st.info(
                f"已點選圖表日期：{clicked_date.date()}。"
                "下方明細優先顯示點選日期；重新點其他資料點即可切換。"
            )
        else:
            st.caption("可點圖上的資料點，或用左側日期欄查詢。")

    requested = clicked_date if clicked_date is not None else pd.Timestamp(query_date)
    actual_dt = nearest_trading_date(strategy_df.index, requested)

    if actual_dt is not None:
        row = strategy_df.loc[actual_dt]
        bench_val = bench_raw.loc[actual_dt] if actual_dt in bench_raw.index else np.nan

        if pd.Timestamp(requested).normalize() != actual_dt.normalize():
            st.caption(
                f"{pd.Timestamp(requested).date()} 非交易日或無資料，"
                f"顯示最近交易日 {actual_dt.date()}。"
            )

        d1, d2, d3, d4, d5 = st.columns(5)
        d1.metric("日期", actual_dt.strftime("%Y/%m/%d"))
        d2.metric("總資產", f"{row['Portfolio']:,.0f}")
        d3.metric("正2", f"{row['Leveraged']:,.0f}")
        d4.metric("現金", f"{row['Cash']:,.0f}")
        d5.metric("正2占比", f"{row['LevWeight']:.2%}")

        extra1, extra2 = st.columns(2)
        extra1.metric(benchmark_name, f"{bench_val:,.0f}" if pd.notna(bench_val) else "—")
        if "Drawdown" in strategy_df.columns:
            extra2.metric("策略自高點回撤", f"{row['Drawdown']:.2%}")

        if event_df is not None and len(event_df) > 0:
            ev = event_df.copy()
            if "Date" in ev.columns:
                ev["Date"] = pd.to_datetime(ev["Date"])
                ev = ev.set_index("Date")
            ev.index = pd.to_datetime(ev.index)
            day_events = ev.loc[ev.index.normalize() == actual_dt.normalize()].copy()
            if not day_events.empty:
                st.markdown("**當日資金調整事件**")
                show_cols = [c for c in [
                    "Event", "TradeDirection", "Amount", "LevBefore", "CashBefore",
                    "LevAfter", "CashAfter", "Trigger", "FNG", "FNGMultiplier"
                ] if c in day_events.columns]
                fmt = {
                    "Amount": "{:,.0f}", "LevBefore": "{:,.0f}", "CashBefore": "{:,.0f}",
                    "LevAfter": "{:,.0f}", "CashAfter": "{:,.0f}", "Trigger": "{:.2%}",
                    "FNG": "{:.1f}", "FNGMultiplier": "{:.2f}×"
                }
                st.dataframe(
                    day_events[show_cols].style.format({k:v for k,v in fmt.items() if k in show_cols}, na_rep=""),
                    use_container_width=True,
                )

    return clicked_date


def render_long_market_chart(long_chart, key_prefix):
    """長期大盤/模擬正2/實際正2比較，重點解決早期振幅被壓扁。"""
    scale = st.radio(
        "長期走勢尺度",
        ["線性", "對數（建議）", "各線起點=100"],
        index=1,
        horizontal=True,
        key=f"{key_prefix}_scale",
    )

    plot_df = long_chart.copy()
    if scale == "各線起點=100":
        for col in plot_df.columns:
            valid = plot_df[col].dropna()
            if not valid.empty and float(valid.iloc[0]) != 0:
                plot_df[col] = 100 * plot_df[col] / float(valid.iloc[0])

    fig = go.Figure()
    for col in plot_df.columns:
        fig.add_trace(go.Scatter(
            x=plot_df.index,
            y=plot_df[col],
            mode="lines",
            name=col,
            hovertemplate="<b>%{x|%Y-%m-%d}</b><br>" + col + "：%{y:,.2f}<extra></extra>",
        ))

    fig.update_layout(
        height=560,
        hovermode="x unified",
        margin=dict(l=10, r=10, t=30, b=10),
        legend=dict(orientation="h"),
        xaxis_title="日期",
        yaxis_title="指數值",
    )
    if scale == "對數（建議）":
        fig.update_yaxes(type="log")

    fig.update_xaxes(
        rangeslider_visible=True,
        rangeselector=dict(
            buttons=[
                dict(count=5, label="5年", step="year", stepmode="backward"),
                dict(count=10, label="10年", step="year", stepmode="backward"),
                dict(count=20, label="20年", step="year", stepmode="backward"),
                dict(step="all", label="全部"),
            ]
        )
    )
    st.plotly_chart(fig, use_container_width=True, key=f"{key_prefix}_long_plot")


# ============================================================
# UI
# ============================================================
st.title("台灣50正2 × 現金：再平衡與股災壓力測試模擬器 v1.8")
st.caption(
    "v1.8：互動圖新增再平衡事件記號，可直接查看每次買進、減碼、重置的日期、"
    "調整金額與正2／現金前後變化。"
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
fng_db = get_local_fng()

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
    "⑥ Fear & Greed研究",
    "⑦ 資料管理",
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
        st.warning(f"{prep_error} 請先到「⑦ 資料管理」建立或匯入資料庫。")
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

        st.subheader("互動資產走勢")
        fixed_events = build_fixed_rebalance_events(strat)
        render_portfolio_chart(
            strategy_df=strat,
            benchmark_series=bench["Portfolio"],
            benchmark_name=bench_name,
            strategy_name=lev_name,
            initial_capital=capital,
            key_prefix="main_portfolio",
            default_query_date=actual_start,
            event_df=fixed_events,
        )

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
            render_long_market_chart(long_chart, "main_long_history")

            # 顯示2014後模型與實際00631L的校驗差距
            compare = long_data["chart"][["模擬正2", "實際00631L（上市後）"]].dropna()
            if not compare.empty:
                compare = compare / compare.iloc[0] * 100
                st.caption("2014/10/31後：模擬正2 vs 實際00631L（同日起點=100）")
                render_long_market_chart(compare, "main_model_actual_compare")

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

        trades = strat.loc[
            strat["Reason"].isin(["下跌門檻", "上漲門檻", "低於比例帶", "高於比例帶"]),
            ["Portfolio", "Leveraged", "Cash", "LevWeight", "Reason", "TradeDirection",
             "TradeAmountSigned", "LevBefore", "CashBefore"]
        ].copy()
        trades["調整金額"] = trades["TradeAmountSigned"].abs()
        trades = trades.rename(columns={
            "Portfolio": "總資產", "LevWeight": "正2占比", "Reason": "觸發條件",
            "TradeDirection": "調整方向", "LevBefore": "調整前正2", "CashBefore": "調整前現金",
            "Leveraged": "調整後正2", "Cash": "調整後現金"
        })
        with st.expander(f"再平衡紀錄（{len(trades)}次）"):
            st.dataframe(
                trades[["總資產", "正2占比", "觸發條件", "調整方向", "調整金額",
                        "調整前正2", "調整前現金", "調整後正2", "調整後現金"]].style.format({
                    "總資產": "{:,.0f}", "正2占比": "{:.2%}", "調整金額": "{:,.0f}",
                    "調整前正2": "{:,.0f}", "調整前現金": "{:,.0f}",
                    "調整後正2": "{:,.0f}", "調整後現金": "{:,.0f}"
                }),
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
                st.error("本地資料庫缺少2000年TAIEX。請到「⑦ 資料管理」按『補齊1999年至今官方TAIEX』。")
                st.stop()
            ur = base.pct_change().dropna()
            br = ur.copy()
            st.info("2000年沒有0050與00631L；使用證交所官方TAIEX重建歷史正2，因此屬壓力測試。")
        else:
            base = actual["twii"]["Close"].loc["2007-01-01":"2009-12-31"].dropna()
            if base.empty or base.index.min() > pd.Timestamp("2007-01-31"):
                st.error("本地資料庫缺少2007–2009年TAIEX。請到「⑦ 資料管理」按『補齊1999年至今官方TAIEX』。")
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

        st.markdown("#### Fear & Greed 加碼倍率（可選）")
        use_fng_ladder = st.checkbox(
            "使用前一個已知 CNN Fear & Greed 調整每階加碼金額",
            value=False,
            key="ladder_use_fng"
        )

        fng_mults = {
            "extreme fear": 1.0,
            "fear": 1.0,
            "neutral": 1.0,
            "greed": 1.0,
            "extreme greed": 1.0,
        }

        if use_fng_ladder:
            if fng_db.empty:
                st.warning(
                    "Fear & Greed 本地資料庫尚未建立。"
                    "請先到「⑦ 資料管理」按『建立／更新 Fear & Greed』。"
                )
            else:
                m1, m2, m3, m4, m5 = st.columns(5)
                with m1:
                    fng_mults["extreme fear"] = st.number_input(
                        "極度恐懼 <25", 0.0, 3.0, 1.50, 0.10,
                        key="fng_mult_ef"
                    )
                with m2:
                    fng_mults["fear"] = st.number_input(
                        "恐懼 25–44", 0.0, 3.0, 1.20, 0.10,
                        key="fng_mult_f"
                    )
                with m3:
                    fng_mults["neutral"] = st.number_input(
                        "中性 45–54", 0.0, 3.0, 1.00, 0.10,
                        key="fng_mult_n"
                    )
                with m4:
                    fng_mults["greed"] = st.number_input(
                        "貪婪 55–74", 0.0, 3.0, 0.70, 0.10,
                        key="fng_mult_g"
                    )
                with m5:
                    fng_mults["extreme greed"] = st.number_input(
                        "極度貪婪 ≥75", 0.0, 3.0, 0.50, 0.10,
                        key="fng_mult_eg"
                    )
                st.caption(
                    "時間對齊採嚴格前一個已知美國交易日；"
                    "台灣同日尚未公布的美國收盤F&G不會被使用。"
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
                st.error("缺少2000年官方TAIEX，請先到「⑦ 資料管理」補齊。")
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
                st.error("缺少2007–2009官方TAIEX，請先到「⑦ 資料管理」補齊。")
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

        ladder_fng_signal = None
        if use_fng_ladder and not fng_db.empty:
            aligned_ladder_fng = align_fng_to_taiwan_dates(
                market_level_ladder.index, fng_db
            )
            ladder_fng_signal = aligned_ladder_fng["FearGreed"]

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
            fng_signal=ladder_fng_signal,
            fng_multipliers=fng_mults,
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

        st.subheader("互動資產走勢與日期查詢")
        render_portfolio_chart(
            strategy_df=ladder_df,
            benchmark_series=bench_ladder["Portfolio"],
            benchmark_name=benchmark_name,
            strategy_name="越跌越買策略",
            initial_capital=capital,
            key_prefix="ladder_portfolio",
            default_query_date=ladder_df.index.min(),
            event_df=ladder_events,
        )

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
                        "FNG": "{:.1f}",
                        "FNGMultiplier": "{:.2f}×",
                        "BaseAmount": "{:,.0f}",
                        "Amount": "{:,.0f}",
                        "Cost": "{:,.0f}",
                        "LevBefore": "{:,.0f}",
                        "CashBefore": "{:,.0f}",
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
                f"每階基本投入：{ladder_buy_pct}% × {ladder_buy_basis}；"
                "現金可用到 0，但不使用融資。"
            )
            if use_fng_ladder:
                st.write(
                    "Fear & Greed倍率："
                    f"極度恐懼 {fng_mults['extreme fear']:.2f}×、"
                    f"恐懼 {fng_mults['fear']:.2f}×、"
                    f"中性 {fng_mults['neutral']:.2f}×、"
                    f"貪婪 {fng_mults['greed']:.2f}×、"
                    f"極度貪婪 {fng_mults['extreme greed']:.2f}×。"
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
# ⑥ Fear & Greed 研究
# ----------------------------
with tabs[5]:
    st.subheader("CNN Fear & Greed｜台灣市場研究")

    if fng_db.empty:
        st.warning(
            "Fear & Greed 資料庫尚未建立。請到「⑦ 資料管理」"
            "按『建立／更新 Fear & Greed』。"
        )
    else:
        fng_start = fng_db["Date"].min()
        fng_end = fng_db["Date"].max()
        reconstructed_n = int((fng_db["Source"] == "reconstructed").sum())
        cnn_n = int((fng_db["Source"] == "cnn_official_via_mirror").sum())

        a, b, c, d = st.columns(4)
        a.metric("最早日期", f"{fng_start.date()}")
        b.metric("最新日期", f"{fng_end.date()}")
        c.metric("第三方重建筆數", f"{reconstructed_n:,}")
        d.metric("CNN端點尾端筆數", f"{cnn_n:,}")

        st.info(
            "2011/01/03–2021/01/29 標為第三方歷史重建；"
            "2021/02/01之後的 canonical dataset 由維護專案"
            "從 CNN 現行端點更新。"
        )

        chart_df = fng_db.set_index("Date")[["FearGreed"]].rename(
            columns={"FearGreed": "Fear & Greed"}
        )
        st.line_chart(chart_df)

        if has_required_data:
            st.markdown("### Fear & Greed 與台股後續報酬")
            market_choice = st.radio(
                "台股研究標的",
                ["TAIEX", "0050"],
                horizontal=True,
                key="fng_market_choice"
            )

            if market_choice == "TAIEX":
                mkt_series = actual["twii"]["Close"].dropna()
            else:
                mkt_series = actual["tr0050"].dropna()

            aligned = align_fng_to_taiwan_dates(mkt_series.index, fng_db)
            study = pd.DataFrame(index=mkt_series.index)
            study["Market"] = mkt_series
            study = study.join(
                aligned[["FNGDate", "FearGreed", "Rating", "Source"]],
                how="left"
            )

            for h in [20, 60, 120]:
                study[f"Fwd{h}"] = study["Market"].shift(-h) / study["Market"] - 1

            def band_from_score(s):
                if pd.isna(s):
                    return np.nan
                if s < 25:
                    return "0–24 極度恐懼"
                if s < 45:
                    return "25–44 恐懼"
                if s < 55:
                    return "45–54 中性"
                if s < 75:
                    return "55–74 貪婪"
                return "75–100 極度貪婪"

            study["Band"] = study["FearGreed"].apply(band_from_score)
            ordered_bands = [
                "0–24 極度恐懼",
                "25–44 恐懼",
                "45–54 中性",
                "55–74 貪婪",
                "75–100 極度貪婪",
            ]
            grouped = study.dropna(subset=["Band"]).groupby("Band").agg(
                樣本數=("FearGreed", "size"),
                FNG平均=("FearGreed", "mean"),
                後20日平均=("Fwd20", "mean"),
                後20日中位數=("Fwd20", "median"),
                後60日平均=("Fwd60", "mean"),
                後60日中位數=("Fwd60", "median"),
                後120日平均=("Fwd120", "mean"),
                後120日中位數=("Fwd120", "median"),
            ).reindex(ordered_bands)

            st.dataframe(
                grouped.style.format({
                    "樣本數": "{:,.0f}",
                    "FNG平均": "{:.1f}",
                    "後20日平均": "{:.2%}",
                    "後20日中位數": "{:.2%}",
                    "後60日平均": "{:.2%}",
                    "後60日中位數": "{:.2%}",
                    "後120日平均": "{:.2%}",
                    "後120日中位數": "{:.2%}",
                }),
                use_container_width=True
            )

            st.caption(
                "時間對齊已避免前視偏誤：台灣每個交易日只使用嚴格早於"
                "該日期的最新美國 Fear & Greed。此表是相關性研究，不代表因果。"
            )

            st.markdown("### 資料來源分段比較")
            source_group = study.dropna(subset=["FearGreed"]).groupby("Source").agg(
                樣本數=("FearGreed", "size"),
                FNG平均=("FearGreed", "mean"),
                後20日平均=("Fwd20", "mean"),
                後60日平均=("Fwd60", "mean"),
                後120日平均=("Fwd120", "mean"),
            )
            st.dataframe(
                source_group.style.format({
                    "樣本數": "{:,.0f}",
                    "FNG平均": "{:.1f}",
                    "後20日平均": "{:.2%}",
                    "後60日平均": "{:.2%}",
                    "後120日平均": "{:.2%}",
                }),
                use_container_width=True
            )


# ----------------------------
# ⑦ 資料管理
# ----------------------------
with tabs[6]:
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
    st.markdown("### CNN Fear & Greed 資料庫")

    if fng_db.empty:
        st.warning("目前 fear_greed_daily.csv 尚未建立資料。")
    else:
        fs = fng_db["Date"].min()
        fe = fng_db["Date"].max()
        st.success(
            f"Fear & Greed：{fs.date()} ～ {fe.date()}，"
            f"共 {len(fng_db):,} 筆。"
        )
        st.dataframe(
            fng_db.groupby("Source").agg(
                起始日=("Date", "min"),
                最後日=("Date", "max"),
                筆數=("Date", "size"),
            ),
            use_container_width=True
        )

    if st.button("建立／更新 Fear & Greed", type="secondary"):
        try:
            with st.spinner("下載 Fear & Greed canonical dataset…"):
                fresh_fng = download_fng_combined()
                save_local_fng(fresh_fng)
            st.success(
                f"Fear & Greed 已更新：{fresh_fng['Date'].min().date()} ～ "
                f"{fresh_fng['Date'].max().date()}，共 {len(fresh_fng):,} 筆。"
            )
            st.rerun()
        except Exception as e:
            st.error(f"Fear & Greed 更新失敗：{e}")

    fng_upload = st.file_uploader(
        "匯入 fear_greed_daily.csv",
        type=["csv"],
        key="fng_upload"
    )
    if fng_upload is not None:
        try:
            incoming_fng = pd.read_csv(fng_upload)
            if st.button("確認匯入 Fear & Greed"):
                incoming_fng = incoming_fng.rename(columns={
                    "Fear Greed": "FearGreed",
                    "date": "Date",
                    "rating": "Rating",
                    "source": "Source",
                })
                save_local_fng(incoming_fng)
                st.success("Fear & Greed 已匯入。")
                st.rerun()
        except Exception as e:
            st.error(f"Fear & Greed 匯入失敗：{e}")

    if not fng_db.empty:
        fng_export = fng_db.copy()
        fng_export["Date"] = pd.to_datetime(
            fng_export["Date"]
        ).dt.strftime("%Y-%m-%d")
        fng_bytes = fng_export.to_csv(
            index=False
        ).encode("utf-8-sig")
        st.download_button(
            "下載最新版 fear_greed_daily.csv",
            data=fng_bytes,
            file_name="fear_greed_daily.csv",
            mime="text/csv",
        )
        st.info(
            "和 market_daily.csv 一樣，網站更新後請把 fear_greed_daily.csv "
            "覆蓋回 GitHub 的 data/ 目錄，才能永久保存。"
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
    "v1.8：新增1999年至今長期模擬回測與可調投資開始日期；越跌越買與反彈分批減碼功能仍保留。"
    "歷史回測與模型最佳化均不代表未來報酬。"
)
