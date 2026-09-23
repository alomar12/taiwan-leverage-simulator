
import math
from datetime import date, timedelta

import numpy as np
import pandas as pd
import streamlit as st
import yfinance as yf

st.set_page_config(page_title="台灣50正2再平衡模擬器", layout="wide")

# -----------------------------
# 資料取得
# -----------------------------
@st.cache_data(ttl=3600, show_spinner=False)
def download_yahoo(ticker: str, start: str, end: str) -> pd.DataFrame:
    """下載 Yahoo Finance 日資料。end 為含當日，因此送給 yfinance 時加一天。"""
    end_dt = pd.Timestamp(end) + pd.Timedelta(days=1)
    df = yf.download(
        ticker,
        start=start,
        end=end_dt.strftime("%Y-%m-%d"),
        auto_adjust=False,
        progress=False,
        actions=False,
        threads=False,
    )
    if df is None or df.empty:
        raise RuntimeError(f"{ticker} 無法取得資料")

    def col(name):
        if isinstance(df.columns, pd.MultiIndex):
            if name not in df.columns.get_level_values(0):
                return None
            x = df[name]
            if isinstance(x, pd.DataFrame):
                x = x.iloc[:, 0]
            return x
        if name not in df.columns:
            return None
        return df[name]

    close = col("Close")
    adj = col("Adj Close")
    if close is None:
        raise RuntimeError(f"{ticker} 缺少 Close 欄位")
    if adj is None:
        adj = close.copy()

    ans = pd.DataFrame({"Close": close.astype(float), "AdjClose": adj.astype(float)})
    ans.index = pd.to_datetime(ans.index).tz_localize(None)
    ans = ans[~ans.index.duplicated(keep="last")].dropna(how="all").sort_index()
    return ans


@st.cache_data(ttl=3600, show_spinner=False)
def load_market_data(end: str):
    # 0050：策略基準與 2008 壓力測試；00631L：2014/10/31 後實際正2；
    # ^TWII：2000 網路泡沫壓力測試代理。
    d0050 = download_yahoo("0050.TW", "2003-06-30", end)
    d2x = download_yahoo("00631L.TW", "2014-10-31", end)
    twii = download_yahoo("^TWII", "1999-01-01", end)
    return d0050, d2x, twii


# -----------------------------
# 正2偏差模型
# -----------------------------
def build_features(r: pd.Series) -> pd.DataFrame:
    x = pd.DataFrame(index=r.index)
    x["r"] = r
    x["abs_r"] = r.abs()
    x["r2"] = r.pow(2)
    x["vol20"] = r.rolling(20, min_periods=10).std() * np.sqrt(252)
    x["mom20"] = (1 + r).rolling(20, min_periods=10).apply(np.prod, raw=True) - 1
    x = x.replace([np.inf, -np.inf], np.nan)
    return x


def fit_gap_model(under_ret: pd.Series, lev_ret: pd.Series):
    z = pd.concat([under_ret.rename("u"), lev_ret.rename("l")], axis=1).dropna()
    f = build_features(z["u"]).dropna()
    z = z.loc[f.index]
    # 實際正2 - 理論每日2倍
    y = z["l"] - 2.0 * z["u"]
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

    # 以校準期波動分位數建立 regime bucket
    q = cal["vol20"].quantile([0.25, 0.50, 0.75]).values
    cal["vol_bucket"] = np.digitize(cal["vol20"].values, q, right=True)
    return coef, cal, q


def predict_gap(r: pd.Series, coef: np.ndarray) -> pd.Series:
    f = build_features(r).copy()
    # 前幾日沒有20日資料時，以可得資料的中位數填補
    for c in ["vol20", "mom20"]:
        f[c] = f[c].fillna(f[c].median())
    X = np.column_stack([
        np.ones(len(f)),
        f["r"].fillna(0).values,
        f["abs_r"].fillna(0).values,
        f["r2"].fillna(0).values,
        f["vol20"].fillna(0).values,
        f["mom20"].fillna(0).values,
    ])
    return pd.Series(X @ coef, index=r.index, name="pred_gap")


def synthetic_2x_returns(
    underlying_ret: pd.Series,
    coef: np.ndarray,
    calibration: pd.DataFrame,
    vol_q: np.ndarray,
    mode: str = "校準平均",
    seed: int = 42,
) -> pd.Series:
    base = 2.0 * underlying_ret.fillna(0)
    gap = predict_gap(underlying_ret.fillna(0), coef)
    out = base + gap

    if mode == "波動狀態抽樣":
        rng = np.random.default_rng(seed)
        f = build_features(underlying_ret.fillna(0))
        f["vol20"] = f["vol20"].fillna(calibration["vol20"].median())
        sign = np.where(underlying_ret.fillna(0).values >= 0, 1, -1)
        bucket = np.digitize(f["vol20"].values, vol_q, right=True)
        sampled = np.zeros(len(f))
        all_eps = calibration["eps"].dropna().values

        # 同方向 + 同波動區間抽取殘差；樣本太少就退回整體殘差
        for i, (s, b) in enumerate(zip(sign, bucket)):
            pool = calibration.loc[
                (calibration["sign"] == s) & (calibration["vol_bucket"] == b), "eps"
            ].dropna().values
            if len(pool) < 20:
                pool = all_eps
            sampled[i] = rng.choice(pool)
        out = out + sampled

    elif mode == "理論2倍":
        out = base

    # 防止數值上出現基金單日跌逾100%的不合理情形
    return out.clip(lower=-0.95, upper=1.50)


def make_level(ret: pd.Series, start=100.0) -> pd.Series:
    return start * (1 + ret.fillna(0)).cumprod()


# -----------------------------
# 投資策略
# -----------------------------
def rebalance_to_target(lev, cash, target, buy_cost, sell_cost):
    total = lev + cash
    desired = total * target

    if desired > lev:
        # L + x = w * (T - buy_cost*x)
        x = max(0.0, (target * total - lev) / (1.0 + target * buy_cost))
        x = min(x, cash / (1.0 + buy_cost))
        cost = x * buy_cost
        lev += x
        cash -= x + cost
        turnover = x
    else:
        # L - x = w * (T - sell_cost*x)
        denom = max(1e-12, 1.0 - target * sell_cost)
        x = max(0.0, (lev - target * total) / denom)
        x = min(x, lev)
        cost = x * sell_cost
        lev -= x
        cash += x - cost
        turnover = x
    return lev, cash, turnover


def simulate_rebalance(
    lev_ret: pd.Series,
    lev_level: pd.Series,
    capital: float,
    target: float,
    trigger_mode: str,
    down_trigger: float,
    up_trigger: float,
    band: float,
    cash_yield: float,
    buy_cost: float,
    sell_cost: float,
    min_days: int = 0,
):
    idx = lev_ret.index.intersection(lev_level.index)
    r = lev_ret.reindex(idx).fillna(0)
    level = lev_level.reindex(idx).ffill()

    lev = capital * target
    cash = capital * (1 - target)
    anchor = float(level.iloc[0])
    last_rebal_i = -10**9
    rows = []
    n_rebal = 0
    turnover_sum = 0.0

    daily_cash = (1 + cash_yield) ** (1 / 252) - 1 if cash_yield > -1 else -1

    for i, dt in enumerate(idx):
        lev *= (1 + float(r.loc[dt]))
        cash *= (1 + daily_cash)
        total = max(lev + cash, 1e-12)
        weight = lev / total
        rel_price = float(level.loc[dt]) / anchor - 1.0 if anchor > 0 else 0.0

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
            before = lev + cash
            lev, cash, turnover = rebalance_to_target(
                lev, cash, target, buy_cost, sell_cost
            )
            turnover_sum += turnover
            n_rebal += 1
            last_rebal_i = i
            anchor = float(level.loc[dt])
            total = lev + cash
            weight = lev / total if total > 0 else 0.0

        rows.append((dt, lev + cash, lev, cash, weight, reason))

    df = pd.DataFrame(
        rows, columns=["Date", "Portfolio", "Leveraged", "Cash", "LevWeight", "Reason"]
    ).set_index("Date")
    df["Peak"] = df["Portfolio"].cummax()
    df["Drawdown"] = df["Portfolio"] / df["Peak"] - 1
    return df, n_rebal, turnover_sum


def simulate_benchmark(ret: pd.Series, capital: float):
    s = capital * (1 + ret.fillna(0)).cumprod()
    df = pd.DataFrame({"Portfolio": s})
    df["Peak"] = df["Portfolio"].cummax()
    df["Drawdown"] = df["Portfolio"] / df["Peak"] - 1
    return df


def metrics(eq: pd.Series, n_rebal=0, turnover=0.0):
    eq = eq.dropna()
    if len(eq) < 2:
        return {}
    years = max((eq.index[-1] - eq.index[0]).days / 365.25, 1 / 252)
    total_ret = eq.iloc[-1] / eq.iloc[0] - 1
    cagr = (eq.iloc[-1] / eq.iloc[0]) ** (1 / years) - 1
    dr = eq.pct_change().dropna()
    vol = dr.std() * np.sqrt(252) if len(dr) else np.nan
    dd = eq / eq.cummax() - 1
    mdd = dd.min()
    sharpe = (dr.mean() / dr.std() * np.sqrt(252)) if len(dr) and dr.std() > 0 else np.nan
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
    # maximize CAGR, maximize MaxDD (less negative)
    x = df[["CAGR", "最大回撤"]].values
    keep = np.ones(len(df), dtype=bool)
    for i in range(len(df)):
        if not keep[i]:
            continue
        dominated = (
            (x[:, 0] >= x[i, 0]) &
            (x[:, 1] >= x[i, 1]) &
            ((x[:, 0] > x[i, 0]) | (x[:, 1] > x[i, 1]))
        )
        if dominated.any():
            keep[i] = False
    return df.loc[keep].copy()


# -----------------------------
# UI
# -----------------------------
st.title("台灣50正2 × 現金：再平衡與股災壓力測試模擬器")
st.caption(
    "核心用途：用00631L自2014/10/31以來的實際偏差校準正2，"
    "再將偏差模型套到2000、2008等歷史行情，測試手動再平衡參數的穩健性。"
)

with st.sidebar:
    st.header("共通設定")
    end_date = st.date_input("資料截止日", value=date.today()).strftime("%Y-%m-%d")
    capital = st.number_input("初始資金（元）", min_value=10000, value=1_000_000, step=100_000)
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

try:
    with st.spinner("下載0050、00631L與台灣加權指數資料…"):
        d0050, d2x, twii = load_market_data(end_date)
except Exception as e:
    st.error(f"資料下載失敗：{e}")
    st.info("Yahoo Finance 偶爾會限制連線。稍後重整即可；也可在程式內改用自行上傳CSV。")
    st.stop()

# 實際比較期
start_actual = max(pd.Timestamp("2014-10-31"), d2x.index.min(), d0050.index.min())
common = d0050.loc[start_actual:].index.intersection(d2x.loc[start_actual:].index)
p0050 = d0050["Close"].reindex(common).dropna()
tr0050 = d0050["AdjClose"].reindex(common).dropna()
p2x = d2x["Close"].reindex(common).dropna()
common = p0050.index.intersection(p2x.index).intersection(tr0050.index)
p0050, tr0050, p2x = p0050.reindex(common), tr0050.reindex(common), p2x.reindex(common)

r0050_price = p0050.pct_change().fillna(0)
r0050_total = tr0050.pct_change().fillna(0)
r2x_actual = p2x.pct_change().fillna(0)

coef, calibration, vol_q = fit_gap_model(r0050_price, r2x_actual)

tabs = st.tabs(["① 實際資料回測", "② 參數最佳化", "③ 2000/2008壓力測試", "④ 正2偏差模型"])

with tabs[0]:
    st.subheader("2014/10/31至今：實際00631L vs 100%0050")
    lev_level_actual = make_level(r2x_actual)

    strat, nreb, turnover = simulate_rebalance(
        r2x_actual, lev_level_actual, capital, target, trigger_mode,
        down_trigger, up_trigger, band, cash_yield, buy_cost, sell_cost, min_days
    )
    bench = simulate_benchmark(r0050_total.reindex(strat.index).fillna(0), capital)

    m_s = metrics(strat["Portfolio"], nreb, turnover)
    m_b = metrics(bench["Portfolio"])

    summary = pd.DataFrame([m_b, m_s], index=["100% 0050", f"{target_pct}% 正2＋現金"])
    fmt = summary.copy()
    st.dataframe(
        fmt.style.format({
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

    dd = pd.concat([
        bench["Drawdown"].rename("100% 0050"),
        strat["Drawdown"].rename("正2＋現金"),
    ], axis=1)
    st.caption("回撤")
    st.line_chart(dd)

    trades = strat.loc[strat["Reason"] != "", ["Portfolio", "LevWeight", "Reason"]]
    with st.expander(f"再平衡紀錄（{len(trades)}次）"):
        st.dataframe(trades, use_container_width=True)

with tabs[1]:
    st.subheader("參數掃描：找『穩健區域』，不是只找單一最高報酬點")
    st.write(
        "建議先以 Calmar（年化報酬 ÷ 最大回撤）與 Pareto frontier 觀察，"
        "再檢查2000、2008壓力測試。單看最高CAGR很容易過度擬合。"
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        t_min, t_max = st.slider("正2比例搜尋範圍（%）", 10, 80, (30, 70), 5)
    with c2:
        d_min, d_max = st.slider("下跌門檻搜尋範圍（%）", 5, 60, (15, 45), 5)
    with c3:
        u_min, u_max = st.slider("上漲門檻搜尋範圍（%）", 5, 100, (20, 60), 5)

    dd_limit = st.slider("可接受最大回撤下限（至少不能差於）", -70, -10, -35, 1) / 100

    if st.button("開始參數掃描", type="primary"):
        rows = []
        targets = np.arange(t_min, t_max + 0.1, 5) / 100
        downs = np.arange(d_min, d_max + 0.1, 5) / 100
        ups = np.arange(u_min, u_max + 0.1, 5) / 100

        prog = st.progress(0)
        total_n = len(targets) * len(downs) * len(ups)
        k = 0
        for t in targets:
            for d in downs:
                for u in ups:
                    s, nr, to = simulate_rebalance(
                        r2x_actual, lev_level_actual, capital, t, "價格漲跌門檻",
                        d, u, 0.075, cash_yield, buy_cost, sell_cost, min_days
                    )
                    m = metrics(s["Portfolio"], nr, to)
                    rows.append({
                        "正2比例": t,
                        "下跌門檻": d,
                        "上漲門檻": u,
                        **m,
                    })
                    k += 1
                    if k % max(1, total_n // 100) == 0:
                        prog.progress(min(k / total_n, 1.0))
        prog.empty()

        res = pd.DataFrame(rows)
        res = res.sort_values(["Calmar", "CAGR"], ascending=False)
        feasible = res[res["最大回撤"] >= dd_limit].copy()
        front = pareto_frontier(res)

        st.session_state["grid_result"] = res
        st.session_state["grid_feasible"] = feasible
        st.session_state["grid_front"] = front

    if "grid_result" in st.session_state:
        res = st.session_state["grid_result"]
        feasible = st.session_state["grid_feasible"]
        front = st.session_state["grid_front"]

        st.markdown("**Calmar前20名**")
        top = res.head(20).copy()
        st.dataframe(
            top.style.format({
                "正2比例": "{:.0%}", "下跌門檻": "{:.0%}", "上漲門檻": "{:.0%}",
                "期末資產": "{:,.0f}", "累積報酬": "{:.1%}", "CAGR": "{:.2%}",
                "年化波動": "{:.2%}", "最大回撤": "{:.2%}",
                "Sharpe(無風險=0)": "{:.2f}", "Calmar": "{:.2f}",
                "再平衡次數": "{:.0f}", "累計換手金額": "{:,.0f}",
            }),
            use_container_width=True
        )

        if len(feasible):
            st.markdown(f"**符合最大回撤 ≥ {dd_limit:.0%} 的最高CAGR前20名**")
            x = feasible.sort_values("CAGR", ascending=False).head(20)
            st.dataframe(
                x.style.format({
                    "正2比例": "{:.0%}", "下跌門檻": "{:.0%}", "上漲門檻": "{:.0%}",
                    "期末資產": "{:,.0f}", "累積報酬": "{:.1%}", "CAGR": "{:.2%}",
                    "年化波動": "{:.2%}", "最大回撤": "{:.2%}",
                    "Sharpe(無風險=0)": "{:.2f}", "Calmar": "{:.2f}",
                }),
                use_container_width=True
            )
        else:
            st.warning("目前搜尋範圍沒有任何參數符合你設定的最大回撤限制。")

        st.markdown("**報酬—回撤 Pareto frontier**")
        st.caption("這些參數沒有被另一組參數同時在CAGR與最大回撤兩方面完全超越。")
        st.dataframe(
            front.sort_values("最大回撤", ascending=False).style.format({
                "正2比例": "{:.0%}", "下跌門檻": "{:.0%}", "上漲門檻": "{:.0%}",
                "CAGR": "{:.2%}", "最大回撤": "{:.2%}", "Calmar": "{:.2f}",
                "期末資產": "{:,.0f}",
            }),
            use_container_width=True
        )

with tabs[2]:
    st.subheader("把2014年至今的『實際偏差』套回歷史股災")
    stress_name = st.selectbox(
        "壓力期間",
        [
            "2000網路泡沫（TAIEX代理）",
            "2008金融海嘯（0050代理）",
            "自訂TAIEX期間",
        ]
    )
    model_mode = st.radio(
        "正2歷史重建方式",
        ["校準平均", "波動狀態抽樣", "理論2倍"],
        horizontal=True
    )

    if stress_name == "2000網路泡沫（TAIEX代理）":
        sdt, edt = pd.Timestamp("2000-01-01"), pd.Timestamp("2002-12-31")
        underlying_price = twii["Close"].loc[sdt:edt]
        benchmark_price = twii["Close"].loc[sdt:edt]
        note = "2000年尚無0050與臺灣50正2，因此此處以TAIEX作為市場代理；存在成分差異。"
    elif stress_name == "2008金融海嘯（0050代理）":
        sdt, edt = pd.Timestamp("2007-01-01"), pd.Timestamp("2009-12-31")
        underlying_price = d0050["Close"].loc[sdt:edt]
        benchmark_price = d0050["AdjClose"].loc[sdt:edt]
        note = "正2尚未上市；以0050價格報酬生成模擬正2，benchmark則使用0050還原報酬。"
    else:
        c1, c2 = st.columns(2)
        with c1:
            custom_start = st.date_input("自訂開始", value=date(1999, 1, 5), key="cs")
        with c2:
            custom_end = st.date_input("自訂結束", value=date(2014, 10, 30), key="ce")
        sdt, edt = pd.Timestamp(custom_start), pd.Timestamp(custom_end)
        underlying_price = twii["Close"].loc[sdt:edt]
        benchmark_price = twii["Close"].loc[sdt:edt]
        note = "自訂期間以TAIEX作為市場代理。"

    st.info(note)
    ur = underlying_price.pct_change().dropna()
    br = benchmark_price.pct_change().reindex(ur.index).fillna(0)

    if model_mode == "波動狀態抽樣":
        n_mc = st.slider("蒙地卡羅路徑數", 50, 1000, 300, 50)
        seed0 = st.number_input("亂數種子", value=42, step=1)

        vals, mdds, cagrs = [], [], []
        for j in range(n_mc):
            sr = synthetic_2x_returns(ur, coef, calibration, vol_q, "波動狀態抽樣", int(seed0+j))
            lvl = make_level(sr)
            sim, nr, to = simulate_rebalance(
                sr, lvl, capital, target, trigger_mode, down_trigger, up_trigger,
                band, cash_yield, buy_cost, sell_cost, min_days
            )
            mm = metrics(sim["Portfolio"], nr, to)
            vals.append(mm["期末資產"])
            mdds.append(mm["最大回撤"])
            cagrs.append(mm["CAGR"])

        dist = pd.DataFrame({"期末資產": vals, "最大回撤": mdds, "CAGR": cagrs})
        qs = dist.quantile([0.05, 0.50, 0.95])
        st.markdown("**蒙地卡羅結果（5% / 中位數 / 95%）**")
        st.dataframe(
            qs.style.format({
                "期末資產": "{:,.0f}",
                "最大回撤": "{:.2%}",
                "CAGR": "{:.2%}",
            }),
            use_container_width=True
        )
    else:
        sr = synthetic_2x_returns(ur, coef, calibration, vol_q, model_mode, 42)
        lvl = make_level(sr)
        sim, nr, to = simulate_rebalance(
            sr, lvl, capital, target, trigger_mode, down_trigger, up_trigger,
            band, cash_yield, buy_cost, sell_cost, min_days
        )
        bm = simulate_benchmark(br, capital)
        out = pd.DataFrame([
            metrics(bm["Portfolio"]),
            metrics(sim["Portfolio"], nr, to)
        ], index=["100%市場/0050代理", "正2＋現金"])
        st.dataframe(
            out.style.format({
                "期末資產": "{:,.0f}", "累積報酬": "{:.1%}", "CAGR": "{:.2%}",
                "年化波動": "{:.2%}", "最大回撤": "{:.2%}",
                "Sharpe(無風險=0)": "{:.2f}", "Calmar": "{:.2f}",
                "再平衡次數": "{:.0f}", "累計換手金額": "{:,.0f}",
            }),
            use_container_width=True
        )
        curves = pd.concat([
            bm["Portfolio"].rename("100%市場/0050代理"),
            sim["Portfolio"].rename("模擬正2＋現金"),
        ], axis=1)
        st.line_chart(curves)

with tabs[3]:
    st.subheader("00631L實際日報酬與『0050日報酬×2』的差距")
    gap = calibration["gap"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("平均日偏差", f"{gap.mean():.4%}")
    c2.metric("偏差日標準差", f"{gap.std():.4%}")
    c3.metric("5%分位", f"{gap.quantile(0.05):.4%}")
    c4.metric("95%分位", f"{gap.quantile(0.95):.4%}")

    st.write("模型使用：日報酬、絕對日報酬、平方報酬、20日波動率、20日動能，估計實際正2相對理論2倍的條件偏差。")
    coef_table = pd.DataFrame({
        "變數": ["常數", "日報酬", "|日報酬|", "日報酬²", "20日年化波動", "20日動能"],
        "係數": coef
    })
    st.dataframe(coef_table, use_container_width=True)

    gap_df = pd.DataFrame({
        "實際偏差": calibration["gap"],
        "模型預測偏差": calibration["pred_gap"],
    })
    st.line_chart(gap_df)

st.divider()
st.caption(
    "重要限制：00631L追蹤臺灣50指數單日正向2倍；2000年壓力測試使用TAIEX代理，"
    "不是臺灣50的真實歷史。歷史校準與參數最佳化不能保證未來結果。"
)
