
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import rcParams

try:
    rcParams["font.family"] = "Noto Sans CJK JP"
except Exception:
    pass

REQUIRED_COLUMNS_MIN = ["開始日時", "使用電力量(ロス後)", "使用電力量(ロス前)"]
OPTIONAL_COLUMNS = ["JEPXスポットプライス"]

def load_excel_to_df(file, sheet_name=None):
    if sheet_name is None or len(str(sheet_name).strip()) == 0:
        xls = pd.ExcelFile(file)
        sheet_name = xls.sheet_names[0]
    df = pd.read_excel(file, sheet_name=sheet_name)
    if "終了日時" in df.columns:
        df = df[df["終了日時"].notna()].copy()
    for c in REQUIRED_COLUMNS_MIN:
        if c not in df.columns:
            raise ValueError(f"必須列が見つかりません: {c}")
    for c in OPTIONAL_COLUMNS:
        if c not in df.columns:
            df[c] = pd.NA
    df["開始日時"] = pd.to_datetime(df["開始日時"], errors="coerce")
    df = df.dropna(subset=["開始日時"]).copy()
    df = df.set_index("開始日時").sort_index()
    # kW列
    df["使用電力量(ロス後)_kW"] = df["使用電力量(ロス後)"] / 0.5
    df["使用電力量(ロス前)_kW"] = df["使用電力量(ロス前)"] / 0.5
    return df

def select_range(df, start=None, end=None):
    tz = df.index.tz
    s = pd.to_datetime(start) if start else None
    e = pd.to_datetime(end) if end else None
    if tz is not None:
        if s is not None and s.tzinfo is None:
            s = s.tz_localize(tz)
        if e is not None and e.tzinfo is None:
            e = e.tz_localize(tz)
    if s is not None:
        df = df.loc[df.index >= s]
    if e is not None:
        df = df.loc[df.index <= e]
    return df

def series_picker(df, series="both", use_kw=True):
    if use_kw:
        cols = {"ロス後": "使用電力量(ロス後)_kW", "ロス前": "使用電力量(ロス前)_kW"}
    else:
        cols = {"ロス後": "使用電力量(ロス後)", "ロス前": "使用電力量(ロス前)"}
    if series == "ロス後":
        out = df[[cols["ロス後"]]].rename(columns={cols["ロス後"]: "ロス後"})
    elif series == "ロス前":
        out = df[[cols["ロス前"]]].rename(columns={cols["ロス前"]: "ロス前"})
    else:
        out = df[[cols["ロス後"], cols["ロス前"]]].rename(
            columns={cols["ロス後"]: "ロス後", cols["ロス前"]: "ロス前"}
        )
    return out

def aggregate_df(df, aggregate=None, how="mean"):
    if aggregate is None:
        return df
    if aggregate == "D":
        return getattr(df.resample("D"), how)()
    if aggregate == "M":
        return getattr(df.resample("MS"), how)()
    raise ValueError("aggregate には None / 'D' / 'M' を指定してください。")

def list_dates(df):
    uniq = pd.to_datetime(df.index.date).unique()
    catalog = pd.DataFrame({"date": pd.to_datetime(uniq)})
    catalog["year"] = catalog["date"].dt.year
    catalog["month"] = catalog["date"].dt.month
    catalog["day"] = catalog["date"].dt.day
    catalog["month_label"] = catalog["date"].dt.strftime("%Y-%m")
    return catalog.sort_values("date")

def get_day_slice(df, date_val):
    tz = df.index.tz
    start = pd.Timestamp(date_val)
    if tz is not None and start.tzinfo is None:
        start = start.tz_localize(tz)
    end = start + pd.Timedelta(days=1)
    return df.loc[(df.index >= start) & (df.index < end)].copy()

def overlay_by_dates(df, dates, which="ロス後"):
    mat = pd.DataFrame(index=range(48))
    for d in dates:
        day = get_day_slice(df, d)
        if day.empty:
            continue
        idx_local = day.index.tz_convert("Asia/Tokyo") if day.index.tz is not None else day.index
        slots = ((idx_local - idx_local.normalize()) / pd.Timedelta(minutes=30)).astype(int)
        ser = day[f"使用電力量({which})_kW"]
        ser = pd.Series(ser.values, index=slots).reindex(range(48))
        mat[str(pd.to_datetime(d).date())] = ser.values
    return mat

def overlay_by_dates_price(df, dates):
    mat = pd.DataFrame(index=range(48))
    for d in dates:
        day = get_day_slice(df, d)
        if day.empty or "JEPXスポットプライス" not in day.columns:
            continue
        idx_local = day.index.tz_convert("Asia/Tokyo") if day.index.tz is not None else day.index
        slots = ((idx_local - idx_local.normalize()) / pd.Timedelta(minutes=30)).astype(int)
        ser = day["JEPXスポットプライス"]
        ser = pd.Series(ser.values, index=slots).reindex(range(48))
        mat[str(pd.to_datetime(d).date())] = ser.values
    return mat

def overlay_price_full_year(df):
    """1年分の各日（JEPX価格）を0..47スロットに並べた行列"""
    mat = pd.DataFrame(index=range(48))
    if df.index.tz is not None:
        uniq_days = pd.to_datetime(df.index.tz_convert("Asia/Tokyo").date).unique()
    else:
        uniq_days = pd.to_datetime(df.index.date).unique()
    for d in uniq_days:
        day = get_day_slice(df, d)
        if day.empty or "JEPXスポットプライス" not in day.columns:
            continue
        idx_local = day.index.tz_convert("Asia/Tokyo") if day.index.tz is not None else day.index
        slots = ((idx_local - idx_local.normalize()) / pd.Timedelta(minutes=30)).astype(int)
        ser = day["JEPXスポットプライス"]
        ser = pd.Series(ser.values, index=slots).reindex(range(48))
        mat[str(pd.to_datetime(d).date())] = ser.values
    return mat

def pick_load_series(df, preferred=None):
    if preferred and preferred in df.columns:
        return df[preferred].astype(float)
    for c in ["需要計画量(ロス前)", "需要計画量", "需要kW"]:
        if c in df.columns and df[c].notna().any():
            return df[c].astype(float)
    return (df["使用電力量(ロス後)"].astype(float) / 0.5)

def pick_generation_series(df, preferred=None):
    if preferred and preferred in df.columns:
        return df[preferred].astype(float)
    for c in ["自家発出力", "PV出力", "太陽光出力", "発電kW"]:
        if c in df.columns and df[c].notna().any():
            return df[c].astype(float)
    return pd.Series(0.0, index=df.index)

def compute_export_offer_def1(df, P_pcs=1000.0, P_exp_max=None, load_col=None, gen_col=None):
    L = pick_load_series(df, preferred=load_col)
    G = pick_generation_series(df, preferred=gen_col)
    offer = (P_pcs - (L - G)).clip(lower=0)
    if P_exp_max is not None:
        offer = offer.clip(upper=float(P_exp_max))
    return offer, L, G

def simulate_soc_with_charge_periodic_reset(
    df, P_pcs=1000.0, P_chg=1000.0, E_nom=2000.0,
    start=None, end=None,
    soc_init_pct=90.0, soc_floor_pct=10.0, reset_every_days=4,
    load_col=None, gen_col=None
):
    # 期間トリム
    if start is not None:
        start = pd.Timestamp(start)
        if df.index.tz is not None and start.tzinfo is None:
            start = start.tz_localize(df.index.tz)
        df = df.loc[df.index >= start]
    if end is not None:
        end = pd.Timestamp(end)
        if df.index.tz is not None and end.tzinfo is None:
            end = end.tz_localize(df.index.tz)
        df = df.loc[df.index <= end]

    L = pick_load_series(df, preferred=load_col)
    G = pick_generation_series(df, preferred=gen_col)
    net_load = (L - G).clip(lower=0.0)
    supply_kW = net_load.clip(upper=P_pcs)
    use_kWh = (supply_kW * 0.5).fillna(0.0)

    E_init = float(soc_init_pct) / 100.0 * E_nom
    E_floor = float(soc_floor_pct) / 100.0 * E_nom

    times = use_kWh.index
    if len(times) == 0:
        return pd.DataFrame(columns=["SOC_kWh", "SOC_%", "charging"])

    start_day = times[0].normalize()

    E = E_init
    charging_mode = False
    charge_deficit = 0.0

    soc_kWh = []
    soc_pct = []
    charging_flags = []

    for t, e_use in use_kWh.items():
        day_num = int((t.normalize() - start_day) / pd.Timedelta(days=1))
        if (t.hour == 0 and t.minute == 0) and (day_num % int(reset_every_days) == 0):
            charge_deficit = max(0.0, E_init - E)
            charging_mode = charge_deficit > 0.0

        if charging_mode:
            add_kWh = min(float(P_chg) * 0.5, charge_deficit)
            E = min(E_init, E + add_kWh)
            charge_deficit -= add_kWh
            if charge_deficit <= 1e-9 or E >= E_init - 1e-9:
                charging_mode = False
            charging_flags.append(True)
        else:
            E = max(E_floor, E - float(e_use))
            charging_flags.append(False)

        soc_kWh.append(E)
        soc_pct.append(100.0 * E / E_nom)

    return pd.DataFrame({"SOC_kWh": soc_kWh, "SOC_%": soc_pct, "charging": charging_flags}, index=times)

def plot_lines(x, y_dict, xlabel, ylabel, title):
    plt.figure(figsize=(12, 6))
    for label, series in y_dict.items():
        plt.plot(x, series, label=label)
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.title(title)
    plt.legend()
    plt.grid(True)
    return plt.gcf()

def derive_charge_cost_series(soc_df, df_price, price_col="JEPXスポットプライス"):
    if price_col not in df_price.columns:
        price_series = pd.Series(0.0, index=soc_df.index)
    else:
        price_series = df_price[price_col].astype(float).reindex(soc_df.index)
    soc_e = soc_df["SOC_kWh"]
    delta = soc_e.diff().fillna(0.0)
    charge_kWh = pd.Series(0.0, index=soc_df.index)
    charge_kWh[soc_df["charging"]] = delta.clip(lower=0.0)[soc_df["charging"]]
    cost = charge_kWh * price_series
    cum_cost = cost.cumsum()
    return charge_kWh, price_series, cost, cum_cost


def simulate_soc_concurrent_price_optimized(
    df, P_pcs=1000.0, P_chg=1000.0, E_nom=2000.0,
    start=None, end=None,
    soc_init_pct=90.0, soc_floor_pct=10.0, reset_every_days=4,
    price_col="JEPXスポットプライス",
    load_col=None, gen_col=None
):
    """
    充電日には「その日の最安コマから」充電量を割当。
    充電中も負荷対応を継続し、(供出kW + 充電kW) <= PCS定格 を満たす。
    充電は初期SOC(=目標)まで。到達不能な場合はその日最大限充電して翌日に繰越。
    """
    # 期間トリム
    if start is not None:
        start = pd.Timestamp(start)
        if df.index.tz is not None and start.tzinfo is None:
            start = start.tz_localize(df.index.tz)
        df = df.loc[df.index >= start]
    if end is not None:
        end = pd.Timestamp(end)
        if df.index.tz is not None and end.tzinfo is None:
            end = end.tz_localize(df.index.tz)
        df = df.loc[df.index <= end]
    if df.empty:
        return pd.DataFrame(columns=["SOC_kWh", "SOC_%", "charging", "charge_kWh", "supply_kW"])

    # 需要・自家発
    L = pick_load_series(df, preferred=load_col)
    G = pick_generation_series(df, preferred=gen_col)
    net_load = (L - G).clip(lower=0.0)  # kW
    # まず負荷に供出（PCS上限）
    supply_kW = net_load.clip(upper=float(P_pcs))

    # 価格
    if price_col in df.columns:
        price = df[price_col].astype(float)
    else:
        price = pd.Series(0.0, index=df.index)

    # 初期/下限
    E_init = float(soc_init_pct) / 100.0 * E_nom
    E_floor = float(soc_floor_pct) / 100.0 * E_nom

    # スロット長
    slot_h = 0.5

    # 出力用ベクトル
    E = []
    E_curr = E_init
    charging_flag = []
    charge_kWh_vec = []

    times = df.index
    start_day = times[0].normalize()

    # 日単位でスケジューリング
    day_groups = df.groupby(df.index.normalize())

    day_counter = 0
    for day, d in day_groups:
        # その日が充電日か？
        is_charge_day = (day_counter % int(reset_every_days) == 0)
        prices = price.loc[d.index]
        sup_kW = supply_kW.loc[d.index]

        # まず「充電ゼロ」で当日終端までSOCを推定
        E_tmp = E_curr
        soc_floor_clip = []
        E_path_nochg = []
        for kW in sup_kW:
            E_tmp = max(E_floor, E_tmp - kW * slot_h)  # 放電のみ
            E_path_nochg.append(E_tmp)
            soc_floor_clip.append(E_tmp <= E_floor + 1e-9)

        # 目標SOC（初期SOC）まで戻すための必要量R
        R = max(0.0, E_init - E_path_nochg[-1]) if is_charge_day else 0.0

        # 充電可能容量（各コマ）：min(P_chg, P_pcs - supply_kW) * slot_h
        chg_cap = (pd.Series(float(P_pcs), index=d.index) - sup_kW).clip(lower=0.0)
        chg_cap = pd.concat([chg_cap, pd.Series(float(P_chg), index=d.index)], axis=1).min(axis=1) * slot_h  # kWh/slot

        # 価格最小から割当（is_charge_day のときだけ）
        alloc = pd.Series(0.0, index=d.index)
        if is_charge_day and R > 0:
            order = prices.sort_values(kind="mergesort").index  # 安定ソート
            remaining = R
            for t in order:
                cap = chg_cap.loc[t]
                if cap <= 0 or remaining <= 0:
                    continue
                add = min(cap, remaining)
                alloc.loc[t] = add
                remaining -= add
            # 残があれば翌日に繰越（ここでは警告せず、残量はE_currに反映されるだけ）

        # 当日の時間軸でSOCを更新（同時供出を考慮）
        for t in d.index:
            # 放電
            E_curr = max(E_floor, E_curr - sup_kW.loc[t] * slot_h)
            # 充電（割当分+容量制約+SOC上限）
            add = alloc.loc[t] if is_charge_day else 0.0
            # 容量上限
            add = min(add, max(0.0, E_init - E_curr))
            E_curr += add

            E.append(E_curr)
            charging_flag.append(add > 1e-12)
            charge_kWh_vec.append(add)

        day_counter += 1

    out = pd.DataFrame({
        "SOC_kWh": E,
        "SOC_%": [100.0 * e / E_nom for e in E],
        "charging": charging_flag,
        "charge_kWh": charge_kWh_vec,
        "supply_kW": supply_kW.values
    }, index=df.index)

    return out
