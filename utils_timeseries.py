
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import rcParams

# ---- Global font: try Noto Sans CJK JP (fallback to default) ----
try:
    rcParams["font.family"] = "Noto Sans CJK JP"
except Exception:
    pass

REQUIRED_COLUMNS_MIN = ["開始日時", "使用電力量(ロス後)", "使用電力量(ロス前)"]
OPTIONAL_COLUMNS = ["JEPXスポットプライス"]  # 円/kWh

def load_excel_to_df(file, sheet_name=None):
    """
    Load the Excel file into a cleaned dataframe.
    Adds kW columns (kWh / 0.5h) and keeps JEPXスポットプライス if present.
    """
    if sheet_name is None or len(str(sheet_name).strip()) == 0:
        xls = pd.ExcelFile(file)
        sheet_name = xls.sheet_names[0]
    df = pd.read_excel(file, sheet_name=sheet_name)

    # Filter out non-time-series rows where '終了日時' is NaN (if exists)
    if "終了日時" in df.columns:
        df = df[df["終了日時"].notna()].copy()

    # Ensure required columns
    for c in REQUIRED_COLUMNS_MIN:
        if c not in df.columns:
            raise ValueError(f"必須列が見つかりません: {c}")

    # Keep optional columns even if missing
    for c in OPTIONAL_COLUMNS:
        if c not in df.columns:
            df[c] = pd.NA

    # Parse datetime and index
    df["開始日時"] = pd.to_datetime(df["開始日時"], errors="coerce")
    df = df.dropna(subset=["開始日時"]).copy()
    df = df.set_index("開始日時").sort_index()

    # kW columns
    df["使用電力量(ロス後)_kW"] = df["使用電力量(ロス後)"] / 0.5  # 30min -> avg kW
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
    """Overlay matrix for JEPXスポットプライス（円/kWh）"""
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

# ---------- 供出可能量（定義①）関連 ----------
def pick_load_series(df, preferred=None):
    """
    優先列を指定できる。なければ需要計画量(ロス前)→使用電力量(ロス後)/0.5 の順で選ぶ。
    """
    if preferred and preferred in df.columns:
        return df[preferred].astype(float)
    candidates = ["需要計画量(ロス前)", "需要計画量", "需要kW"]
    for c in candidates:
        if c in df.columns and df[c].notna().any():
            return df[c].astype(float)
    # fallback: 使用電力量(ロス後)をkW換算
    return (df["使用電力量(ロス後)"].astype(float) / 0.5)

def pick_generation_series(df, preferred=None):
    """
    自家発がなければゼロ系列を返す。
    """
    if preferred and preferred in df.columns:
        return df[preferred].astype(float)
    for c in ["自家発出力", "PV出力", "太陽光出力", "発電kW"]:
        if c in df.columns and df[c].notna().any():
            return df[c].astype(float)
    return pd.Series(0.0, index=df.index)

def compute_export_offer_def1(df, P_pcs=1000.0, P_exp_max=None, load_col=None, gen_col=None):
    """
    供出可能量（定義①）: max(0, P_pcs - (L - G)), 逆潮上限でクリップ
    """
    L = pick_load_series(df, preferred=load_col)
    G = pick_generation_series(df, preferred=gen_col)
    offer = (P_pcs - (L - G)).clip(lower=0)
    if P_exp_max is not None:
        offer = offer.clip(upper=float(P_exp_max))
    return offer, L, G

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
