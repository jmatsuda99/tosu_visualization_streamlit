
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import rcParams

# ---- Global font: try Noto Sans CJK JP (fallback to default) ----
try:
    rcParams["font.family"] = "Noto Sans CJK JP"
except Exception:
    pass

REQUIRED_COLUMNS = ["開始日時", "使用電力量(ロス後)", "使用電力量(ロス前)"]

def load_excel_to_df(file, sheet_name=None):
    """
    Load the Excel file into a cleaned dataframe:
    - Ensures required columns exist
    - Drops total/annotation rows by excluding NaN '終了日時' if present
    - Parses '開始日時' with timezone if included, sorts and sets as index
    - Adds kW columns (kWh / 0.5h)
    """
    if sheet_name is None:
        xls = pd.ExcelFile(file)
        sheet_name = xls.sheet_names[0]
    df = pd.read_excel(file, sheet_name=sheet_name)

    # Filter out non-time-series rows where '終了日時' is NaN (if column exists)
    if "終了日時" in df.columns:
        df = df[df["終了日時"].notna()].copy()

    # Ensure required columns
    for c in REQUIRED_COLUMNS:
        if c not in df.columns:
            raise ValueError(f"必須列が見つかりません: {c}")

    # Parse datetime and index
    df["開始日時"] = pd.to_datetime(df["開始日時"], errors="coerce")
    df = df.dropna(subset=["開始日時"]).copy()
    df = df.set_index("開始日時").sort_index()

    # kW columns
    df["使用電力量(ロス後)_kW"] = df["使用電力量(ロス後)"] / 0.5  # 30min interval -> average kW
    df["使用電力量(ロス前)_kW"] = df["使用電力量(ロス前)"] / 0.5

    return df

def select_range(df, start=None, end=None):
    """
    Select by time range (inclusive). If df index has tz, localize naive boundaries to the same tz.
    """
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
    """
    Return dataframe with selected series.
    """
    if use_kw:
        cols = {
            "ロス後": "使用電力量(ロス後)_kW",
            "ロス前": "使用電力量(ロス前)_kW",
        }
    else:
        cols = {
            "ロス後": "使用電力量(ロス後)",
            "ロス前": "使用電力量(ロス前)",
        }

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
    """
    Aggregate time series.
      aggregate: None (raw 30min), 'D' (daily), 'M' (monthly, MS)
      how: 'mean' (default) or 'sum'/'max' etc.
    """
    if aggregate is None:
        return df
    if aggregate == "D":
        return getattr(df.resample("D"), how)()
    if aggregate == "M":
        return getattr(df.resample("MS"), how)()
    raise ValueError("aggregate には None / 'D' / 'M' を指定してください。")

def list_dates(df):
    """
    Return DataFrame with unique dates and year/month/day helpers.
    """
    uniq = pd.to_datetime(df.index.date).unique()
    catalog = pd.DataFrame({"date": pd.to_datetime(uniq)})
    catalog["year"] = catalog["date"].dt.year
    catalog["month"] = catalog["date"].dt.month
    catalog["day"] = catalog["date"].dt.day
    catalog["month_label"] = catalog["date"].dt.strftime("%Y-%m")
    return catalog.sort_values("date")

def get_day_slice(df, date_val):
    """
    Return the one-day slice for date_val (string or date), inclusive [00:00,24:00).
    """
    # keep timezone if exists
    tz = df.index.tz
    start = pd.Timestamp(date_val)
    if tz is not None and start.tzinfo is None:
        start = start.tz_localize(tz)
    end = start + pd.Timedelta(days=1)
    return df.loc[(df.index >= start) & (df.index < end)].copy()

def overlay_by_dates(df, dates, which="ロス後"):
    """
    Build matrix for overlay plot by specific date list.
    Returns index slots (0..47) and matrix DataFrame with each column as a date.
    """
    mat = pd.DataFrame(index=range(48))
    for d in dates:
        day = get_day_slice(df, d)
        if day.empty:
            continue
        # normalize to 30min slots 0..47 in local time
        idx_local = day.index.tz_convert("Asia/Tokyo") if day.index.tz is not None else day.index
        slots = ((idx_local - idx_local.normalize()) / pd.Timedelta(minutes=30)).astype(int)
        ser = day[f"使用電力量({which})_kW"]
        ser = pd.Series(ser.values, index=slots).reindex(range(48))
        mat[str(pd.to_datetime(d).date())] = ser.values
    return mat

def plot_lines(x, y_dict, xlabel, ylabel, title):
    plt.figure(figsize=(12, 6))
    for label, series in y_dict.items():
        plt.plot(x, series, label=label)
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.title(title)
    plt.legend()
    plt.grid(True)
    st = plt.gcf()
    return st
