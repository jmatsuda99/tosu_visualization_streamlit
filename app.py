
import io
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from utils_timeseries import (
    load_excel_to_df, select_range, series_picker, aggregate_df,
    list_dates, get_day_slice, overlay_by_dates, plot_lines
)

st.set_page_config(page_title="鳥栖PO1期 可視化ツール", layout="wide")

st.title("鳥栖PO1期 可視化ツール（kW換算・オーバレイ／単独表示）")

with st.sidebar:
    st.header("データ入力")
    up = st.file_uploader("Excel（.xlsx）をアップロード", type=["xlsx"])
    sheet_name = st.text_input("シート名（未入力なら先頭シート）", value="")

if up is None:
    st.info("左のサイドバーからExcelファイルをアップロードしてください。")
    st.stop()

# Load
try:
    df = load_excel_to_df(up, sheet_name if sheet_name.strip() else None)
except Exception as e:
    st.error(f"読み込みエラー: {e}")
    st.stop()

st.success("データの読み込みに成功しました。")

# Common widgets
min_t, max_t = df.index.min(), df.index.max()
st.caption(f"データ期間: {min_t} 〜 {max_t}")

tab1, tab2, tab3, tab4 = st.tabs([
    "1) 基本プロット（kW）",
    "2) 集計（kWベース）",
    "4) オーバレイ（kW）",
    "5) 単独表示（kW・範囲指定）",
])

# ---------- Tab1: Basic plot ----------
with tab1:
    st.subheader("基本プロット（kW・30分）")
    c1, c2, c3 = st.columns(3)
    with c1:
        series = st.selectbox("系列", ["both", "ロス後", "ロス前"], index=0, key="t1_series")
    with c2:
        start = st.date_input("開始日", value=min_t.date(), min_value=min_t.date(), max_value=max_t.date(), key="t1_start")
    with c3:
        end = st.date_input("終了日", value=max_t.date(), min_value=min_t.date(), max_value=max_t.date(), key="t1_end")

    dfr = select_range(df, pd.Timestamp(start), pd.Timestamp(end) + pd.Timedelta(days=1))
    plot_df = series_picker(dfr, series=series, use_kw=True)

    fig = plt.figure(figsize=(12,6))
    for col in plot_df.columns:
        plt.plot(plot_df.index, plot_df[col], label=col)
    plt.xlabel("時刻")
    plt.ylabel("平均出力 (kW)")
    plt.title("平均出力（30分・kW）")
    plt.legend()
    plt.grid(True)
    st.pyplot(fig)

    # download
    csv = plot_df.to_csv().encode("utf-8-sig")
    st.download_button("CSVをダウンロード", data=csv, file_name="basic_plot_kw.csv", mime="text/csv")

# ---------- Tab2: Aggregation ----------
with tab2:
    st.subheader("集計（kWベース）")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        agg = st.selectbox("粒度", ["日平均(D)", "月平均(M)"], index=0, key="t2_agg")
    with c2:
        how = st.selectbox("集計方法", ["mean", "max", "min"], index=0, key="t2_how")
    with c3:
        series2 = st.selectbox("系列", ["both", "ロス後", "ロス前"], index=0, key="t2_series")
    with c4:
        start2 = st.date_input("開始日", value=min_t.date(), key="t2_start")
    end2 = st.date_input("終了日", value=max_t.date(), key="t2_end")

    dfr2 = select_range(df, pd.Timestamp(start2), pd.Timestamp(end2) + pd.Timedelta(days=1))
    plot_df2 = series_picker(dfr2, series=series2, use_kw=True)
    agg_code = "D" if agg.startswith("日") else "M"
    plot_df2 = aggregate_df(plot_df2, aggregate=agg_code, how=how)

    fig2 = plt.figure(figsize=(12,6))
    for col in plot_df2.columns:
        plt.plot(plot_df2.index, plot_df2[col], label=col)
    plt.xlabel("日付" if agg_code=="D" else "年月")
    ylabel = f"{how} kW（{'日平均' if agg_code=='D' else '月平均'}）"
    plt.ylabel(ylabel)
    plt.title(f"kW {('日' if agg_code=='D' else '月')}集計（{how}）")
    plt.legend()
    plt.grid(True)
    st.pyplot(fig2)

    csv2 = plot_df2.to_csv().encode("utf-8-sig")
    st.download_button("CSVをダウンロード", data=csv2, file_name="aggregate_kw.csv", mime="text/csv")

# ---------- Tab4: Overlay ----------
with tab3:
    st.subheader("オーバレイ（kW）")
    catalog = list_dates(df)

    mode = st.radio("オーバレイ種別", ["指定日", "月ごと同日", "年ごと同月日"], horizontal=True, key="t4_mode")
    which = st.selectbox("ロス前/後", ["ロス後", "ロス前"], index=0, key="t4_which")

    if mode == "指定日":
        choices = st.multiselect("日付を選択", catalog["date"].dt.strftime("%Y-%m-%d").tolist(), max_selections=20)
        dates = choices
    elif mode == "月ごと同日":
        day_of_month = st.number_input("日（1〜31）", min_value=1, max_value=31, value=15, step=1)
        months = st.multiselect("対象月（YYYY-MM）", catalog["month_label"].unique().tolist(), default=catalog["month_label"].unique().tolist())
        # build concrete dates
        dates = []
        for m in months:
            try:
                y, mo = m.split("-")
                dates.append(f"{y}-{mo}-{day_of_month}")
            except Exception:
                pass
    else:  # 年ごと同月日
        md = st.text_input("月日（MM-DD）", value="08-15")
        years = st.multiselect("対象年", sorted(catalog["year"].unique().tolist()), default=sorted(catalog["year"].unique().tolist()))
        dates = [f"{y}-{md}" for y in years]

    if st.button("プロット", type="primary"):
        mat = overlay_by_dates(df, dates, which=which)
        if mat.empty:
            st.warning("該当するデータがありません。")
        else:
            fig3 = plt.figure(figsize=(12,6))
            for col in mat.columns:
                plt.plot(range(48), mat[col], label=col)
            plt.xlabel("時刻（30分刻み、0=0:00 … 47=23:30）")
            plt.ylabel("平均出力 (kW)")
            plt.title(f"日曲線オーバレイ（{which}）")
            plt.legend()
            plt.grid(True)
            st.pyplot(fig3)

            csv3 = mat.to_csv(index_label="slot(30min)").encode("utf-8-sig")
            st.download_button("CSVをダウンロード", data=csv3, file_name="overlay_kw.csv", mime="text/csv")

# ---------- Tab5: Single (non-overlay) ----------
with tab4:
    st.subheader("単独表示（kW・範囲指定）")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        agg5 = st.selectbox("粒度", ["30分(raw)", "日平均(D)", "月平均(M)"], index=0, key="t5_agg")
    with c2:
        series5 = st.selectbox("系列", ["both", "ロス後", "ロス前"], index=0, key="t5_series")
    with c3:
        start5 = st.date_input("開始日", value=min_t.date(), key="t5_start")
    with c4:
        end5 = st.date_input("終了日", value=max_t.date(), key="t5_end")

    dfr5 = select_range(df, pd.Timestamp(start5), pd.Timestamp(end5) + pd.Timedelta(days=1))
    plot_df5 = series_picker(dfr5, series=series5, use_kw=True)
    agg_code5 = None if agg5.startswith("30") else ("D" if agg5.startswith("日") else "M")
    plot_df5 = aggregate_df(plot_df5, aggregate=agg_code5, how="mean")

    fig5 = plt.figure(figsize=(12,6))
    for col in plot_df5.columns:
        plt.plot(plot_df5.index if agg_code5 else plot_df5.index, plot_df5[col], label=col)
    plt.xlabel("時刻" if agg_code5 is None else ("日付" if agg_code5=="D" else "年月"))
    plt.ylabel("平均出力 (kW)")
    plt.title("単独表示（kW）")
    plt.legend()
    plt.grid(True)
    st.pyplot(fig5)

    csv5 = plot_df5.to_csv().encode("utf-8-sig")
    st.download_button("CSVをダウンロード", data=csv5, file_name="single_kw.csv", mime="text/csv")

st.caption("© Tosu PO1 Visualizer — built with Streamlit & Matplotlib")
