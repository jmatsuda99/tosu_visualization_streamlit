
import io
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from utils_timeseries import (
    load_excel_to_df, select_range, series_picker, aggregate_df,
    list_dates, get_day_slice, overlay_by_dates, overlay_by_dates_price, plot_lines,
    compute_export_offer_def1
)

st.set_page_config(page_title="鳥栖PO1期 可視化ツール", layout="wide")

st.title("鳥栖PO1期 可視化ツール（kW換算／価格／オーバレイ／単独表示／供出可能量①）")

with st.sidebar:
    st.header("データ入力")
    up = st.file_uploader("Excel（.xlsx）をアップロード", type=["xlsx"])
    sheet_name = st.text_input("シート名（未入力なら先頭シート）", value="")
    st.divider()
    st.subheader("供出可能量（①）計算パラメータ")
    P_pcs = st.number_input("PCS定格（kW）", min_value=1, value=1000, step=10)
    P_exp_max = st.text_input("逆潮上限（kW／空欄=無制限）", value="")

if up is None:
    st.info("左のサイドバーからExcelファイルをアップロードしてください。")
    st.stop()

# Load
try:
    df = load_excel_to_df(up, sheet_name)
except Exception as e:
    st.error(f"読み込みエラー: {e}")
    st.stop()

st.success("データの読み込みに成功しました。")
has_price = "JEPXスポットプライス" in df.columns and df["JEPXスポットプライス"].notna().any()

# Common widgets
min_t, max_t = df.index.min(), df.index.max()
st.caption(f"データ期間: {min_t} 〜 {max_t}（JEPX価格列: {'あり' if has_price else 'なし'}）")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "1) 基本プロット（kW + 価格）",
    "2) 集計（kW/価格）",
    "4) オーバレイ（kW/価格）",
    "5) 単独表示（kW/価格・範囲指定）",
    "6) 供出可能量（①：1000-(L-G)）",
])

# ---------- Tab1: Basic plot ----------
with tab1:
    st.subheader("基本プロット（30分）")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        series = st.selectbox("系列（出力）", ["both", "ロス後", "ロス前"], index=0, key="t1_series")
    with c2:
        start = st.date_input("開始日", value=min_t.date(), min_value=min_t.date(), max_value=max_t.date(), key="t1_start")
    with c3:
        end = st.date_input("終了日", value=max_t.date(), min_value=min_t.date(), max_value=max_t.date(), key="t1_end")
    with c4:
        show_price = st.checkbox("JEPXスポットプライスを右軸に表示", value=has_price, disabled=not has_price)

    dfr = select_range(df, pd.Timestamp(start), pd.Timestamp(end) + pd.Timedelta(days=1))
    plot_df = series_picker(dfr, series=series, use_kw=True)

    fig, ax = plt.subplots(figsize=(12,6))
    for col in plot_df.columns:
        ax.plot(plot_df.index, plot_df[col], label=col)
    ax.set_xlabel("時刻")
    ax.set_ylabel("平均出力 (kW)")
    title = "平均出力（30分・kW）"
    if show_price:
        ax2 = ax.twinx()
        ax2.plot(dfr.index, dfr["JEPXスポットプライス"])
        ax2.set_ylabel("JEPXスポットプライス (円/kWh)")
        title += " + JEPX価格"
    ax.set_title(title)
    ax.legend(loc="upper left")
    ax.grid(True)
    st.pyplot(fig)

# ---------- Tab2: Aggregation ----------
with tab2:
    st.subheader("集計（kW/価格）")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        agg = st.selectbox("粒度", ["日平均(D)", "月平均(M)"], index=0, key="t2_agg")
    with c2:
        how = st.selectbox("集計方法（kW）", ["mean", "max", "min"], index=0, key="t2_how")
    with c3:
        series2 = st.selectbox("系列（kW）", ["both", "ロス後", "ロス前"], index=0, key="t2_series")
    with c4:
        start2 = st.date_input("開始日", value=min_t.date(), key="t2_start")
    with c5:
        end2 = st.date_input("終了日", value=max_t.date(), key="t2_end")
    show_price2 = st.checkbox("JEPX価格も表示（右軸：平均）", value=has_price, disabled=not has_price, key="t2_price")

    dfr2 = select_range(df, pd.Timestamp(start2), pd.Timestamp(end2) + pd.Timedelta(days=1))
    plot_df2 = series_picker(dfr2, series=series2, use_kw=True)
    agg_code = "D" if agg.startswith("日") else "M"
    plot_df2 = aggregate_df(plot_df2, aggregate=agg_code, how=how)

    fig2, ax = plt.subplots(figsize=(12,6))
    for col in plot_df2.columns:
        ax.plot(plot_df2.index, plot_df2[col], label=col)
    ax.set_xlabel("日付" if agg_code=="D" else "年月")
    ylabel = f"{how} kW（{'日平均' if agg_code=='D' else '月平均'}）"
    ax.set_ylabel(ylabel)
    title2 = f"kW {('日' if agg_code=='D' else '月')}集計（{how}）"

    if show_price2:
        price_series = aggregate_df(dfr2[["JEPXスポットプライス"]], aggregate=agg_code, how="mean")
        ax2 = ax.twinx()
        ax2.plot(price_series.index, price_series["JEPXスポットプライス"])
        ax2.set_ylabel("JEPXスポットプライス 平均 (円/kWh)")
        title2 += " + 価格(平均)"
    ax.set_title(title2)
    ax.legend(loc="upper left")
    ax.grid(True)
    st.pyplot(fig2)

# ---------- Tab3: Overlay ----------
with tab3:
    st.subheader("オーバレイ")
    from utils_timeseries import list_dates
    catalog = list_dates(df)

    target = st.radio("対象", ["出力(kW)", "JEPXスポットプライス"], horizontal=True, key="t4_target")
    if target == "出力(kW)":
        which = st.selectbox("ロス前/後", ["ロス後", "ロス前"], index=0, key="t4_which")
    mode = st.radio("オーバレイ種別", ["指定日", "月ごと同日", "年ごと同月日"], horizontal=True, key="t4_mode")

    if mode == "指定日":
        choices = st.multiselect("日付を選択", catalog["date"].dt.strftime("%Y-%m-%d").tolist(), max_selections=20, key="t4_dates")
        dates = choices
    elif mode == "月ごと同日":
        day_of_month = st.number_input("日（1〜31）", min_value=1, max_value=31, value=15, step=1, key="t4_dom")
        months = st.multiselect("対象月（YYYY-MM）", catalog["month_label"].unique().tolist(), default=catalog["month_label"].unique().tolist(), key="t4_months")
        dates = []
        for m in months:
            try:
                y, mo = m.split("-")
                dates.append(f"{y}-{mo}-{day_of_month}")
            except Exception:
                pass
    else:  # 年ごと同月日
        md = st.text_input("月日（MM-DD）", value="08-15", key="t4_md")
        years = st.multiselect("対象年", sorted(catalog["year"].unique().tolist()), default=sorted(catalog["year"].unique().tolist()), key="t4_years")
        dates = [f"{y}-{md}" for y in years]

    if st.button("プロット", type="primary", key="t4_btn"):
        if target == "出力(kW)":
            from utils_timeseries import overlay_by_dates
            mat = overlay_by_dates(df, dates, which=which)
            ylabel = "平均出力 (kW)"
            title = f"日曲線オーバレイ（{which}）"
        else:
            from utils_timeseries import overlay_by_dates_price
            mat = overlay_by_dates_price(df, dates)
            ylabel = "JEPXスポットプライス (円/kWh)"
            title = "日曲線オーバレイ（JEPX価格）"

        if mat.empty:
            st.warning("該当するデータがありません。")
        else:
            fig3, ax = plt.subplots(figsize=(12,6))
            for col in mat.columns:
                ax.plot(range(48), mat[col], label=col)
            ax.set_xlabel("時刻（30分刻み、0=0:00 … 47=23:30）")
            ax.set_ylabel(ylabel)
            ax.set_title(title)
            ax.legend()
            ax.grid(True)
            st.pyplot(fig3)
            st.download_button("CSVをダウンロード", data=mat.to_csv(index_label="slot(30min)").encode("utf-8-sig"),
                               file_name=("overlay_kw.csv" if target=="出力(kW)" else "overlay_jepx.csv"), mime="text/csv")

# ---------- Tab4: Single (non-overlay) ----------
with tab4:
    st.subheader("単独表示（範囲指定）")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        agg5 = st.selectbox("粒度", ["30分(raw)", "日平均(D)", "月平均(M)"], index=0, key="t5_agg")
    with c2:
        series5 = st.selectbox("系列（kW）", ["both", "ロス後", "ロス前"], index=0, key="t5_series")
    with c3:
        start5 = st.date_input("開始日", value=min_t.date(), key="t5_start")
    with c4:
        end5 = st.date_input("終了日", value=max_t.date(), key="t5_end")
    with c5:
        show_price5 = st.checkbox("JEPX価格も表示（右軸：平均/そのまま）", value=has_price, disabled=not has_price, key="t5_price")

    dfr5 = select_range(df, pd.Timestamp(start5), pd.Timestamp(end5) + pd.Timedelta(days=1))
    plot_df5 = series_picker(dfr5, series=series5, use_kw=True)
    agg_code5 = None if agg5.startswith("30") else ("D" if agg5.startswith("日") else "M")
    plot_df5 = aggregate_df(plot_df5, aggregate=agg_code5, how="mean")

    fig5, ax = plt.subplots(figsize=(12,6))
    for col in plot_df5.columns:
        ax.plot(plot_df5.index, plot_df5[col], label=col)
    ax.set_xlabel("時刻" if agg_code5 is None else ("日付" if agg_code5=="D" else "年月"))
    ax.set_ylabel("平均出力 (kW)")
    title5 = "単独表示（kW）"

    if show_price5:
        price_plot = dfr5[["JEPXスポットプライス"]].copy()
        price_plot = aggregate_df(price_plot, aggregate=agg_code5, how="mean")
        ax2 = ax.twinx()
        ax2.plot(price_plot.index, price_plot["JEPXスポットプライス"])
        ax2.set_ylabel("JEPXスポットプライス (円/kWh)")
        title5 += " + 価格"

    ax.set_title(title5)
    ax.legend(loc="upper left")
    ax.grid(True)
    st.pyplot(fig5)

# ---------- Tab5: 供出可能量（定義①） ----------
with tab5:
    st.subheader("供出可能量（定義①：1000-(L-G)）")
    c1, c2, c3 = st.columns(3)
    with c1:
        start6 = st.date_input("開始日", value=min_t.date(), key="t6_start")
    with c2:
        end6 = st.date_input("終了日", value=max_t.date(), key="t6_end")
    with c3:
        load_col = st.selectbox("需要列の選択（なければ自動推定）", ["自動", "需要計画量(ロス前)", "需要計画量", "需要kW"], index=0)
    gen_col = st.selectbox("自家発列の選択（無ければなし）", ["自動", "自家発出力", "PV出力", "太陽光出力", "発電kW"], index=0)

    # parse exp max
    P_exp_max_val = None
    try:
        P_exp_max_val = float(P_exp_max) if len(P_exp_max.strip()) > 0 else None
    except Exception:
        P_exp_max_val = None

    dfr6 = select_range(df, pd.Timestamp(start6), pd.Timestamp(end6) + pd.Timedelta(days=1))
    offer, L, G = compute_export_offer_def1(
        dfr6,
        P_pcs=float(P_pcs),
        P_exp_max=P_exp_max_val,
        load_col=(None if load_col=="自動" else load_col),
        gen_col=(None if gen_col=="自動" else gen_col),
    )

    # 時系列グラフ + 最小値の点線表示とラベル
    min_val = offer.min()
    min_ts = offer.idxmin()
    fig6, ax = plt.subplots(figsize=(12,6))
    ax.plot(offer.index, offer.values, label="供出可能量(①)")
    ax.axhline(min_val, linestyle="--", label=f"最小値 {min_val:.1f} kW")
    # 注記を右寄りに配置
    x_pos = offer.index[int(len(offer)*0.6)]
    ax.text(x_pos, float(min_val), f"最小値 {min_val:.1f} kW @ {min_ts}", bbox=dict(facecolor="white", alpha=0.7))
    ax.set_xlabel("時刻")
    ax.set_ylabel("供出可能量 (kW)")
    ax.set_title("一次調整力 供出可能量（定義①）— 推移と最小値")
    ax.grid(True)
    ax.legend()
    st.pyplot(fig6)

    # CSVダウンロード
    out_df = pd.DataFrame({
        "供出可能量kW(①=PCS-(L-G))": offer,
        "需要kW(L)": L,
        "自家発kW(G)": G,
    })
    st.download_button("CSVをダウンロード", data=out_df.to_csv().encode("utf-8-sig"),
                       file_name="export_offer_def1.csv", mime="text/csv")
