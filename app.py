
import io
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from utils_timeseries import (
    load_excel_to_df, select_range, series_picker, aggregate_df,
    list_dates, get_day_slice, overlay_by_dates, overlay_by_dates_price, overlay_price_full_year,
    plot_lines, compute_export_offer_def1,
    simulate_soc_with_charge_periodic_reset
)

st.set_page_config(page_title="鳥栖PO1期 可視化ツール", layout="wide")
st.title("鳥栖PO1期 可視化ツール（kW/価格/オーバレイ/単独/供出可能量①/SOC充電対応）")

with st.sidebar:
    st.header("データ入力")
    up = st.file_uploader("Excel（.xlsx）をアップロード", type=["xlsx"])
    sheet_name = st.text_input("シート名（未入力なら先頭シート）", value="")
    st.divider()
    st.subheader("共通パラメータ")
    P_pcs_common = st.number_input("PCS定格（kW）", min_value=1, value=1000, step=10)

if up is None:
    st.info("左のサイドバーからExcelファイルをアップロードしてください。")
    st.stop()

try:
    df = load_excel_to_df(up, sheet_name)
except Exception as e:
    st.error(f"読み込みエラー: {e}")
    st.stop()

st.success("データの読み込みに成功しました。")
has_price = "JEPXスポットプライス" in df.columns and df["JEPXスポットプライス"].notna().any()

min_t, max_t = df.index.min(), df.index.max()
st.caption(f"データ期間: {min_t} 〜 {max_t}（JEPX価格列: {'あり' if has_price else 'なし'}）")

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "1) 基本プロット（kW + 価格）",
    "2) 集計（kW/価格）",
    "4) オーバレイ（kW/価格）",
    "5) 単独表示（kW/価格・範囲指定）",
    "6) 供出可能量（①：1000-(L-G)）",
    "7) 価格：1年分オーバレイ",
    "8) SOCシミュレーション（充電コマ考慮・期間指定）",
])

# --- Tab1 ---
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
    ax.set_xlabel("時刻"); ax.set_ylabel("平均出力 (kW)")
    title = "平均出力（30分・kW）"
    if show_price:
        ax2 = ax.twinx()
        ax2.plot(dfr.index, dfr["JEPXスポットプライス"])
        ax2.set_ylabel("JEPXスポットプライス (円/kWh)")
        title += " + 価格"
    ax.set_title(title); ax.legend(loc="upper left"); ax.grid(True)
    st.pyplot(fig)

# --- Tab2 ---
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
    ax.set_ylabel(f"{how} kW（{'日平均' if agg_code=='D' else '月平均'}）")
    title2 = f"kW {('日' if agg_code=='D' else '月')}集計（{how}）"
    if show_price2:
        price_series = aggregate_df(dfr2[["JEPXスポットプライス"]], aggregate=agg_code, how="mean")
        ax2 = ax.twinx(); ax2.plot(price_series.index, price_series["JEPXスポットプライス"])
        ax2.set_ylabel("JEPXスポットプライス 平均 (円/kWh)"); title2 += " + 価格(平均)"
    ax.set_title(title2); ax.legend(loc="upper left"); ax.grid(True)
    st.pyplot(fig2)

# --- Tab3 ---
with tab3:
    st.subheader("オーバレイ（kW/価格）")
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
        dates = [f"{m}-{day_of_month:02d}" for m in months]
    else:
        md = st.text_input("月日（MM-DD）", value="08-15", key="t4_md")
        years = st.multiselect("対象年", sorted(catalog["year"].unique().tolist()), default=sorted(catalog["year"].unique().tolist()), key="t4_years")
        dates = [f"{y}-{md}" for y in years]
    if st.button("プロット", type="primary", key="t4_btn"):
        if target == "出力(kW)":
            mat = overlay_by_dates(df, dates, which=which); ylabel = "平均出力 (kW)"; title = f"日曲線オーバレイ（{which}）"
        else:
            mat = overlay_by_dates_price(df, dates); ylabel = "JEPXスポットプライス (円/kWh)"; title = "日曲線オーバレイ（JEPX価格）"
        if mat.empty:
            st.warning("該当するデータがありません。")
        else:
            fig3, ax = plt.subplots(figsize=(12,6))
            for col in mat.columns:
                ax.plot(range(48), mat[col], label=col)
            ax.set_xlabel("時刻（30分刻み、0=0:00 … 47=23:30）"); ax.set_ylabel(ylabel); ax.set_title(title); ax.legend(); ax.grid(True)
            st.pyplot(fig3)
            st.download_button("CSVをダウンロード", data=mat.to_csv(index_label="slot(30min)").encode("utf-8-sig"),
                               file_name=("overlay_kw.csv" if target=="出力(kW)" else "overlay_jepx.csv"), mime="text/csv")

# --- Tab4 ---
with tab4:
    st.subheader("単独表示（kW/価格・範囲指定）")
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
        show_price5 = st.checkbox("JEPX価格も表示（右軸）", value=has_price, disabled=not has_price, key="t5_price")
    dfr5 = select_range(df, pd.Timestamp(start5), pd.Timestamp(end5) + pd.Timedelta(days=1))
    plot_df5 = series_picker(dfr5, series=series5, use_kw=True)
    agg_code5 = None if agg5.startswith("30") else ("D" if agg5.startswith("日") else "M")
    plot_df5 = aggregate_df(plot_df5, aggregate=agg_code5, how="mean")
    fig5, ax = plt.subplots(figsize=(12,6))
    for col in plot_df5.columns: ax.plot(plot_df5.index, plot_df5[col], label=col)
    ax.set_xlabel("時刻" if agg_code5 is None else ("日付" if agg_code5=="D" else "年月")); ax.set_ylabel("平均出力 (kW)"); title5 = "単独表示（kW）"
    if show_price5:
        price_plot = aggregate_df(dfr5[["JEPXスポットプライス"]], aggregate=agg_code5, how="mean")
        ax2 = ax.twinx(); ax2.plot(price_plot.index, price_plot["JEPXスポットプライス"]); ax2.set_ylabel("JEPXスポットプライス (円/kWh)"); title5 += " + 価格"
    ax.set_title(title5); ax.legend(loc="upper left"); ax.grid(True); st.pyplot(fig5)

# --- Tab5: Export offer def1 ---
with tab5:
    st.subheader("供出可能量（定義①：1000-(L-G)）")
    c1, c2, c3 = st.columns(3)
    with c1:
        start6 = st.date_input("開始日", value=min_t.date(), key="t6_start")
    with c2:
        end6 = st.date_input("終了日", value=max_t.date(), key="t6_end")
    with c3:
        P_exp_max = st.text_input("逆潮上限（kW／空欄=無制限）", value="")
    load_col = st.selectbox("需要列の選択（なければ自動推定）", ["自動", "需要計画量(ロス前)", "需要計画量", "需要kW"], index=0, key="t6_load")
    gen_col = st.selectbox("自家発列の選択（無ければなし）", ["自動", "自家発出力", "PV出力", "太陽光出力", "発電kW"], index=0, key="t6_gen")
    P_exp_max_val = None
    try:
        P_exp_max_val = float(P_exp_max) if len(P_exp_max.strip()) > 0 else None
    except Exception:
        P_exp_max_val = None
    dfr6 = select_range(df, pd.Timestamp(start6), pd.Timestamp(end6) + pd.Timedelta(days=1))
    offer, L, G = compute_export_offer_def1(dfr6, P_pcs=P_pcs_common, P_exp_max=P_exp_max_val,
                                            load_col=(None if load_col=="自動" else load_col),
                                            gen_col=(None if gen_col=="自動" else gen_col))
    min_val = offer.min(); min_ts = offer.idxmin()
    fig6, ax = plt.subplots(figsize=(12,6))
    ax.plot(offer.index, offer.values, label="供出可能量(①)"); ax.axhline(min_val, linestyle="--", label=f"最小値 {min_val:.1f} kW")
    x_pos = offer.index[int(len(offer)*0.6)]; ax.text(x_pos, float(min_val), f"最小値 {min_val:.1f} kW @ {min_ts}", bbox=dict(facecolor="white", alpha=0.7))
    ax.set_xlabel("時刻"); ax.set_ylabel("供出可能量 (kW)"); ax.set_title("一次調整力 供出可能量（定義①）— 推移と最小値"); ax.grid(True); ax.legend(); st.pyplot(fig6)
    out_df = pd.DataFrame({"供出可能量kW(①=PCS-(L-G))": offer, "需要kW(L)": L, "自家発kW(G)": G})
    st.download_button("CSVをダウンロード", data=out_df.to_csv().encode("utf-8-sig"), file_name="export_offer_def1.csv", mime="text/csv")

# --- Tab6: Price full-year overlay ---
with tab6:
    st.subheader("JEPXスポットプライス：1年分オーバレイ（各日×48スロット）")
    ymax = st.number_input("縦軸上限（円/kWh）", min_value=10, value=40, step=5)
    mat = overlay_price_full_year(df)
    if mat.empty:
        st.warning("価格列が見つからないか、データがありません。")
    else:
        fig7, ax = plt.subplots(figsize=(12,6))
        for col in mat.columns:
            ax.plot(range(48), mat[col], alpha=0.2, linewidth=0.7)
        ax.set_xlabel("時刻スロット (0=0:00, ..., 47=23:30)"); ax.set_ylabel("JEPXスポットプライス (円/kWh)")
        ax.set_title("JEPXスポットプライス 日曲線オーバレイ（全日）"); ax.grid(True); ax.set_xlim(0,47); ax.set_ylim(0, ymax)
        ax.set_xticks(range(0, 48, 4))
        st.pyplot(fig7)
        st.download_button("CSVをダウンロード（48×日数）", data=mat.to_csv(index_label="slot(30min)").encode("utf-8-sig"),
                           file_name="jepx_overlay_full_year.csv", mime="text/csv")

# --- Tab7: SOC simulation with charge and period selection ---
with tab7:
    st.subheader("SOCシミュレーション（充電コマ考慮・期間指定）")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        start_soc = st.date_input("開始日", value=min_t.date(), key="t7_start")
    with c2:
        end_soc = st.date_input("終了日", value=max_t.date(), key="t7_end")
    with c3:
        soc_init_pct = st.number_input("初期SOC（%）", min_value=1.0, max_value=100.0, value=90.0, step=1.0)
    with c4:
        soc_floor_pct = st.number_input("下限SOC（%）", min_value=0.0, max_value=90.0, value=10.0, step=1.0)
    with c5:
        reset_days = st.number_input("充電間隔（日）", min_value=1, value=4, step=1)
    c6, c7, c8 = st.columns(3)
    with c6:
        E_nom = st.number_input("電池容量（kWh）", min_value=100, value=2000, step=100)
    with c7:
        P_chg = st.number_input("充電出力（kW）", min_value=1, value=1000, step=10)
    with c8:
        P_pcs_for_soc = st.number_input("PCS定格（kW）", min_value=1, value=1000, step=10)

    load_col7 = st.selectbox("需要列（自動推定可）", ["自動", "需要計画量(ロス前)", "需要計画量", "需要kW"], index=0, key="t7_load")
    gen_col7 = st.selectbox("自家発列（無ければなし）", ["自動", "自家発出力", "PV出力", "太陽光出力", "発電kW"], index=0, key="t7_gen")

    soc_df = simulate_soc_with_charge_periodic_reset(
        df,
        P_pcs=P_pcs_for_soc, P_chg=P_chg, E_nom=E_nom,
        start=pd.Timestamp(start_soc), end=pd.Timestamp(end_soc) + pd.Timedelta(days=1) - pd.Timedelta(minutes=30),
        soc_init_pct=soc_init_pct, soc_floor_pct=soc_floor_pct, reset_every_days=reset_days,
        load_col=(None if load_col7=="自動" else load_col7),
        gen_col=(None if gen_col7=="自動" else gen_col7)
    )

    if soc_df.empty:
        st.warning("SOCシミュレーションに必要なデータが不足しています。")
    else:
        fig8, ax = plt.subplots(figsize=(12,6))
        ax.plot(soc_df.index, soc_df["SOC_%"], drawstyle="steps-post", label="SOC（%）")
        chg_idx = soc_df.index[soc_df["charging"]]
        if len(chg_idx) > 0:
            ax.scatter(chg_idx, soc_df.loc[chg_idx, "SOC_%"], s=20, label="充電スロット")
        ax.axhline(soc_floor_pct, linestyle="--", label=f"下限 {soc_floor_pct:.1f}%")
        ax.axhline(soc_init_pct, linestyle="--", label=f"初期 {soc_init_pct:.1f}%")
        ax.set_xlabel("時刻"); ax.set_ylabel("SOC (%)"); ax.set_title("SOCの推移（充電コマ考慮）")
        ax.grid(True)
        handles, labels = ax.get_legend_handles_labels()
        uniq = dict(zip(labels, handles))
        ax.legend(uniq.values(), uniq.keys())
        st.pyplot(fig8)

        st.download_button("CSVをダウンロード（SOC/充電コマ）", data=soc_df.to_csv().encode("utf-8-sig"),
                           file_name="soc_with_charge_and_period.csv", mime="text/csv")
