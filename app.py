import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import fitz  # PyMuPDF
import tempfile
import io
import os

st.set_page_config(page_title="後方数値データ分析", layout="wide")
st.title("📊 後方数値データ分析ダッシュボード")

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"], key="excel_upload")
if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    df.columns = [str(c).strip().replace('\u3000', '').replace('\xa0', '') for c in df.columns]

    numeric_cols = ['年齢', '年収', '同借希望額', '住宅ローン返済月額', '勤続年数', '他社借入件数',
                    '取扱金額_申込当月', '取扱金額_申込翌月末', '取扱金額_申込翌々月末']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    if '申込日' in df.columns:
        df['申込日'] = pd.to_datetime(df['申込日'], errors='coerce')

    df['取扱高'] = df[['取扱金額_申込当月', '取扱金額_申込翌月末', '取扱金額_申込翌々月末']].sum(axis=1)

    st.sidebar.header("フィルタ設定")
    start_date, end_date = st.sidebar.date_input("申込日範囲", [df['申込日'].min(), df['申込日'].max()])
    media_codes = df['媒体コード'].dropna().unique().tolist() if '媒体コード' in df.columns else []
    selected_codes = st.sidebar.multiselect("媒体コードを選択（ALL選択で全件）", ["ALL"] + media_codes, default=["ALL"])

    filtered_df = df[(df['申込日'] >= pd.to_datetime(start_date)) & (df['申込日'] <= pd.to_datetime(end_date))]
    if "ALL" not in selected_codes:
        filtered_df = filtered_df[filtered_df['媒体コード'].isin(selected_codes)]

    st.write(f"件数: {len(filtered_df)}")

    def group_age(x):
        if pd.isna(x): return "不明"
        if x < 20: return "10代"
        elif x < 30: return "20代"
        elif x < 40: return "30代"
        elif x < 50: return "40代"
        elif x < 60: return "50代"
        else: return "60代以上"

    def group_income(x):
        if pd.isna(x): return "不明"
        if x < 500: return "0-499"
        elif x < 1000: return "500-999"
        else: return "1000以上"

    filtered_df['年代'] = filtered_df['年齢'].apply(group_age)
    filtered_df['年収帯'] = filtered_df['年収'].apply(group_income)

    def create_dual_axis_chart(df, category_col, title):
        count_data = df[category_col].value_counts().sort_index()
        sum_data = df.groupby(category_col)['取扱高'].sum().reindex(count_data.index)

        fig = go.Figure()
        fig.add_trace(go.Bar(x=count_data.index, y=count_data.values, name="件数", marker_color="skyblue", yaxis="y"))
        fig.add_trace(go.Bar(x=sum_data.index, y=sum_data.values, name="取扱高（円）", marker_color="orange", yaxis="y2"))
        fig.update_layout(title=title, xaxis=dict(title=category_col),
                          yaxis=dict(title="件数"), yaxis2=dict(title="取扱高（円）", overlaying="y", side="right"),
                          barmode="group")
        return fig

    st.subheader("📈 二軸棒グラフ")
    figs = []
    for col, title in [("年代", "年代別"), ("年収帯", "年収帯別")]:
        if col in filtered_df.columns:
            fig = create_dual_axis_chart(filtered_df, col, title)
            st.plotly_chart(fig, use_container_width=True)
            figs.append((fig, title))

    st.subheader("🔍 クロス集計（年代 × 年収帯）")
    pivot_count = pd.pivot_table(filtered_df, index='年代', columns='年収帯', aggfunc='size', fill_value=0)
    pivot_sum = pd.pivot_table(filtered_df, index='年代', columns='年収帯', values='取扱高', aggfunc='sum', fill_value=0)
    st.write("件数")
    st.dataframe(pivot_count)
    st.write("取扱高（円）")
    st.dataframe(pivot_sum)

    def create_pdf(figs, pivot_count, pivot_sum):
        pdf = fitz.open()

        # 表紙
        page = pdf.new_page()
        text = f"後方数値データ分析レポート\n\n期間: {start_date} ～ {end_date}\n媒体コード: {'ALL' if 'ALL' in selected_codes else '媒体コード指定'}\n件数: {len(filtered_df)}"
        page.insert_text((50, 50), text, fontsize=12)

        # グラフページ
        for fig, title in figs:
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
                fig.write_image(tmpfile.name, format="png")
                img_rect = fitz.Rect(50, 100, 550, 500)
                page = pdf.new_page()
                page.insert_text((50, 50), title, fontsize=14)
                page.insert_image(img_rect, filename=tmpfile.name)
                os.unlink(tmpfile.name)

        # クロス集計ページ（件数）
        page = pdf.new_page()
        page.insert_text((50, 50), "クロス集計：件数（年代 × 年収帯）", fontsize=14)
        table_text = pivot_count.to_string()
        page.insert_text((50, 80), table_text, fontsize=8)

        # クロス集計ページ（取扱高）
        page = pdf.new_page()
        page.insert_text((50, 50), "クロス集計：取扱高（年代 × 年収帯）", fontsize=14)
        table_text2 = pivot_sum.to_string()
        page.insert_text((50, 80), table_text2, fontsize=8)

        pdf_bytes = pdf.write()
        return io.BytesIO(pdf_bytes)

    pdf_stream = create_pdf(figs, pivot_count, pivot_sum)
    st.download_button("📥 PDFレポートをダウンロード", data=pdf_stream, file_name="分析レポート.pdf", mime="application/pdf")

else:
    st.info("Excelファイルをアップロードしてください。")