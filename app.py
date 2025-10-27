import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import os

st.set_page_config(page_title="後方数値データ分析", layout="wide")
st.title("📊 後方数値データ分析ダッシュボード")

# カテゴリ順序定義
category_orders = {
    "年収帯": ['0-499', '500-999', '1000以上'],
    "借入希望額帯": ['0', '1-9', '10-19', '20-29', '30-39', '40-49', '50-59', '60-69', '70-79', '80-89', '90-99', '100-199', '200-299', '300以上'],
    "住宅ローン帯": ['0', '1-9', '10-19', '20-29', '30-39', '40-49', '50-59', '60-69', '70-79', '80-89', '90-99', '100以上'],
    "勤続年数帯": ['0', '1-3', '4-9', '10-20', '21以上']
}

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
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

    if len(filtered_df) == 0:
        st.warning("データがありません。フィルタ条件を確認してください。")
    else:
        # カテゴリ分け関数
        def group_income(x):
            if pd.isna(x): return "不明"
            if x < 500: return "0-499"
            elif x < 1000: return "500-999"
            else: return "1000以上"

        def group_mortgage(x):
            if pd.isna(x): return "不明"
            if x == 0: return "0"
            elif x < 10: return "1-9"
            elif x < 100: return f"{int(x//10)*10}-{int(x//10)*10+9}"
            else: return "100以上"

        if '年収' in filtered_df.columns:
            filtered_df['年収帯'] = filtered_df['年収'].apply(group_income)
        if '住宅ローン返済月額' in filtered_df.columns:
            filtered_df['住宅ローン帯'] = filtered_df['住宅ローン返済月額'].apply(group_mortgage)

        st.subheader("🔍 クロス集計（件数＋取扱高）")
        selected_cols = st.multiselect("クロス集計する項目を選択", ['年収帯', '住宅ローン帯'])
        if len(selected_cols) >= 2:
            pivot_count = pd.pivot_table(filtered_df, index=selected_cols[0], columns=selected_cols[1], aggfunc='size', fill_value=0)
            pivot_sum = pd.pivot_table(filtered_df, index=selected_cols[0], columns=selected_cols[1], values='取扱高', aggfunc='sum', fill_value=0)
            st.write("件数")
            st.dataframe(pivot_count)
            st.write("取扱高（円）")
            st.dataframe(pivot_sum)

            # グラフ作成
            count_melted = pivot_count.reset_index().melt(id_vars=selected_cols[0], var_name=selected_cols[1], value_name="件数")
            sum_melted = pivot_sum.reset_index().melt(id_vars=selected_cols[0], var_name=selected_cols[1], value_name="取扱高")
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=count_melted[selected_cols[0]] + "-" + count_melted[selected_cols[1]],
                y=count_melted["件数"],
                name="件数",
                marker_color="skyblue",
                offsetgroup=0,
                yaxis="y"
            ))
            fig.add_trace(go.Bar(
                x=sum_melted[selected_cols[0]] + "-" + sum_melted[selected_cols[1]],
                y=sum_melted["取扱高"],
                name="取扱高（円）",
                marker_color="orange",
                offsetgroup=1,
                yaxis="y2"
            ))
            fig.update_layout(
                title="クロス集計（件数＋取扱高）",
                xaxis=dict(title="組み合わせ"),
                yaxis=dict(title="件数", side="left"),
                yaxis2=dict(title="取扱高（円）", overlaying="y", side="right"),
                barmode="group"
            )
            st.plotly_chart(fig, use_container_width=True)

            # PowerPoint生成関数
            def generate_ppt_with_chart(filtered_df, selected_cols):
                prs = Presentation()
                slide1 = prs.slides.add_slide(prs.slide_layouts[1])
                slide1.shapes.title.text = "フィルタ後のデータ概要"
                slide1.placeholders[1].text = f"件数: {len(filtered_df)}\n取扱高合計: {filtered_df['取扱高'].sum():,.0f} 円"

                slide2 = prs.slides.add_slide(prs.slide_layouts[1])
                slide2.shapes.title.text = "クロス集計（件数）"
                slide2.placeholders[1].text = pivot_count.to_string()

                slide3 = prs.slides.add_slide(prs.slide_layouts[1])
                slide3.shapes.title.text = "クロス集計（取扱高）"
                slide3.placeholders[1].text = pivot_sum.to_string(float_format='{:,.0f}'.format)

                # グラフ画像保存
                fig.write_image("cross_chart.png")
                slide4 = prs.slides.add_slide(prs.slide_layouts[5])
                slide4.shapes.title.text = "クロス集計グラフ"
                slide4.shapes.add_picture("cross_chart.png", Inches(1), Inches(1.5), height=Inches(4.5))

                ppt_io = BytesIO()
                prs.save(ppt_io)
                ppt_io.seek(0)
                return ppt_io

            ppt_file = generate_ppt_with_chart(filtered_df, selected_cols)
            st.download_button("📤 PowerPointをダウンロード", ppt_file, file_name="filtered_summary.pptx")

else:
    st.info("Excelファイルをアップロードしてください。")