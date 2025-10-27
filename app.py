import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import fitz  # PyMuPDF
from PIL import Image

st.set_page_config(page_title="後方数値データ分析", layout="wide")
st.title("📊 後方数値データ分析ダッシュボード")

# ✅ ファイルアップロード
uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # ✅ 列名整形
    df.columns = [str(c).strip().replace('\u3000', '').replace('\xa0', '') for c in df.columns]

    # ✅ 日付変換
    if '申込日' in df.columns:
        df['申込日'] = pd.to_datetime(df['申込日'], errors='coerce')

    # ✅ 取扱高の定義（申込当月＋翌月末＋翌々月末）
    amount_cols = ['取扱金額_申込当月', '取扱金額_申込翌月末', '取扱金額_申込翌々月末']
    missing_cols = [col for col in amount_cols if col not in df.columns]
    if missing_cols:
        st.error(f"以下の列が不足しています: {', '.join(missing_cols)}")
    else:
        df['取扱高'] = df[amount_cols].sum(axis=1)

    # ✅ サイドバー：フィルタ設定
    st.sidebar.header("フィルタ設定")
    start_date, end_date = st.sidebar.date_input("申込日範囲", [df['申込日'].min(), df['申込日'].max()])
    media_codes = df['媒体コード'].dropna().unique().tolist() if '媒体コード' in df.columns else []
    selected_codes = st.sidebar.multiselect("媒体コードを選択（ALL選択で全件）", ["ALL"] + media_codes, default=["ALL"])

    # ✅ フィルタ適用
    filtered_df = df[(df['申込日'] >= pd.to_datetime(start_date)) & (df['申込日'] <= pd.to_datetime(end_date))]
    if "ALL" not in selected_codes:
        filtered_df = filtered_df[filtered_df['媒体コード'].isin(selected_codes)]

    st.write(f"件数: {len(filtered_df)}")

    if len(filtered_df) == 0:
        st.warning("データがありません。フィルタ条件を確認してください。")
    else:
        # ✅ グラフ作成
        st.subheader("📈 項目別インタラクティブグラフ")
        figs = []

        def create_dual_axis_chart(df, category_col, title):
            count_data = df[category_col].value_counts().sort_index()
            sum_data = df.groupby(category_col)['取扱高'].sum().reindex(count_data.index)
            fig = go.Figure()
            fig.add_trace(go.Bar(x=count_data.index, y=count_data.values, name="件数", marker_color="skyblue", offsetgroup=0))
            fig.add_trace(go.Bar(x=sum_data.index, y=sum_data.values, name="取扱高（円）", marker_color="orange", offsetgroup=1, yaxis="y2"))
            fig.update_layout(title=f"{title}の分布", yaxis=dict(title="件数"), yaxis2=dict(title="取扱高（円）", overlaying="y", side="right"), barmode="group")
            return fig

        chart_cols = [("性別", "性別"), ("年代別", "年代"), ("年収帯", "年収帯")]
        for title, col in chart_cols:
            if col in filtered_df.columns and '取扱高' in filtered_df.columns:
                fig = create_dual_axis_chart(filtered_df, col, title)
                st.plotly_chart(fig, use_container_width=True)
                figs.append((fig, title))

        # ✅ クロス集計
        st.subheader("🔍 クロス集計")
        selected_cols = st.multiselect("クロス集計する項目を選択", [c for _, c in chart_cols])
        if len(selected_cols) >= 2:
            pivot_count = pd.pivot_table(filtered_df, index=selected_cols[0], columns=selected_cols[1], aggfunc='size', fill_value=0)
            pivot_sum = pd.pivot_table(filtered_df, index=selected_cols[0], columns=selected_cols[1], values='取扱高', aggfunc='sum', fill_value=0)
            st.write("件数")
            st.dataframe(pivot_count)
            st.write("取扱高（円）")
            st.dataframe(pivot_sum)

        # ✅ CSV出力
        csv_buffer = io.StringIO()
        filtered_df.to_csv(csv_buffer, index=False)
        st.download_button("📄 データをCSVでダウンロード", data=csv_buffer.getvalue(), file_name="filtered_data.csv", mime="text/csv")

        # ✅ PDF出力（グラフ＋クロス集計）
        if figs:
            pdf_buffer = io.BytesIO()
            doc = fitz.open()
            title_page = doc.new_page()
            title_text = f"後方数値データ分析レポート\n\n期間: {start_date} ～ {end_date}\n媒体コード: {'ALL' if 'ALL' in selected_codes else '媒体コード指定'}\n件数: {len(filtered_df)}"
            title_page.insert_text((72, 72), title_text, fontsize=14)

            for fig, title in figs:
                img_bytes = fig.to_image(format="png", scale=2)
                rect = fitz.Rect(50, 50, 550, 550)
                page = doc.new_page()
                page.insert_text((72, 30), title, fontsize=16)
                page.insert_image(rect, stream=img_bytes)

            doc.save(pdf_buffer)
            pdf_buffer.seek(0)
            st.download_button("📄 レポートをPDFでダウンロード", data=pdf_buffer, file_name="analysis_report.pdf", mime="application/pdf")

else:
    st.info("Excelファイルをアップロードしてください。")