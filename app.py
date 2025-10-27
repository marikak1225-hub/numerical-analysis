import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

st.set_page_config(page_title="後方数値データ分析", layout="wide")
st.title("📊 後方数値データ分析ダッシュボード")

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # 文字化け修正
    df.columns = [str(c).strip().replace('\u3000', '').replace('\xa0', '') for c in df.columns]

    # 数値変換
    numeric_cols = ['年齢', '年収', '同借希望額', '住宅ローン返済月額', '勤続年数', '他社借入件数',
                    '取扱金額_申込当月', '取扱金額_申込翌月末', '取扱金額_申込翌々月末']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # 日付変換
    if '申込日' in df.columns:
        df['申込日'] = pd.to_datetime(df['申込日'], errors='coerce')

    # 取扱高計算
    df['取扱高'] = df[['取扱金額_申込当月', '取扱金額_申込翌月末', '取扱金額_申込翌々月末']].sum(axis=1)

    # ✅ サイドバー：フィルタ設定
    st.sidebar.header("フィルタ設定")

    # 日付範囲フィルタ
    start_date, end_date = st.sidebar.date_input("申込日範囲", [df['申込日'].min(), df['申込日'].max()])

    # 媒体コードフィルタ
    media_codes = df['媒体コード'].dropna().unique().tolist() if '媒体コード' in df.columns else []
    selected_codes = st.sidebar.multiselect("媒体コードを選択（ALL選択で全件）", ["ALL"] + media_codes, default=["ALL"])

    # フィルタ適用
    filtered_df = df[(df['申込日'] >= pd.to_datetime(start_date)) & (df['申込日'] <= pd.to_datetime(end_date))]
    if "ALL" not in selected_codes:
        filtered_df = filtered_df[filtered_df['媒体コード'].isin(selected_codes)]

    st.write(f"件数: {len(filtered_df)}")

    if len(filtered_df) == 0:
        st.warning("データがありません。フィルタ条件を確認してください。")
    else:
        # ✅ グループ化関数
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

        # ✅ カテゴリ変換
        filtered_df['年代'] = filtered_df['年齢'].apply(group_age)
        filtered_df['年収帯'] = filtered_df['年収'].apply(group_income)

        # ✅ グラフ生成
        st.subheader("📈 項目別インタラクティブグラフ")
        chart_cols = [
            ("性別", "性別"),
            ("年代別", "年代"),
            ("年収帯", "年収帯"),
            ("都道府県", "都道府県")
        ]

        figs = []
        for title, col in chart_cols:
            if col in filtered_df.columns and filtered_df[col].dropna().shape[0] > 0:
                count_data = filtered_df[col].value_counts().sort_index()
                sum_data = filtered_df.groupby(col)['取扱高'].sum().reindex(count_data.index)
                fig = go.Figure()
                fig.add_trace(go.Bar(x=count_data.index, y=count_data.values, name="件数", marker_color="skyblue"))
                fig.add_trace(go.Bar(x=sum_data.index, y=sum_data.values, name="取扱高（円）", marker_color="orange"))
                fig.update_layout(title=f"{title}の分布", barmode="group")
                st.plotly_chart(fig, use_container_width=True)
                figs.append(fig)

        # ✅ CSV & PDF ダウンロード
        if figs:
            csv_buffer = io.StringIO()
            pd.DataFrame({'グラフ数': [len(figs)]}).to_csv(csv_buffer, index=False)
            st.download_button('📄 CSVでダウンロード', data=csv_buffer.getvalue(), file_name='graph_data.csv', mime='text/csv')

            pdf_buffer = io.BytesIO()
            c = canvas.Canvas(pdf_buffer, pagesize=A4)
            c.drawString(100, 800, 'グラフレポート')
            c.save()
            pdf_buffer.seek(0)
            st.download_button('📄 PDFでダウンロード', data=pdf_buffer, file_name='graph_report.pdf', mime='application/pdf')
else:
    st.info("Excelファイルをアップロードしてください。")