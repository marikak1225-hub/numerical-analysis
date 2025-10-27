import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches, Pt
import io

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

        def group_loan(x):
            if pd.isna(x): return "不明"
            if x == 0: return "0"
            elif x < 10: return "1-9"
            elif x < 20: return "10-19"
            elif x < 30: return "20-29"
            elif x < 40: return "30-39"
            elif x < 50: return "40-49"
            elif x < 60: return "50-59"
            elif x < 70: return "60-69"
            elif x < 80: return "70-79"
            elif x < 90: return "80-89"
            elif x < 100: return "90-99"
            elif x < 200: return "100-199"
            elif x < 300: return "200-299"
            else: return "300以上"

        def group_mortgage(x):
            if pd.isna(x): return "不明"
            if x == 0: return "0"
            elif x < 10: return "1-9"
            elif x < 20: return "10-19"
            elif x < 30: return "20-29"
            elif x < 40: return "30-39"
            elif x < 50: return "40-49"
            elif x < 60: return "50-59"
            elif x < 70: return "60-69"
            elif x < 80: return "70-79"
            elif x < 90: return "80-89"
            elif x < 100: return "90-99"
            else: return "100以上"

        def group_years(x):
            if pd.isna(x): return "不明"
            if x == 0: return "0"
            elif x <= 3: return "1-3"
            elif x <= 9: return "4-9"
            elif x <= 20: return "10-20"
            else: return "21以上"
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
import io
import tempfile

st.set_page_config(page_title="後方数値データ分析", layout="wide")
st.title("📊 後方数値データ分析ダッシュボード")

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])
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

    if len(filtered_df) == 0:
        st.warning("データがありません。フィルタ条件を確認してください。")
    else:
        # カテゴリ変換関数
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

        def group_loan(x):
            if pd.isna(x): return "不明"
            if x == 0: return "0"
            elif x < 10: return "1-9"
            elif x < 20: return "10-19"
            elif x < 30: return "20-29"
            elif x < 40: return "30-39"
            elif x < 50: return "40-49"
            elif x < 60: return "50-59"
            elif x < 70: return "60-69"
            elif x < 80: return "70-79"
            elif x < 90: return "80-89"
            elif x < 100: return "90-99"
            elif x < 200: return "100-199"
            elif x < 300: return "200-299"
            else: return "300以上"

        def group_mortgage(x):
            if pd.isna(x): return "不明"
            if x == 0: return "0"
            elif x < 10: return "1-9"
            elif x < 20: return "10-19"
            elif x < 30: return "20-29"
            elif x < 40: return "30-39"
            elif x < 50: return "40-49"
            elif x < 60: return "50-59"
            elif x < 70: return "60-69"
            elif x < 80: return "70-79"
            elif x < 90: return "80-89"
            elif x < 100: return "90-99"
            else: return "100以上"

        def group_years(x):
            if pd.isna(x): return "不明"
            if x == 0: return "0"
            elif x <= 3: return "1-3"
            elif x <= 9: return "4-9"
            elif x <= 20: return "10-20"
            else: return "21以上"

        # カテゴリ変換
        filtered_df['年代'] = filtered_df['年齢'].apply(group_age)
        filtered_df['年収帯'] = filtered_df['年収'].apply(group_income)
        filtered_df['借入希望額帯'] = filtered_df['同借希望額'].apply(group_loan)
        filtered_df['住宅ローン帯'] = filtered_df['住宅ローン返済月額'].apply(group_mortgage)
        filtered_df['勤続年数帯'] = filtered_df['勤続年数'].apply(group_years)

        # グラフ作成関数
        def create_dual_axis_grouped_chart(df, category_col, title):
            count_data = df[category_col].value_counts().sort_index()
            sum_data = df.groupby(category_col)['取扱高'].sum().reindex(count_data.index)

            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=count_data.index,
                y=count_data.values,
                name="件数",
                marker_color="skyblue",
                text=[f"{v}" for v in count_data.values],
                textposition="outside",
                offsetgroup=0,
                yaxis="y"
            ))
            fig.add_trace(go.Bar(
                x=sum_data.index,
                y=sum_data.values,
                name="取扱高（円）",
                marker_color="orange",
                text=[f"{v/1_000_000:.1f}M" for v in sum_data.values],
                textposition="outside",
                offsetgroup=1,
                yaxis="y2"
            ))
            fig.update_layout(
                title=f"{title}の分布（件数＋取扱高）",
                xaxis=dict(title=category_col),
                yaxis=dict(title="件数", side="left"),
                yaxis2=dict(title="取扱高（円）", overlaying="y", side="right"),
                barmode="group"
            )
            return fig

        st.subheader("📈 項目別二軸横並びインタラクティブ棒グラフ")
        chart_cols = [
            ("性別", "性別"),
            ("年代別", "年代"),
            ("年収帯", "年収帯"),
            ("都道府県", "都道府県"),
            ("利用目的", "利用目的"),
            ("借入希望額帯", "借入希望額帯"),
            ("家族構成", "家族構成"),
            ("子供数", "子供数"),
            ("住宅ローン帯", "住宅ローン帯"),
            ("勤務状況", "勤務状況"),
            ("勤続年数帯", "勤続年数帯"),
            ("他社借入件数", "他社借入件数")
        ]

        figs = []
        for title, col in chart_cols:
            if col in filtered_df.columns and filtered_df[col].dropna().shape[0] > 0:
                fig = create_dual_axis_grouped_chart(filtered_df, col, title)
                st.plotly_chart(fig, use_container_width=True)
                figs.append((fig, title))

        # PDF作成関数
        def create_pdf(fig_list):
            buffer = io.BytesIO()
            c = canvas.Canvas(buffer, pagesize=A4)
            width, height = A4

            c.setFont("Helvetica-Bold", 16)
            c.drawString(50, height - 50, "後方数値データ分析レポート")
            c.setFont("Helvetica", 12)
            c.drawString(50, height - 80, f"期間: {start_date} ～ {end_date}")
            media_text = "ALL" if "ALL" in selected_codes else "媒体コード指定"
            c.drawString(50, height - 100, f"媒体コード: {media_text}")
            c.drawString(50, height - 120, f"件数: {len(filtered_df)}")
            c.showPage()

            for fig, title in fig_list:
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
                    fig.write_image(tmpfile.name, format="png")
                    c.drawImage(tmpfile.name, 50, 150, width=500, height=400)
                    c.setFont("Helvetica-Bold", 14)
                    c.drawString(50, height - 50, title)
                    c.showPage()

            c.save()
            buffer.seek(0)
            return buffer

        if figs:
            pdf_file = create_pdf(figs)
            st.download_button(
                label="📥 PDFレポートをダウンロード",
                data=pdf_file,
                file_name="後方数値データ分析レポート.pdf",
                mime="application/pdf"
            )

else:
    st.info("Excelファイルをアップロードしてください。")