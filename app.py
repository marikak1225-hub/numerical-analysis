import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
from PIL import Image
import openpyxl
from openpyxl.drawing.image import Image as XLImage

# ページ設定
st.set_page_config(page_title="後方数値データ分析", layout="wide")
st.title("📊 後方数値データ分析ダッシュボード")

# カテゴリ順序定義
category_orders = {
    "年収帯": ['0-499', '500-999', '1000以上'],
    "借入希望額帯": ['0', '1-9', '10-19', '20-29', '30-39', '40-49', '50-59', '60-69', '70-79', '80-89', '90-99', '100-199', '200-299', '300以上'],
    "住宅ローン帯": ['0', '1-9', '10-19', '20-29', '30-39', '40-49', '50-59', '60-69', '70-79', '80-89', '90-99', '100以上'],
    "勤続年数帯": ['0', '1-3', '4-9', '10-20', '21以上']
}

# サイドバー：ファイルアップロード
st.sidebar.header("ファイルアップロード")
uploaded_data = st.sidebar.file_uploader("後方数値データをアップロード", type=["xlsx"])

# マスタファイル読み込み（GitHub固定）
master_path = "媒体コードマスタ.xlsx"
master = pd.read_excel(master_path)

# 列名正規化
master.columns = [str(c).strip().replace('\u3000', '').replace('\xa0', '') for c in master.columns]

# 「会社名」を「媒体名」に変更
master.rename(columns={"会社名": "媒体名"}, inplace=True)

# id_varsとコード列を動的に取得
id_vars = [col for col in master.columns if col in ["媒体名", "カテゴリ"]]
code_cols = [col for col in master.columns if col not in id_vars]

# 縦持ち変換
master_long = master.melt(id_vars=id_vars, value_vars=code_cols,
                          var_name="コード列", value_name="媒体コード").dropna(subset=["媒体コード"])

if uploaded_data:
    # 後方数値データ読み込み
    df = pd.read_excel(uploaded_data)
    df.columns = [str(c).strip().replace('\u3000', '').replace('\xa0', '') for c in df.columns]

    # 性別整形
    if '性別' in df.columns:
        df['性別'] = df['性別'].astype(str).str.extract(r'_(男性|女性)')

    # 数値列変換
    numeric_cols = ['年齢', '年収', '同借希望額', '住宅ローン返済月額', '勤続年数',
                    '他社借入件数', '取扱金額_申込当月', '取扱金額_申込翌月末', '取扱金額_申込翌々月末']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    if '申込日' in df.columns:
        df['申込日'] = pd.to_datetime(df['申込日'], errors='coerce')

    # 取扱高計算
    df['取扱高'] = df[['取扱金額_申込当月', '取扱金額_申込翌月末', '取扱金額_申込翌々月末']].sum(axis=1)

    # 承認区分のNULL処理
    if '承認区分' in df.columns:
        df['承認区分'] = df['承認区分'].fillna('NULL')
    else:
        df['承認区分'] = 'NULL'

    # マスタと突合
    merged_df = df.merge(master_long, on="媒体コード", how="left")

    # フィルタUI
    st.sidebar.header("フィルタ設定")
    start_date, end_date = st.sidebar.date_input("申込日範囲", [merged_df['申込日'].min(), merged_df['申込日'].max()])
    gender_options = ["ALL", "男性", "女性"]
    selected_genders = st.sidebar.multiselect("性別を選択", gender_options, default=["ALL"])

    company_options = ["ALL"] + (merged_df["媒体名"].dropna().unique().tolist() if "媒体名" in merged_df.columns else [])
    selected_companies = st.sidebar.multiselect("媒体名を選択", company_options, default=["ALL"])

    category_options = ["ALL"] + (merged_df["カテゴリ"].dropna().unique().tolist() if "カテゴリ" in merged_df.columns else [])
    selected_categories = st.sidebar.multiselect("カテゴリを選択", category_options, default=["ALL"])

    approval_options = ["ALL", "承認", "スモール", "NULL"]
    selected_approval = st.sidebar.multiselect("承認区分を選択", approval_options, default=["ALL"])

    # フィルタ処理
    filtered_df = merged_df[(merged_df['申込日'] >= pd.to_datetime(start_date)) & (merged_df['申込日'] <= pd.to_datetime(end_date))]
    if "ALL" not in selected_genders and '性別' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['性別'].isin(selected_genders)]
    if "ALL" not in selected_companies and "媒体名" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["媒体名"].isin(selected_companies)]
    if "ALL" not in selected_categories and "カテゴリ" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["カテゴリ"].isin(selected_categories)]
    if "ALL" not in selected_approval and "承認区分" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["承認区分"].isin(selected_approval)]

    st.write(f"件数: {len(filtered_df)}")

    # 年齢を10刻みでグループ化
    def group_age_10(x):
        if pd.isna(x): return "不明"
        try:
            x = int(x)
        except:
            return "不明"
        if x < 10: return "0-9"
        elif x < 20: return "10-19"
        elif x < 30: return "20-29"
        elif x < 40: return "30-39"
        elif x < 50: return "40-49"
        elif x < 60: return "50-59"
        elif x < 70: return "60-69"
        elif x < 80: return "70-79"
        elif x < 90: return "80-89"
        else: return "90以上"

    filtered_df['年齢'] = filtered_df['年齢'].apply(group_age_10)

    # 年収帯・借入希望額帯・住宅ローン帯・勤続年数帯も分類
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

    filtered_df['年収帯'] = filtered_df['年収'].apply(group_income)
    filtered_df['借入希望額帯'] = filtered_df['同借希望額'].apply(group_loan)
    filtered_df['住宅ローン帯'] = filtered_df['住宅ローン返済月額'].apply(group_mortgage)
    filtered_df['勤続年数帯'] = filtered_df['勤続年数'].apply(group_years)

    # グラフ表示
    st.subheader("📈 項目別インタラクティブグラフ")
    chart_cols = [
        ("性別", "性別"),
        ("年齢", "年齢"),
        ("年収", "年収帯"),
        ("都道府県", "都道府県"),
        ("利用目的", "利用目的"),
        ("同借希望額", "借入希望額帯"),
        ("家族構成", "家族構成"),
        ("子供数", "子供数"),
        ("住宅ローン返済月額", "住宅ローン帯"),
        ("勤務状況", "勤務状況"),
        ("勤続年数", "勤続年数帯"),
        ("他社借入件数", "他社借入件数"),
        ("媒体名", "媒体名"),
        ("承認区分", "承認区分")
    ]

    def create_dual_axis_grouped_chart(df, category_col, title):
        if category_col not in df.columns or df[category_col].dropna().shape[0] == 0:
            return go.Figure()
        if category_col in category_orders:
            ordered_categories = category_orders[category_col]
            count_data = df[category_col].value_counts().reindex(ordered_categories).fillna(0)
            sum_data = df.groupby(category_col)['取扱高'].sum().reindex(ordered_categories).fillna(0)
        else:
            count_data = df[category_col].value_counts().sort_index()
            sum_data = df.groupby(category_col)['取扱高'].sum().reindex(count_data.index)

        fig = go.Figure()
        fig.add_trace(go.Bar(x=count_data.index, y=count_data.values, name="件数", marker_color="skyblue"))
        fig.add_trace(go.Bar(x=sum_data.index, y=sum_data.values, name="取扱高（円）", marker_color="orange", yaxis="y2"))
        fig.update_layout(title=f"{title}（件数＋取扱高）", barmode="group", yaxis=dict(title="件数"), yaxis2=dict(title="取扱高（円）", overlaying="y", side="right"))
        return fig

    # グラフ生成とExcel貼り付け用リスト
    figs = []
    for title, col in chart_cols:
        if col in filtered_df.columns and filtered_df[col].dropna().shape[0] > 0:
            fig = create_dual_axis_grouped_chart(filtered_df, col, title)
            st.plotly_chart(fig, use_container_width=True)
            figs.append((title, fig))

    # Excelに画像貼り付け
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "グラフ一覧"
    row = 1
    for title, fig in figs:
        img_bytes = fig.to_image(format="png")
        img = Image.open(io.BytesIO(img_bytes))
        img_path = f"{title}.png"
        img.save(img_path)
        xl_img = XLImage(img_path)
        ws.add_image(xl_img, f"A{row}")
        row += 20  # 次の画像の位置をずらす

    excel_bytes = io.BytesIO()
    wb.save(excel_bytes)

    st.download_button("📥 グラフをExcelでダウンロード", data=excel_bytes.getvalue(), file_name="charts.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("Excelファイルをアップロードしてください。")