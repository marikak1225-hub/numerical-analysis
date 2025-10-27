import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

st.set_page_config(page_title="後方数値データ分析", layout="wide")
st.title("📊 後方数値データ分析ダッシュボード")

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    df.columns = [str(c).strip().replace('　', '').replace(' ', '') for c in df.columns]

    if '申込日' in df.columns:
        df['申込日'] = pd.to_datetime(df['申込日'], errors='coerce')

    df['取扱高'] = df[[c for c in df.columns if '取扱金額' in c]].sum(axis=1)

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
        def create_chart(df, category_col, title):
            count_data = df[category_col].value_counts().sort_index()
            sum_data = df.groupby(category_col)['取扱高'].sum().reindex(count_data.index)
            fig = go.Figure()
            fig.add_trace(go.Bar(x=count_data.index, y=count_data.values, name="件数", marker_color="skyblue", yaxis="y"))
            fig.add_trace(go.Bar(x=sum_data.index, y=sum_data.values, name="取扱高（円）", marker_color="orange", yaxis="y2"))
            fig.update_layout(title=f"{title}の分布（件数＋取扱高）", xaxis=dict(title=category_col),
                              yaxis=dict(title="件数", side="left"),
                              yaxis2=dict(title="取扱高（円）", overlaying="y", side="right"),
                              barmode="group")
            return fig

        chart_cols = [("性別", "性別"), ("年代別", "年代"), ("都道府県", "都道府県")]
        figs = []
        for title, col in chart_cols:
            if col in filtered_df.columns:
                fig = create_chart(filtered_df, col, title)
                st.plotly_chart(fig, use_container_width=True)
                figs.append((fig, title, f"{title}の件数と取扱高の二軸グラフ"))

        # CSV出力
        if figs:
            csv_data = []
            for fig, title, desc in figs:
                for trace in fig.data:
                    csv_data.append(pd.DataFrame({
                        'カテゴリ': trace.x,
                        '値': trace.y,
                        '系列': trace.name,
                        'グラフタイトル': title
                    }))
            csv_combined = pd.concat(csv_data)
            csv_buffer = io.StringIO()
            csv_combined.to_csv(csv_buffer, index=False)
            st.download_button("📄 グラフデータをCSVでダウンロード", data=csv_buffer.getvalue(), file_name="グラフデータ.csv", mime="text/csv")

            # PDF出力
            pdf_buffer = io.BytesIO()
            c = canvas.Canvas(pdf_buffer, pagesize=A4)
            width, height = A4
            for fig, title, desc in figs:
                img_bytes = fig.to_image(format="png", scale=2)
                image = ImageReader(io.BytesIO(img_bytes))
                c.setFont("Helvetica-Bold", 16)
                c.drawString(40, height - 40, title)
                c.setFont("Helvetica", 12)
                c.drawString(40, height - 60, desc)
                c.drawImage(image, 40, 100, width=500, preserveAspectRatio=True, mask='auto')
                c.showPage()
            c.save()
            pdf_buffer.seek(0)
            st.download_button("📄 グラフをPDFでダウンロード", data=pdf_buffer, file_name="グラフレポート.pdf", mime="application/pdf")
else:
    st.info("Excelファイルをアップロードしてください。")
