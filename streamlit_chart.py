import os
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from streamlit_autorefresh import st_autorefresh

st.set_page_config(page_title="Rain Odds & Chart", layout="wide")
st.title("Real-time Rain Odds and Chart")

# 文件上传（替换本地路径）
uploaded_file = st.sidebar.file_uploader("上传 Excel 文件 (如 JUL28_06-18.xlsx)", type=["xlsx"])

# 自动每 N 秒刷新（Sidebar可调整）
refresh_interval = st.sidebar.slider("Refresh interval (seconds)", 2, 30, 2)
st_autorefresh(interval=refresh_interval*1000, key="datarefresh")

def read_data(uploaded_file):
    if uploaded_file is None:
        st.warning("请上传 Excel 文件以查看数据。")
        return pd.DataFrame()
    
    try:
        wb = load_workbook(uploaded_file, data_only=True)
        ws = wb.active
        data = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is not None:
                data.append({
                    "Rain Odds": float(row[0]),
                    "50MA": float(row[1]) if row[1] is not None else None,
                    "150MA": float(row[2]) if row[2] is not None else None,
                    "Time": row[3]
                })
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"发生错误：{e}")
        return pd.DataFrame()

df = read_data(uploaded_file)

if not df.empty:
    st.line_chart(df[["Rain Odds", "50MA", "150MA"]])
    st.dataframe(df)
    st.info(f"資料最新時間：{df['Time'].iloc[-1]}")