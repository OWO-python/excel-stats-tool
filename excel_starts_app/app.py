import streamlit as st
import pandas as pd
from io import BytesIO

# ----------------------
# 頁面標題
# ----------------------
st.title("📊 Excel 數據統計與分析工具")
st.write("上傳你的 Excel 檔案，進行處理後下載更新版本")

# ----------------------
# 上傳 Excel
# ----------------------
uploaded_file = st.file_uploader("選擇一個 Excel 檔案", type=["xlsx", "xls"])

if uploaded_file is not None:
    # 讀取 Excel
    df = pd.read_excel(uploaded_file)
    st.success("✅ 成功讀取 Excel 檔案")

    # ----------------------
    # 顯示原始數據
    # ----------------------
    st.subheader("📋 原始數據預覽")
    st.dataframe(df.head(10))

    # ----------------------
    # 你的數據處理邏輯（這裡隨意修改）
    # 範例：新增一個總和欄位
    # ----------------------
    st.subheader("🔧 數據處理")
    st.write("目前範例：新增一個「備註」欄位（你可自行改邏輯）")

    # 簡單處理範例
    df["備註"] = "已處理"  # 新增欄位

    st.dataframe(df.head(10))

    # ----------------------
    # ✅ 正確下載 Excel（不會報錯）
    # ----------------------
    st.subheader("📥 下載更新後的 Excel")

    # 創建內存緩存
    buffer = BytesIO()

    # 寫入 Excel
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    # 提供下載按鈕
    st.download_button(
        label="📌 點擊下載更新後的 Excel",
        data=buffer.getvalue(),
        file_name="更新後的數據.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("ℹ️ 請上傳一個 Excel 檔案開始使用")
