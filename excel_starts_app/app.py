import pandas as pd
import re
import streamlit as st
from datetime import time as dt_time


def get_class(name):
    if pd.isna(name):
        return "無班別"
    match = re.search(r'(\d+[A-Za-z])', str(name))
    if match:
        return match.group(1).upper()
    return "無班別"

def get_level(name):
    if pd.isna(name):
        return "無班別"
    match = re.search(r'(\d+)[A-Za-z]', str(name))
    if match:
        return f"中{match.group(1)}"
    return "無班別"

def time_to_sec(t):
    if isinstance(t, dt_time):
        return t.hour * 3600 + t.minute * 60 + t.second
    elif isinstance(t, str):
        try:
            h, m, s = map(int, t.split(":"))
            return h * 3600 + m * 60 + s
        except:
            return float('inf')
    return float('inf')

def load_data(df):
    df["級別"] = df["Player Name"].apply(get_level)
    df["班別"] = df["Player Name"].apply(get_class)
    df["時間_秒"] = df["Total Time Taken"].apply(time_to_sec)
    df["準確率_數值"] = df["Accuracy"].str.replace("%", "").astype(float)
    df = df[(df["時間_秒"] > 0) & (df["準確率_數值"] > 0)]
    return df

st.set_page_config(page_title="成績統計", layout="wide")
st.title("📊 Excel 數據統計網站")
st.subheader("上傳 Excel 檔案，自選工作表分析，彈性查詢級別/班別排名")


if "df" not in st.session_state:
    st.session_state.df = None
if "sheet_names" not in st.session_state:
    st.session_state.sheet_names = []


uploaded_file = st.file_uploader("選擇 Excel 檔", type=["xlsx", "xls"])


if uploaded_file is not None and not st.session_state.sheet_names:
    try:
        xl = pd.ExcelFile(uploaded_file)
        st.session_state.sheet_names = xl.sheet_names
        st.success("✅ 檔案讀取成功，請選擇要分析的工作表")
    except Exception as e:
        st.error(f"讀取 Excel 失敗：{e}")


selected_sheet = None
if st.session_state.sheet_names:
    selected_sheet = st.selectbox("請選擇要分析的工作表", st.session_state.sheet_names)


col1, col2 = st.columns(2)
with col1:
    x_level = st.number_input("各級查詢前 X 名", min_value=1, max_value=50, value=3, step=1)
with col2:
    y_class = st.number_input("各班查詢前 Y 名", min_value=1, max_value=50, value=3, step=1)


if uploaded_file is not None and selected_sheet and st.session_state.df is None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        df = load_data(df)
        st.session_state.df = df
        st.success(f"✅ 已載入工作表「{selected_sheet}」並完成分析！")
    except Exception as e:
        st.error(f"分析數據失敗：{e}")


if st.session_state.df is not None:
    df = st.session_state.df
    st.divider()

    #
    missing = df[df["班別"] == "無班別"]
    if not missing.empty:
        st.warning(f"⚠️ 偵測到 {len(missing)} 位同學沒有班別，請手動補充：")
        with st.expander("點此補充班別"):
            for idx, row in missing.iterrows():
                name = row["Player Name"]
                col_a, col_b = st.columns([3, 2])
                with col_a:
                    st.write(f"**{name}**")
                with col_b:
                    new_class = st.text_input(f"輸入班別（1A-6S）", key=f"class_{idx}")
                    if new_class:
                        new_class = new_class.strip().upper()
                        if re.match(r"\d+[A-Za-z]", new_class):
                            df.at[idx, "班別"] = new_class
                            df.at[idx, "級別"] = get_level(new_class)
                            st.session_state.df = df
                            st.success(f"✅ 已更新 {name} 為 {new_class}")
                        else:
                            st.error("❌ 班別格式錯誤，請輸入如 1A/1B/1C.. 的格式")
    else:
        st.info("✅ 所有同學都已抓到班別")

    st.divider()


    st.subheader(f"🏆 各級準確率最高前 {x_level} 位")
    normal = df[df["級別"] != "無班別"]
    for level, group in normal.groupby("級別"):
        with st.expander(f"{level} 級"):
            top_x = group.sort_values(by=["準確率_數值", "時間_秒"], ascending=[False, True]).head(x_level)
            for i, (_, stu) in enumerate(top_x.iterrows(), 1):
                st.write(f"第{i}名：{stu['Player Name']} | 準確率：{stu['Accuracy']} | 時間：{stu['Total Time Taken']}")

    st.divider()


    st.subheader(f"🏆 各班準確率最高前 {y_class} 位")
    class_normal = df[df["班別"] != "無班別"]
    for cls, group in class_normal.groupby("班別"):
        with st.expander(f"{cls} 班"):
            top_y = group.sort_values(by=["準確率_數值", "時間_秒"], ascending=[False, True]).head(y_class)
            for i, (_, stu) in enumerate(top_y.iterrows(), 1):
                st.write(f"第{i}名：{stu['Player Name']} | 準確率：{stu['Accuracy']} | 時間：{stu['Total Time Taken']}")

    st.divider()
    
    if st.button("匯出更新後的 Excel"):
        output = df.to_excel(index=False)
        st.download_button(
            label="下載更新後的檔案",
            data=output,
            file_name="updated_scores.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    if not uploaded_file:
        st.info("請上傳 Excel 檔案")
    elif uploaded_file and not selected_sheet:
        st.info("請選擇要分析的工作表")
