import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
import requests

st.set_page_config(page_title="後龍國中課表彙整系統", layout="wide")

# --- 1. 從 GitHub 抓取檔案的函數 ---
RAW_URL = "https://raw.githubusercontent.com/longdada124/autoclass/main/"

@st.cache_data(ttl=600)
def fetch_excel_from_github(filename):
    try:
        r = requests.get(RAW_URL + filename)
        r.raise_for_status()
        return r.content
    except Exception as e:
        st.error(f"無法讀取 {filename}: {e}")
        return None

# --- 2. 核心邏輯：讀取所有班級工作表 ---
def load_all_data():
    assign_data = fetch_excel_from_github("配課表.xlsx")
    table_data = fetch_excel_from_github("課表.xlsx")
    
    if not assign_data or not table_data:
        return None

    # 讀取 Excel 中所有的工作表
    # xls_a: 每個 key 是班級名稱，value 是該班級的配課 Dataframe
    xls_a = pd.read_excel(BytesIO(assign_data), sheet_name=None)
    xls_t = pd.read_excel(BytesIO(table_data), sheet_name=None)
    
    teacher_data = {}
    class_data = {}
    all_teachers = set()

    # 處理各班課表
    for class_name, df_t in xls_t.items():
        if class_name not in xls_a: continue # 若配課表沒這班就跳過
        
        df_a = xls_a[class_name].astype(str).apply(lambda x: x.str.strip())
        df_t = df_t.astype(str).apply(lambda x: x.str.strip())
        
        day_map = {"週一":1, "週二":2, "週三":3, "週四":4, "週五":5}
        class_data[class_name] = {}

        for _, row in df_t.iterrows():
            d_str = row['星期']
            p_val = row['節次']
            subj = row['科目']
            
            if d_str in day_map and str(p_val).isdigit():
                d, p = day_map[d_str], int(p_val)
                
                # 從該班配課頁面找出老師
                match = df_a[df_a['科目'] == subj]
                t_name = match.iloc[0]['教師'] if not match.empty else "未定"
                
                # 存入班級預覽資料
                class_data[class_name][(d, p)] = f"{subj}\n({t_name})"
                
                # 分解教師（處理如 葉麗君/張素梅）
                for t in [x.strip() for x in t_name.split('/')]:
                    if t == "未定": continue
                    all_teachers.add(t)
                    if t not in teacher_data: teacher_data[t] = {}
                    teacher_data[t][(d, p)] = {"subj": subj, "class": class_name}
                    
    return teacher_data, class_data, sorted(list(all_teachers)), sorted(list(class_data.keys()))

# --- 3. 執行加載 ---
data = load_all_data()

if data:
    t_db, c_db, teachers, classes = data
    
    tab1, tab2 = st.tabs(["🏫 班級課表預覽", "👨‍🏫 教師課表預覽"])
    
    with tab1:
        sel_c = st.selectbox("選擇班級", classes)
        df_c = pd.DataFrame(index=[f"第{i}節" for i in range(1, 9)], columns=["週一", "週二", "週三", "週四", "週五"])
        for d in range(1, 6):
            for p in range(1, 9):
                df_c.iloc[p-1, d-1] = c_db.get(sel_c, {}).get((d, p), "")
        st.table(df_c)
        
    with tab2:
        sel_t = st.selectbox("選擇教師", teachers)
        df_t = pd.DataFrame(index=[f"第{i}節" for i in range(1, 9)], columns=["週一", "週二", "週三", "週四", "週五"])
        for d in range(1, 6):
            for p in range(1, 9):
                v = t_db.get(sel_t, {}).get((d, p))
                df_t.iloc[p-1, d-1] = f"{v['class']}\n{v['subj']}" if v else ""
        st.table(df_t)
else:
    st.info("請確認 GitHub 上的 配課表.xlsx 與 課表.xlsx 是否已準備就緒。")
