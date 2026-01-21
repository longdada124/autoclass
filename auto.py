import streamlit as st
import pandas as pd
import requests
import base64
from io import BytesIO
from docx import Document
import re

# --- 1. 設定與 GitHub 連接 ---
REPO = "longdada124/autoclass"
TOKEN = st.secrets["G_TOKEN"]  # 請確保已在 Streamlit Secrets 設定此變數
FILES = {
    "assign": "配課表.xlsx",
    "timetable": "課表.xlsx"
}

def push_to_github(content, filename):
    url = f"https://api.github.com/repos/{REPO}/contents/{filename}"
    headers = {"Authorization": f"token {TOKEN}"}
    r = requests.get(url, headers=headers)
    sha = r.json().get("sha") if r.status_code == 200 else None
    encoded = base64.b64encode(content).decode("utf-8")
    data = {"message": f"Web Update {filename}", "content": encoded, "branch": "main"}
    if sha: data["sha"] = sha
    res = requests.put(url, headers=headers, json=data)
    return res.status_code in [200, 201]

def pull_from_github(filename):
    url = f"https://raw.githubusercontent.com/{REPO}/main/{filename}"
    r = requests.get(url)
    return r.content if r.status_code == 200 else None

# --- 2. 頁面配置 ---
st.set_page_config(page_title="後龍國中課表雲端系統", layout="wide")

# --- 3. 側邊欄：僅在需要更新時使用 ---
with st.sidebar:
    st.header("⚙️ 雲端資料更新")
    st.info("上傳後點擊同步，資料將永久儲存於 GitHub。")
    up_a = st.file_uploader("1. 更新配課表 (Excel)", type="xlsx")
    up_t = st.file_uploader("2. 更新全校課表 (Excel)", type="xlsx")
    
    if st.button("🚀 同步並儲存至雲端"):
        with st.spinner("同步中..."):
            if up_a: push_to_github(up_a.getvalue(), FILES["assign"])
            if up_t: push_to_github(up_t.getvalue(), FILES["timetable"])
        st.success("✅ 同步成功！下次開啟不需再上傳。")
        st.rerun()

# --- 4. 資料讀取與解析邏輯 ---
@st.cache_data(ttl=600)
def load_system_data():
    a_bytes = pull_from_github(FILES["assign"])
    t_bytes = pull_from_github(FILES["timetable"])
    
    if not a_bytes or not t_bytes:
        return None

    xls_a = pd.read_excel(BytesIO(a_bytes), sheet_name=None)
    xls_t = pd.read_excel(BytesIO(t_bytes), sheet_name=None)
    
    t_db, c_db = {}, {}
    all_t, all_c = set(), sorted(list(xls_t.keys()))
    day_map = {"週一":1, "週二":2, "週三":3, "週四":4, "週五":5}

    for c_name in all_c:
        if c_name not in xls_a: continue
        df_a = xls_a[c_name].astype(str).apply(lambda x: x.str.strip())
        df_t = xls_t[c_name].astype(str).apply(lambda x: x.str.strip())
        c_db[c_name] = {}
        
        for _, row in df_t.iterrows():
            d_str, p_val, subj = row['星期'], row['節次'], row['科目']
            if d_str in day_map and str(p_val).isdigit():
                d, p = day_map[d_str], int(p_val)
                # 對應配課老師
                match = df_a[df_a['科目'] == subj]
                t_raw = match.iloc[0]['教師'] if not match.empty else "未定"
                c_db[c_name][(d, p)] = f"{subj}\n({t_raw})"
                
                # 建立教師索引
                for t in [x.strip() for x in t_raw.split('/')]:
                    if t == "未定": continue
                    all_t.add(t)
                    if t not in t_db: t_db[t] = {}
                    t_db[t][(d, p)] = {"c": c_name, "s": subj}
                    
    return t_db, c_db, sorted(list(all_t)), all_c

# --- 5. 主介面顯示 ---
data_package = load_system_data()

if data_package:
    t_db, c_db, teachers, classes = data_package
    tab1, tab2 = st.tabs(["🏫 班級課表預覽 (一班一頁格式)", "👨‍🏫 教師個人課表"])

    with tab1:
        sel_c = st.selectbox("請選擇班級", classes)
        view_c = pd.DataFrame(index=[f"第{i}節" for i in range(1, 9)], columns=["週一", "週二", "週三", "週四", "週五"])
        for d in range(1, 6):
            for p in range(1, 9):
                view_c.iloc[p-1, d-1] = c_db.get(sel_c, {}).get((d, p), "")
        st.table(view_c)

    with tab2:
        sel_t = st.selectbox("請選擇教師", teachers)
        view_t = pd.DataFrame(index=[f"第{i}節" for i in range(1, 9)], columns=["週一", "週二", "週三", "週四", "週五"])
        for d in range(1, 6):
            for p in range(1, 9):
                item = t_db.get(sel_t, {}).get((d, p))
                view_t.iloc[p-1, d-1] = f"{item['c']}\n{item['s']}" if item else ""
        st.table(view_t)
        
        # 額外功能：如果需要 Word 輸出可在這裡加入之前給您的 master_replace 邏輯
else:
    st.warning("👋 歡迎使用！偵測到雲端尚無資料，請先在左側上傳 Excel 檔案並點擊同步。")
    st.image("https://img.icons8.com/clouds/200/database.png")
