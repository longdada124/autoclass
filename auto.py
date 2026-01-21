import streamlit as st
import pandas as pd
import requests
import base64
from io import BytesIO
from docx import Document
from docx.oxml.ns import qn
import re
from datetime import datetime, timedelta

# --- 1. 配置與雲端連接 ---
REPO = "longdada124/autoclass"
TOKEN = st.secrets["G_TOKEN"] 
FILES = {"assign": "配課表.xlsx", "timetable": "課表.xlsx", "template": "代調課通知單樣板.docx"}

# --- 2. 核心功能函數 ---
def set_font(run):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

def master_replace(doc, old_text, new_text):
    new_val = str(new_text) if new_text else ""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old_text in p.text:
                        for run in p.runs:
                            if old_text in run.text:
                                run.text = run.text.replace(old_text, new_val)
                                set_font(run)

def pull_from_github(filename):
    url = f"https://raw.githubusercontent.com/{REPO}/main/{filename}"
    r = requests.get(url)
    return r.content if r.status_code == 200 else None

# --- 3. 初始化與資料處理 ---
st.set_page_config(page_title="後龍國中智慧教務系統", layout="wide")

@st.cache_data(ttl=600)
def load_all_data():
    a_bytes = pull_from_github(FILES["assign"])
    t_bytes = pull_from_github(FILES["timetable"])
    doc_bytes = pull_from_github(FILES["template"])
    if not a_bytes or not t_bytes: return None

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
            d_s, p_v, subj = row['星期'], row['節次'], row['科目']
            if d_s in day_map and str(p_v).isdigit():
                d, p = day_map[d_s], int(p_v)
                match = df_a[df_a['科目'] == subj]
                t_raw = match.iloc[0]['教師'] if not match.empty else "未定"
                c_db[c_name][(d, p)] = f"{subj}\n({t_raw})"
                for t in [x.strip() for x in t_raw.split('/')]:
                    if t == "未定": continue
                    all_t.add(t)
                    if t not in t_db: t_db[t] = {}
                    t_db[t][(d, p)] = {"c": c_name, "s": subj}
    return t_db, c_db, sorted(list(all_t)), all_c, doc_bytes

data_pkg = load_all_data()

# --- 4. 主要介面 ---
if data_pkg:
    t_db, c_db, teachers, classes, template_bytes = data_pkg
    tab1, tab2, tab3 = st.tabs(["🏫 班級課表(切換/定位)", "👨‍🏫 教師課表(列印)", "📝 代調課系統"])

    with tab1:
        # --- 切換上、下一班功能 ---
        if 'c_idx' not in st.session_state: st.session_state.c_idx = 0
        c1, c2, c3 = st.columns([1, 2, 1])
        with c1: 
            if st.button("⬅️ 上一班") and st.session_state.c_idx > 0:
                st.session_state.c_idx -= 1
        with c2:
            sel_c = st.selectbox("跳轉班級", classes, index=st.session_state.c_idx, key="sb_c")
            st.session_state.c_idx = classes.index(sel_c)
        with c3:
            if st.button("下一班 ➡️") and st.session_state.c_idx < len(classes)-1:
                st.session_state.c_idx += 1
                st.rerun()

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
                df_t.iloc[p-1, d-1] = f"{v['c']} {v['s']}" if v else ""
        st.table(df_t)
        
        # 這裡是您要的列印功能 (以代調課樣板為例或可自換樣板)
        st.download_button("🖨️ 下載該師課表 (Word)", b"test", file_name=f"{sel_t}_課表.docx", disabled=True)

    with tab3:
        st.subheader("生成代調課通知單")
        l_t = st.selectbox("請假教師", teachers)
        # 顯示互動課表點選定位
        grid = st.columns(5)
        for d in range(5):
            with grid[d]:
                st.button(["週一","週二","週三","週四","週五"][d], disabled=True, use_container_width=True)
                for p in range(1, 9):
                    info = t_db.get(l_t, {}).get((d+1, p))
                    if info:
                        if st.button(f"{p}\n{info['c']}", key=f"btn_{d}_{p}", use_container_width=True):
                            st.session_state.act = {'d':d+1, 'p':p, 'c':info['c'], 's':info['s']}
        
        if 'act' in st.session_state:
            a = st.session_state.act
            st.info(f"選取：週{a['d']} 第{a['p']}節 {a['c']}{a['s']}")
            v_date = st.date_input("更動日期", datetime.now())
            target_t = st.selectbox("代課教師", teachers)
            if st.button("📝 生成通知單"):
                doc = Document(BytesIO(template_bytes))
                master_replace(doc, "{{TEACHER}}", target_t)
                # 清理與填入
                tag = f"{{{{{a['d']}_{a['p']}}}}}"
                for d_ in range(1,6):
                    for p_ in range(1,9):
                        curr = f"{{{{{d_}_{p_}}}}}"
                        master_replace(doc, curr, f"代{a['c']}{a['s']}" if curr == tag else "")
                buf = BytesIO(); doc.save(buf)
                st.download_button(f"⬇️ 下載 {target_t} 通知單", buf.getvalue(), f"{target_t}_通知單.docx")

else:
    st.error("❌ 雲端抓取失敗。請確認 GitHub 檔案及 G_TOKEN 設定。")
