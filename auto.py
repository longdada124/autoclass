import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from io import BytesIO
import re
from datetime import datetime, timedelta

# --- 1. 核心字體與 Word 邏輯 ---
def set_font_style(run, font_name="標楷體"):
    """確保中文字體強制鎖定為標楷體"""
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def master_replace(doc_obj, old_text, new_text):
    """替換文字並套用標楷體"""
    new_val = str(new_text) if new_text is not None else ""
    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old_text in p.text:
                        for run in p.runs:
                            if old_text in run.text:
                                run.text = run.text.replace(old_text, new_val)
                                set_font_style(run)
                        if old_text in p.text:
                            p.text = p.text.replace(old_text, new_val)
                            for r in p.runs: set_font_style(r)

# --- 2. 模擬「拖曳」的視覺化介面 ---
def render_interactive_grid(teacher_name, t_db):
    """模仿 DM 建立可點擊的互動課表"""
    days = ["一", "二", "三", "四", "五"]
    st.write(f"### 📅 {teacher_name} 老師的週課表")
    st.caption("請點擊下方課表中的「藍色按鈕」來發動調代課")
    
    # 建立 8 節課的網格
    cols = st.columns(5)
    for d_idx, day in enumerate(days):
        with cols[d_idx]:
            st.button(day, disabled=True, use_container_width=True) # 標題
            for p in range(1, 9):
                info = t_db.get(teacher_name, {}).get((d_idx + 1, p))
                if info:
                    # 如果該節有課，顯示藍色按鈕
                    btn_label = f"第{p}節\n{info['c']}\n{info['s']}"
                    if st.button(btn_label, key=f"btn_{d_idx}_{p}", use_container_width=True, type="primary"):
                        st.session_state.selected_lesson = {
                            'day': d_idx + 1, 'period': p, 'c': info['c'], 's': info['s']
                        }
                else:
                    # 無課則顯示空白按鈕
                    st.button(f"第{p}節", key=f"empty_{d_idx}_{p}", disabled=True, use_container_width=True)

# --- 3. 主程式架構 ---
st.set_page_config(page_title="後龍國中智慧代調課系統", layout="wide")

# 初始化 session state
if 'selected_lesson' not in st.session_state: st.session_state.selected_lesson = None

with st.sidebar:
    st.header("📂 數據與樣板")
    f_assign = st.file_uploader("1. 上傳配課表", type=["xlsx"])
    f_time = st.file_uploader("2. 上傳課表", type=["xlsx"])
    f_temp = st.file_uploader("3. 上傳 Word 樣板", type=["docx"])
    
    if f_assign and f_time and f_temp:
        if st.button("🔄 執行整合"):
            # (資料處理邏輯同前，簡化展示)
            df_a = pd.read_excel(f_assign).astype(str).apply(lambda x: x.str.strip())
            df_t = pd.read_excel(f_time).astype(str).apply(lambda x: x.str.strip())
            t_db = {}
            all_t = set()
            day_map = {"一":1, "二":2, "三":3, "四":4, "五":5}
            for _, r in df_t.iterrows():
                d_match = re.search(r'[一二三四五]', r['星期'])
                p_match = re.search(r'\d+', r['節次'])
                if d_match and p_match:
                    d, p = day_map[d_match.group()], int(p_match.group())
                    c, s = r['班級'], r['科目']
                    match = df_a[(df_a['班級'] == c) & (df_a['科目'] == s)]
                    ts = str(match.iloc[0]['教師']).split('/') if not match.empty else ["未知"]
                    for t in [x.strip() for x in ts]:
                        all_t.add(t)
                        if t not in t_db: t_db[t] = {}
                        t_db[t][(d, p)] = {"c": c, "s": s}
            st.session_state.update({"t_db": t_db, "all_t": sorted(list(all_t)), "template": f_temp.read(), "ready": True})

if st.session_state.get("ready"):
    st.title("📑 智慧代調課管理系統 (互動版)")
    
    t_list = st.session_state.all_t
    sel_teacher = st.selectbox("🔍 請選擇請假教師", t_list)
    
    # 顯示互動課表
    render_interactive_grid(sel_teacher, st.session_state.t_db)
    
    # 如果使用者點擊了某節課
    if st.session_state.selected_lesson:
        l = st.session_state.selected_lesson
        st.divider()
        st.success(f"📍 已選取：週{l['day']} 第{l['period']}節 - {l['c']} {l['s']}")
        
        col_a, col_b = st.columns(2)
        with col_a:
            date_sel = st.date_input("🗓️ 實際變動日期", datetime.now())
            mode = st.radio("🔄 變動性質", ["代課", "調課"], horizontal=True)
        
        with col_b:
            # 智慧過濾衝堂 (DM 功能)
            avail_ts = [t for t in t_list if (l['day'], l['period']) not in st.session_state.t_db.get(t, {})]
            to_t = st.selectbox("👤 接收教師 (系統已過濾衝堂者)", avail_ts)
        
        if st.button("🚀 確認並生成通知單"):
            # 計算日期
            mon = date_sel - timedelta(days=date_sel.weekday())
            w_strs = [f"{mon.year-1911}.{(mon+timedelta(days=i)).month:02d}.{(mon+timedelta(days=i)).day:02d}" for i in range(5)]
            
            # 生成檔案
            doc = Document(BytesIO(st.session_state.template))
            master_replace(doc, "{{TEACHER}}", to_t)
            for i, d_str in enumerate(w_strs): master_replace(doc, f"{{{{D{i+1}}}}}", d_str)
            
            # 填寫內容並清理所有格子
            target_tag = f"{{{{{l['day']}_{l['period']}}}}}"
            content = f"{mode[:1]}{l['c']}\n{l['s']}"
            for d in range(1, 6):
                for p in range(1, 9):
                    tag = f"{{{{{d}_{p}}}}}"
                    master_replace(doc, tag, content if tag == target_tag else "")
            
            buf = BytesIO()
            doc.save(buf)
            st.download_button(f"⬇️ 下載 {to_t} 老師的通知單", buf.getvalue(), f"{to_t}_通知單.docx")
