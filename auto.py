import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
from datetime import datetime, timedelta

# --- 系統配置 ---
st.set_page_config(page_title="後龍國中智慧代調課系統", layout="wide")

# --- 核心邏輯：Word 處理 (保留字體與換行) ---
def master_replace(doc_obj, old_text, new_text):
    new_val = str(new_text) if new_text is not None else ""
    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old_text in p.text:
                        for run in p.runs:
                            if old_text in run.text:
                                if "\n" in new_val:
                                    parts = new_val.split("\n")
                                    run.text = run.text.replace(old_text, parts[0])
                                    for part in parts[1:]:
                                        run.add_break()
                                        run.add_text(part)
                                else:
                                    run.text = run.text.replace(old_text, new_val)

def generate_sub_notice(template_bytes, target_teacher, change_data, week_dates):
    doc = Document(BytesIO(template_bytes))
    # 填寫抬頭與日期 
    master_replace(doc, "{{TEACHER}}", target_teacher)
    for i, d_str in enumerate(week_dates):
        master_replace(doc, f"{{{{D{i+1}}}}}", d_str)
    
    # 填寫目標格子，其餘清空 
    target_tag = f"{{{{{change_data['day']}_{change_data['period']}}}}}"
    for d in range(1, 6):
        for p in range(1, 9):
            tag = f"{{{{{d}_{p}}}}}"
            if tag == target_tag:
                master_replace(doc, tag, change_data['content'])
            else:
                master_replace(doc, tag, "")
    return doc

# --- UI 輔助函數 ---
def get_roc_week(base_date):
    start = base_date - timedelta(days=base_date.weekday())
    return [f"{d.year-1911}.{d.month:02d}.{d.day:02d}" for d in [start + timedelta(days=i) for i in range(5)]]

# --- 側邊欄：資料中心 ---
with st.sidebar:
    st.header("📂 數據與樣板管理")
    f_assign = st.file_uploader("1. 上傳配課表 (Excel)", type=["xlsx"])
    f_time = st.file_uploader("2. 上傳課表 (Excel)", type=["xlsx"])
    f_temp = st.file_uploader("3. 上傳代調課樣板 (.docx)", type=["docx"])
    
    if f_assign and f_time and f_temp:
        if st.button("🔄 整合數據"):
            # 解析邏輯 (簡化版)
            df_a = pd.read_excel(f_assign)
            df_t = pd.read_excel(f_time)
            
            # 建立教師與班級資料庫
            t_db = {} # 教師課表
            all_t = set()
            day_map = {"一":1,"二":2,"三":3,"四":4,"五":5}
            
            for _, r in df_t.iterrows():
                d = day_map.get(str(r['星期'])[-1], 0)
                p = int(re.search(r'\d+', str(r['節次'])).group())
                c, s = str(r['班級']), str(r['科目'])
                # 從配課表抓教師 
                t_list = df_a[(df_a['班級']==c) & (df_a['科目']==s)]['教師'].iloc[0].split('/')
                for t in t_list:
                    t = t.strip()
                    all_t.add(t)
                    if t not in t_db: t_db[t] = {}
                    t_db[t][(d, p)] = {"c": c, "s": s}
            
            st.session_state.update({"t_db": t_db, "all_t": sorted(list(all_t)), "template": f_temp.read(), "ready": True})
            st.success("✅ 系統已就緒")

# --- 主畫面：仿 DM 調代課作業 ---
if st.session_state.get("ready"):
    st.title("🗂️ 智慧調代課作業中心")
    
    # --- Step 1: 選擇欲代課課程 ---
    st.markdown("### **Step.1 選擇欲代課課程**")
    c1, c2, c3 = st.columns([2, 2, 3])
    with c1:
        sel_date = st.date_input("請假/調動日期", datetime.now())
        w_idx = sel_date.weekday() + 1
    with c2:
        absent_t = st.selectbox("請假/被調動教師", st.session_state.all_t)
    
    # 顯示該員當日課程
    lessons = []
    t_sched = st.session_state.t_db.get(absent_t, {})
    for p in range(1, 9):
        if (w_idx, p) in t_sched:
            info = t_sched[(w_idx, p)]
            lessons.append({"p": p, "c": info['c'], "s": info['s'], "label": f"第 {p} 節: {info['c']} {info['s']}"})
    
    if not lessons:
        st.warning("該教師當日無課程。")
    else:
        # 使用表格樣式顯示可選課程
        selected_l = st.radio("選擇課程：", lessons, format_func=lambda x: x['label'], horizontal=True)
        
        st.divider()
        
        # --- Step 2: 選擇代課教師 (衝堂檢查) ---
        st.markdown("### **Step.2 選擇代課教師**")
        
        # 自動篩選：該節次沒課的老師
        available_ts = []
        conflicted_ts = []
        for t in st.session_state.all_t:
            if (w_idx, selected_l['p']) in st.session_state.t_db.get(t, {}):
                conflicted_ts.append(t)
            else:
                available_ts.append(t)
        
        col_left, col_right = st.columns([1, 1])
        with col_left:
            mode = st.radio("變動類型", ["代課", "調課"], horizontal=True)
            sub_t = st.selectbox("🔍 選擇任課教師 (已過濾衝堂)", available_ts)
            
            # 顯示預覽內容 
            prefix = "代" if mode == "代課" else "調"
            content = f"{prefix}{selected_l['c']}\n{selected_l['s']}"
            st.info(f"💡 將於通知單填入：\n**{content.replace(chr(10), ' ')}**")
            
        with col_right:
            st.caption(f"📊 {sub_t} 老師的當日課表預覽")
            sub_day_sched = {f"第{i}節": "" for i in range(1, 9)}
            for (d, p), info in st.session_state.t_db.get(sub_t, {}).items():
                if d == w_idx: sub_day_sched[f"第{p}節"] = f"{info['c']} {info['s']}"
            st.table(pd.DataFrame([sub_day_sched]).T.rename(columns={0: "課程"}))

        if st.button("🪄 生成調代課通知單"):
            w_dates = get_roc_week(sel_date)
            change_data = {'day': w_idx, 'period': selected_l['p'], 'content': content}
            
            final_doc = generate_sub_notice(st.session_state.template, sub_t, change_data, w_dates)
            
            buf = BytesIO()
            final_doc.save(buf)
            st.success(f"🎉 {sub_t} 的通知單已準備就緒！")
            st.download_button(f"⬇️ 下載通知單 ({sub_t})", buf.getvalue(), f"{sel_date.strftime('%m%d')}_{sub_t}_通知單.docx")

else:
    st.info("請於左側上傳必要之數據檔案與 Word 樣板以啟動智慧系統。")
