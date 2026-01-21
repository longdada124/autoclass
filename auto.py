import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from io import BytesIO
import re
from datetime import datetime, timedelta

# --- 1. 核心字體與 Word 邏輯 (支援標楷體與自動清理) ---

def set_font_style(run, font_name="標楷體"):
    """強制鎖定中文字體為標楷體 (Word 底層東亞字體設定)"""
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def master_replace(doc_obj, old_text, new_text):
    """替換標籤並套用格式，支援換行並保留標楷體"""
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
                                    set_font_style(run)
                                    for part in parts[1:]:
                                        run.add_break()
                                        new_run = run.add_text(part)
                                        set_font_style(new_run)
                                else:
                                    run.text = run.text.replace(old_text, new_val)
                                    set_font_style(run)
                        # 保險機制：處理被拆分的 Run
                        if old_text in p.text:
                            p.text = p.text.replace(old_text, new_val)
                            for r in p.runs: set_font_style(r)

def generate_docx(template_bytes, target_teacher, change_info, week_dates):
    """產製通知單並徹底清理 40 個格子標籤"""
    doc = Document(BytesIO(template_bytes))
    
    # 填寫抬頭與日期
    master_replace(doc, "{{TEACHER}}", target_teacher)
    for i, d_str in enumerate(week_dates):
        master_replace(doc, f"{{{{D{i+1}}}}}", d_str)
    
    # 填寫 40 個課程格子 (1_1 到 5_8)
    target_tag = f"{{{{{change_info['day']}_{change_info['period']}}}}}"
    for d in range(1, 6):
        for p in range(1, 9):
            tag = f"{{{{{d}_{p}}}}}"
            # 僅在目標格子填入內容，其餘一律清空
            content = change_info['content'] if tag == target_tag else ""
            master_replace(doc, tag, content)
            
    return doc

# --- 2. 系統設定與持久化儲存邏輯 ---

st.set_page_config(page_title="後龍國中智慧代調課系統", layout="wide")

# 初始化 Session State (讓資料留在網頁中)
if 'db_ready' not in st.session_state:
    st.session_state.db_ready = False
if 'selected_lesson' not in st.session_state:
    st.session_state.selected_lesson = None

with st.sidebar:
    st.header("📂 資料管理中心")
    
    # 如果已經有資料，顯示狀態而非重新上傳
    if st.session_state.db_ready:
        st.success("✅ 資料庫與樣板已就緒")
        if st.button("🗑️ 清除資料並重新上傳"):
            st.session_state.db_ready = False
            st.rerun()
    else:
        f_assign = st.file_uploader("1. 上傳配課表 (Excel)", type=["xlsx"])
        f_time = st.file_uploader("2. 上傳課表 (Excel)", type=["xlsx"])
        f_temp = st.file_uploader("3. 上傳 Word 樣板 (.docx)", type=["docx"])
        
        if f_assign and f_time and f_temp:
            if st.button("🚀 啟動系統 (儲存至網頁)"):
                # 處理資料
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
                
                # 存入 Session State 達成持久化
                st.session_state.update({
                    "t_db": t_db,
                    "all_t": sorted(list(all_t)),
                    "template": f_temp.read(),
                    "db_ready": True
                })
                st.rerun()

# --- 3. 主畫面：互動作業區 ---

if st.session_state.db_ready:
    st.title("📑 智慧代調課作業系統")
    
    # 步驟 1：選擇老師與課程
    t_list = st.session_state.all_t
    sel_teacher = st.selectbox("🔍 步驟 1：請選擇「請假/受調動」教師", t_list)
    
    st.write(f"### 📅 {sel_teacher} 老師的週課表")
    st.caption("請點擊下方藍色課程按鈕發動作業：")
    
    # 顯示互動格網
    grid_cols = st.columns(5)
    for d_idx in range(5):
        with grid_cols[d_idx]:
            st.button(f"週{['一','二','三','四','五'][d_idx]}", disabled=True, use_container_width=True)
            for p in range(1, 9):
                info = st.session_state.t_db.get(sel_teacher, {}).get((d_idx + 1, p))
                if info:
                    label = f"第{p}節\n{info['c']}\n{info['s']}"
                    if st.button(label, key=f"btn_{d_idx}_{p}", use_container_width=True, type="primary"):
                        st.session_state.selected_lesson = {
                            'day': d_idx + 1, 'period': p, 'c': info['c'], 's': info['s']
                        }
                else:
                    st.button(f"第{p}節\n-", key=f"empty_{d_idx}_{p}", disabled=True, use_container_width=True)

    # 步驟 2：處理代調課細節
    if st.session_state.selected_lesson:
        l = st.session_state.selected_lesson
        st.divider()
        st.success(f"📍 已選定：週{l['day']} 第{l['period']}節 - {l['c']} {l['s']}")
        
        c1, c2, c3 = st.columns(3)
        with c1:
            change_date = st.date_input("🗓️ 變動日期", datetime.now())
            mode = st.radio("🔄 性質", ["代課", "調課"], horizontal=True)
        
        with c2:
            # 智慧衝堂檢查 (DM 同款功能)
            avail_ts = [t for t in t_list if (l['day'], l['period']) not in st.session_state.t_db.get(t, {})]
            to_teacher = st.selectbox("👤 選擇「代課/接收」教師", avail_ts)
            st.caption(f"💡 已排除第 {l['period']} 節衝堂者")
            
        with c3:
            st.write("📝 預覽內容：")
            content = f"{mode[:1]}{l['c']}\n{l['s']}"
            st.info(content.replace("\n", " "))
            
            if st.button("🚀 生成通知單並下載", use_container_width=True):
                # 計算該週日期
                monday = change_date - timedelta(days=change_date.weekday())
                week_strs = [f"{monday.year-1911}.{(monday+timedelta(days=i)).month:02d}.{(monday+timedelta(days=i)).day:02d}" for i in range(5)]
                
                # 生成檔案
                final_docx = generate_docx(
                    st.session_state.template,
                    to_teacher,
                    {'day': l['day'], 'period': l['period'], 'content': content},
                    week_strs
                )
                
                buf = BytesIO()
                final_docx.save(buf)
                st.download_button(
                    f"⬇️ 下載 {to_teacher} 通知單",
                    buf.getvalue(),
                    f"{change_date.strftime('%m%d')}_{to_teacher}_通知單.docx",
                    use_container_width=True
                )
else:
    st.info("👋 您好！請從左側上傳 Excel 課表與 Word 樣板，系統將自動儲存資料供您連續作業。")
