import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
from datetime import datetime, timedelta

# --- 1. 系統設定 ---
st.set_page_config(page_title="後龍國中智慧代調課系統", layout="wide")

# --- 2. 核心功能函數 ---

def master_replace(doc_obj, old_text, new_text):
    """安全替換 Word 標籤，支援換行並保留格式"""
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

def generate_docx(template_bytes, teacher, change_data, week_dates):
    """產製通知單並清除所有未使用的 {{d_p}} 標籤"""
    doc = Document(BytesIO(template_bytes))
    master_replace(doc, "{{TEACHER}}", teacher)
    for i, d_str in enumerate(week_dates):
        master_replace(doc, f"{{{{D{i+1}}}}}", d_str)
    
    # 填寫內容並「徹底排空」其他格子
    target_tag = f"{{{{{change_data['day']}_{change_data['period']}}}}}"
    for d in range(1, 6):
        for p in range(1, 9):
            tag = f"{{{{{d}_{p}}}}}"
            content = change_data['content'] if tag == target_tag else ""
            master_replace(doc, tag, content)
    return doc

# --- 3. 側邊欄：資料整合 (修正 IndexError) ---
with st.sidebar:
    st.header("📂 數據中心")
    f_assign = st.file_uploader("1. 上傳配課表 (Excel)", type=["xlsx"])
    f_time = st.file_uploader("2. 上傳課表 (Excel)", type=["xlsx"])
    f_temp = st.file_uploader("3. 上傳 Word 樣板", type=["docx"])
    
    if f_assign and f_time and f_temp:
        if st.button("🚀 執行資料彙整"):
            try:
                df_a = pd.read_excel(f_assign).astype(str).apply(lambda x: x.str.strip())
                df_t = pd.read_excel(f_time).astype(str).apply(lambda x: x.str.strip())
                
                t_db = {}     # 教師課表索引
                class_db = {} # 班級課表索引
                all_teachers = set()
                day_map = {"一":1, "二":2, "三":3, "四":4, "五":5}

                for _, r in df_t.iterrows():
                    # 抓取星期與節次
                    d_match = re.search(r'[一二三四五]', r['星期'])
                    p_match = re.search(r'\d+', r['節次'])
                    if not d_match or not p_match: continue
                    
                    d, p = day_map[d_match.group()], int(p_match.group())
                    c, s = r['班級'], r['科目']
                    
                    # 修正 IndexError: 先檢查配課表是否存在該班級與科目
                    match_rows = df_a[(df_a['班級'] == c) & (df_a['科目'] == s)]
                    if not match_rows.empty:
                        t_list = str(match_rows.iloc[0]['教師']).split('/')
                    else:
                        t_list = ["未知"] # 找不到老師時顯示未知，不崩潰

                    # 建立索引
                    for t in [x.strip() for x in t_list]:
                        all_teachers.add(t)
                        if t not in t_db: t_db[t] = {}
                        t_db[t][(d, p)] = {"c": c, "s": s}
                    
                    if c not in class_db: class_db[c] = {}
                    class_db[c][(d, p)] = {"s": s, "t": "/".join(t_list)}

                st.session_state.update({
                    "t_db": t_db, "class_db": class_db, 
                    "all_t": sorted(list(all_teachers)), 
                    "template": f_temp.read(), "ready": True
                })
                st.success("✅ 彙整完畢！")
            except Exception as e:
                st.error(f"解析失敗：{e}")

# --- 4. 主畫面：分頁系統 ---
if st.session_state.get("ready"):
    tab1, tab2, tab3 = st.tabs(["👩‍🏫 教師課表彙整", "🏫 班級課表彙整", "🔄 智慧代調課"])

    # --- 教師課表 ---
    with tab1:
        sel_t = st.selectbox("選擇教師", st.session_state.all_t)
        grid = {d: [""]*8 for d in ["一","二","三","四","五"]}
        for (d, p), info in st.session_state.t_db.get(sel_t, {}).items():
            if 1 <= p <= 8: grid[list(grid.keys())[d-1]][p-1] = f"{info['c']}\n{info['s']}"
        st.table(pd.DataFrame(grid, index=[f"第{i}節" for i in range(1,9)]))

    # --- 班級課表 ---
    with tab2:
        sel_c = st.selectbox("選擇班級", sorted(list(st.session_state.class_db.keys())))
        grid_c = {d: [""]*8 for d in ["一","二","三","四","五"]}
        for (d, p), info in st.session_state.class_db.get(sel_c, {}).items():
            if 1 <= p <= 8: grid_c[list(grid_c.keys())[d-1]][p-1] = f"{info['s']}\n{info['t']}"
        st.table(pd.DataFrame(grid_c, index=[f"第{i}節" for i in range(1,9)]))

    # --- 智慧代調課 (仿 DM 功能) ---
    with tab3:
        st.subheader("Step 1. 選擇原始課程")
        c1, c2 = st.columns(2)
        with c1:
            date_sel = st.date_input("變動日期", datetime.now())
            w_idx = date_sel.weekday() + 1
        with c2:
            from_t = st.selectbox("原任課老師 (請假方)", st.session_state.all_t)
        
        # 抓取該老師當天課程
        daily_lessons = []
        for p in range(1, 9):
            if (w_idx, p) in st.session_state.t_db.get(from_t, {}):
                info = st.session_state.t_db[from_t][(w_idx, p)]
                daily_lessons.append({"p": p, "c": info['c'], "s": info['s'], "label": f"第{p}節 {info['c']}{info['s']}"})
        
        if not daily_lessons:
            st.warning("該教師此日無課。")
        else:
            sel_l = st.radio("選擇欲處理的節次", daily_lessons, format_func=lambda x: x['label'], horizontal=True)
            
            st.divider()
            st.subheader("Step 2. 安排代調課教師")
            
            mode = st.radio("變動性質", ["代課", "調課"], horizontal=True)
            
            # 智慧過濾衝堂教師
            available_ts = [t for t in st.session_state.all_t if (w_idx, sel_l['p']) not in st.session_state.t_db.get(t, {})]
            to_t = st.selectbox(f"選擇接收教師 (已自動過濾第{sel_l['p']}節衝堂者)", available_ts)
            
            content = f"{mode[:1]}{sel_l['c']}\n{sel_l['s']}"
            st.info(f"📝 預覽內容：**{content.replace(chr(10), ' ')}** (將填入 {to_t} 的通知單)")

            if st.button("🖨️ 產生通知單"):
                # 計算該週日期
                monday = date_sel - timedelta(days=date_sel.weekday())
                week_strs = [f"{monday.year-1911}.{(monday+timedelta(days=i)).month:02d}.{(monday+timedelta(days=i)).day:02d}" for i in range(5)]
                
                final_doc = generate_docx(st.session_state.template, to_t, {'day': w_idx, 'period': sel_l['p'], 'content': content}, week_strs)
                
                output = BytesIO()
                final_doc.save(output)
                st.success("通知單產製完成！")
                st.download_button(f"⬇️ 下載 {to_t} 的通知單", output.getvalue(), f"{to_t}_通知單.docx")

else:
    st.info("👋 請先於左側上傳 Excel 課表與 Word 樣板，並點擊「執行資料彙整」。")
