import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
from datetime import datetime, timedelta

# --- 1. 系統設定 ---
st.set_page_config(page_title="後龍國中智慧代調課系統", layout="wide")

# --- 2. 核心 Word 處理邏輯 (強化版) ---

def master_replace(doc_obj, old_text, new_text):
    """更強大的替換邏輯，確保能搜尋到所有文字區段"""
    new_val = str(new_text) if new_text is not None else ""
    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old_text in p.text:
                        # 處理 Run 級別的替換，這能保留樣板的字體格式
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
                        # 二次檢查：如果 Run 被 Word 切太碎導致沒換到，直接對 Paragraph 處理
                        if old_text in p.text:
                            p.text = p.text.replace(old_text, new_val)

def generate_docx(template_bytes, teacher, change_data, week_dates):
    """產製通知單，確保所有日期與 40 個格子標籤都被處理"""
    doc = Document(BytesIO(template_bytes))
    
    # 1. 填寫教師名稱 
    master_replace(doc, "{{TEACHER}}", teacher)
    
    # 2. 填寫週一至週五日期 ({{D1}} ~ {{D5}}) 
    for i, d_str in enumerate(week_dates):
        tag = f"{{{{D{i+1}}}}}"
        master_replace(doc, tag, d_str)
    
    # 3. 遍歷 5 天 * 8 節課 = 40 個標籤 
    target_tag = f"{{{{{change_data['day']}_{change_data['period']}}}}}"
    for d in range(1, 6):
        for p in range(1, 9):
            tag = f"{{{{{d}_{p}}}}}"
            if tag == target_tag:
                # 填入代/調課內容 [cite: 7, 9]
                master_replace(doc, tag, change_data['content'])
            else:
                # 沒用到的標籤「必須」清空，確保畫面空白
                master_replace(doc, tag, "")
    return doc

# --- 3. 側邊欄與資料整合 (修正 IndexError) ---
with st.sidebar:
    st.header("📂 數據與樣板管理")
    f_assign = st.file_uploader("1. 上傳配課表 (Excel)", type=["xlsx"])
    f_time = st.file_uploader("2. 上傳課表 (Excel)", type=["xlsx"])
    f_temp = st.file_uploader("3. 上傳 Word 樣板", type=["docx"])
    
    if f_assign and f_time and f_temp:
        if st.button("🔄 執行資料整合"):
            try:
                # 讀取並清除前後空白
                df_a = pd.read_excel(f_assign).astype(str).apply(lambda x: x.str.strip())
                df_t = pd.read_excel(f_time).astype(str).apply(lambda x: x.str.strip())
                
                t_db = {}
                all_t = set()
                day_map = {"一":1, "二":2, "三":3, "四":4, "五":5}

                for _, r in df_t.iterrows():
                    d_match = re.search(r'[一二三四五]', r['星期'])
                    p_match = re.search(r'\d+', r['節次'])
                    if not d_match or not p_match: continue
                    
                    d, p = day_map[d_match.group()], int(p_match.group())
                    c, s = r['班級'], r['科目']
                    
                    # 修正 IndexError：先過濾再確認是否有資料
                    match_data = df_a[(df_a['班級'] == c) & (df_a['科目'] == s)]
                    if not match_data.empty:
                        teachers = str(match_data.iloc[0]['教師']).split('/')
                    else:
                        teachers = ["(無配課)"]

                    for t in [x.strip() for x in teachers]:
                        all_t.add(t)
                        if t not in t_db: t_db[t] = {}
                        t_db[t][(d, p)] = {"c": c, "s": s}
                
                st.session_state.update({
                    "t_db": t_db, "all_t": sorted(list(all_t)), 
                    "template": f_temp.read(), "ready": True
                })
                st.success("✅ 數據載入成功！")
            except Exception as e:
                st.error(f"整合發生錯誤：{e}")

# --- 4. 主畫面 (代調課邏輯) ---
if st.session_state.get("ready"):
    st.title("📑 智慧代調課管理系統")
    
    st.subheader("Step 1. 選擇要處理的課程")
    c1, c2 = st.columns(2)
    with c1:
        date_sel = st.date_input("變動日期", datetime.now())
        w_idx = date_sel.weekday() + 1
    with c2:
        from_t = st.selectbox("請假教師", st.session_state.all_t)
    
    # 取得請假教師當日課程
    lessons = []
    for p in range(1, 9):
        info = st.session_state.t_db.get(from_t, {}).get((w_idx, p))
        if info:
            lessons.append({"p": p, "c": info['c'], "s": info['s'], "label": f"第{p}節 {info['c']}{info['s']}"})
    
    if not lessons:
        st.warning("該教師當天沒有課程。")
    else:
        sel_l = st.radio("欲調整的節次", lessons, format_func=lambda x: x['label'], horizontal=True)
        
        st.divider()
        st.subheader("Step 2. 安排代調課細節")
        
        mode = st.radio("變動性質", ["代課", "調課"], horizontal=True)
        
        # 智慧排除衝堂教師
        avail_ts = [t for t in st.session_state.all_t if (w_idx, sel_l['p']) not in st.session_state.t_db.get(t, {})]
        to_t = st.selectbox("代課/接收教師 (已排除衝堂)", avail_ts)
        
        content = f"{mode[:1]}{sel_l['c']}\n{sel_l['s']}"
        st.info(f"📋 即將填入：**{content.replace(chr(10), ' ')}** 到 **{to_t}** 老師的通知單中。")

        if st.button("🚀 生成通知單"):
            # 計算該週日期字串
            mon = date_sel - timedelta(days=date_sel.weekday())
            w_strs = [f"{mon.year-1911}.{(mon+timedelta(days=i)).month:02d}.{(mon+timedelta(days=i)).day:02d}" for i in range(5)]
            
            # 生成檔案
            final_doc = generate_docx(
                st.session_state.template, 
                to_t, 
                {'day': w_idx, 'period': sel_l['p'], 'content': content}, 
                w_strs
            )
            
            buf = BytesIO()
            final_doc.save(buf)
            st.success("產製成功！")
            st.download_button(f"⬇️ 下載通知單 ({to_t})", buf.getvalue(), f"{to_t}_代調課單.docx")
else:
    st.info("請於左側上傳數據後開始作業。")
