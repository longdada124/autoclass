import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from io import BytesIO
import re
from datetime import datetime, timedelta

# --- 1. 字體與 Word 核心邏輯 ---
def set_font_style(run, font_name="標楷體"):
    """強制鎖定中文字體為標楷體"""
    run.font.name = font_name
    # 針對 Word 的東亞文字屬性進行設定
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def master_replace(doc_obj, old_text, new_text):
    """替換文字並套用標楷體，支援換行"""
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
                        # 二次保險：處理可能被切碎的 Run
                        if old_text in p.text:
                            p.text = p.text.replace(old_text, new_val)
                            for r in p.runs:
                                set_font_style(r)

def generate_docx(template_bytes, teacher, change_data, week_dates):
    """產製通知單：填入內容並強制清除所有剩餘標籤"""
    doc = Document(BytesIO(template_bytes))
    
    # 1. 填寫抬頭
    master_replace(doc, "{{TEACHER}}", teacher)
    
    # 2. 填寫日期標籤 (D1~D5)
    for i, d_str in enumerate(week_dates):
        master_replace(doc, f"{{{{D{i+1}}}}}", d_str)
    
    # 3. 遍歷 40 個課程格子 (1_1 ~ 5_8)
    target_tag = f"{{{{{change_data['day']}_{change_data['period']}}}}}"
    for d in range(1, 6):
        for p in range(1, 9):
            tag = f"{{{{{d}_{p}}}}}"
            # 是選中那格就填代調課資訊，其餘一律變空白
            content = change_data['content'] if tag == target_tag else ""
            master_replace(doc, tag, content)
            
    return doc

# --- 2. 頁面配置與資料處理 ---
st.set_page_config(page_title="後龍國中智慧代調課系統", layout="wide")

# 初始化 Session State 用於記憶「選取的課程」
if 'selected_lesson' not in st.session_state:
    st.session_state.selected_lesson = None

with st.sidebar:
    st.header("📂 資料上傳區")
    f_assign = st.file_uploader("1. 上傳配課表 (Excel)", type=["xlsx"])
    f_time = st.file_uploader("2. 上傳課表 (Excel)", type=["xlsx"])
    f_temp = st.file_uploader("3. 上傳通知單樣板 (.docx)", type=["docx"])
    
    if f_assign and f_time and f_temp:
        if st.button("🔄 執行數據彙整"):
            try:
                # 讀取並清洗資料
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
                    
                    # 搜尋配課老師
                    match = df_a[(df_a['班級'] == c) & (df_a['科目'] == s)]
                    teachers = str(match.iloc[0]['教師']).split('/') if not match.empty else ["未知"]
                    
                    for t in [x.strip() for x in teachers]:
                        all_t.add(t)
                        if t not in t_db: t_db[t] = {}
                        t_db[t][(d, p)] = {"c": c, "s": s}
                
                st.session_state.update({
                    "t_db": t_db, 
                    "all_t": sorted(list(all_t)), 
                    "template": f_temp.read(), 
                    "ready": True
                })
                st.success("✅ 資料整合完成！")
            except Exception as e:
                st.error(f"整合發生錯誤：{e}")

# --- 3. 主畫面：互動式操作 ---
if st.session_state.get("ready"):
    st.title("📑 智慧代調課互動作業")
    
    # 選擇請假老師
    sel_teacher = st.selectbox("🔍 步驟 1：請選擇「請假/受調動」教師", st.session_state.all_t)
    
    # 繪製視覺化互動課表
    st.write(f"### 📅 {sel_teacher} 老師的週課表")
    st.caption("按下方按鈕選取要「被代」或「被調」的課程：")
    
    days_labels = ["一", "二", "三", "四", "五"]
    grid_cols = st.columns(5)
    
    for d_idx, day_name in enumerate(days_labels):
        with grid_cols[d_idx]:
            st.button(f"週{day_name}", disabled=True, use_container_width=True)
            for p in range(1, 9):
                info = st.session_state.t_db.get(sel_teacher, {}).get((d_idx + 1, p))
                if info:
                    # 這是該老師有課的格子，點擊可選取
                    btn_label = f"第{p}節\n{info['c']}\n{info['s']}"
                    if st.button(btn_label, key=f"btn_{d_idx}_{p}", use_container_width=True, type="primary"):
                        st.session_state.selected_lesson = {
                            'day': d_idx + 1, 'period': p, 'c': info['c'], 's': info['s']
                        }
                else:
                    # 空堂
                    st.button(f"第{p}節\n(空)", key=f"empty_{d_idx}_{p}", disabled=True, use_container_width=True)

    # 如果已經選取了一門課，顯示下一步
    if st.session_state.selected_lesson:
        l = st.session_state.selected_lesson
        st.divider()
        st.success(f"📍 已選定：週{l['day']} 第{l['period']}節 - {l['c']} {l['s']}")
        
        c1, c2, c3 = st.columns([2, 2, 2])
        with c1:
            date_sel = st.date_input("🗓️ 實際變動日期", datetime.now())
            mode = st.radio("🔄 變動性質", ["代課", "調課"], horizontal=True)
        
        with c2:
            # 智慧衝堂檢查：過濾出該時段沒課的老師
            avail_ts = [t for t in st.session_state.all_t if (l['day'], l['period']) not in st.session_state.t_db.get(t, {})]
            to_teacher = st.selectbox("👤 步驟 2：選擇「代課/接收」教師", avail_ts)
            st.caption(f"💡 系統已自動排除第 {l['period']} 節衝堂者")
            
        with c3:
            st.write("📝 內容預覽")
            content = f"{mode[:1]}{l['c']}\n{l['s']}"
            st.code(content)
            
            if st.button("🚀 生成通知單", use_container_width=True):
                # 計算日期 (週一到週五)
                monday = date_sel - timedelta(days=date_sel.weekday())
                week_strs = [f"{monday.year-1911}.{(monday+timedelta(days=i)).month:02d}.{(monday+timedelta(days=i)).day:02d}" for i in range(5)]
                
                # 生成檔案
                final_doc = generate_docx(
                    st.session_state.template, 
                    to_teacher, 
                    {'day': l['day'], 'period': l['period'], 'content': content}, 
                    week_strs
                )
                
                buf = BytesIO()
                final_doc.save(buf)
                st.download_button(
                    f"⬇️ 下載 {to_teacher} 的通知單", 
                    buf.getvalue(), 
                    f"{date_sel.strftime('%m%d')}_{to_teacher}_通知單.docx",
                    use_container_width=True
                )
else:
    st.info("👋 您好！請先於左側上傳 Excel 與 Word 樣板，完成資料整合後即可開始作業。")
