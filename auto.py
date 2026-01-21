import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="後龍國中課表暨調代課系統", layout="wide")

# --- 核心工具函數 ---
def master_replace(doc_obj, old_text, new_text):
    """替換 Word 內的文字，包含表格與段落 (精準標籤替換版)"""
    if not new_text: new_text = ""
    # 1. 替換段落中的文字
    for p in doc_obj.paragraphs:
        if old_text in p.text:
            # 嘗試保留格式的替換
            for run in p.runs:
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)
            # 如果 run 切割太碎導致沒換到，強制整段替換 (會重置該段格式，但通常標籤是獨立的所以還好)
            if old_text in p.text: 
                p.text = p.text.replace(old_text, new_text)

    # 2. 替換表格中的文字
    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                # 針對儲存格內的每個段落檢查
                for p in cell.paragraphs:
                    if old_text in p.text:
                        p.text = p.text.replace(old_text, new_text)

def get_week_dates(base_date):
    """計算該週週一至週五的日期字串"""
    start_of_week = base_date - timedelta(days=base_date.weekday())
    dates = []
    for i in range(5):
        d = start_of_week + timedelta(days=i)
        # 轉換為民國年格式 115.02.09
        roc_year = d.year - 1911
        dates.append(f"{roc_year}.{d.month:02d}.{d.day:02d}")
    return dates

def fill_sub_notice(template_bytes, teacher_name, changes, week_dates):
    """
    填寫代調課通知單 (標籤版)
    changes: list of dict [{'day': 1-5, 'period': 1-8, 'text': '代702 國文'}]
    """
    doc = Document(BytesIO(template_bytes))
    
    # 1. 替換基本資料
    master_replace(doc, "{{TEACHER}}", teacher_name)
    
    # 2. 替換日期 {{D1}} ~ {{D5}}
    for i, d_str in enumerate(week_dates):
        master_replace(doc, f"{{{{D{i+1}}}}}", d_str)
    
    # 3. 先清空所有未使用的課表標籤 (避免印出來還有 {{1_1}} 這種字)
    # 我們先建立一個 "要填寫的格子清單"
    fill_map = {}
    for chg in changes:
        tag = f"{{{{{chg['day']}_{chg['period']}}}}}" # 格式: {{1_1}}
        fill_map[tag] = chg['text']

    # 4. 執行替換
    # 掃描所有可能的標籤 {{1_1}} 到 {{5_8}}
    for d in range(1, 6):
        for p in range(1, 9):
            tag = f"{{{{{d}_{p}}}}}"
            if tag in fill_map:
                # 如果這格有課，填入內容
                master_replace(doc, tag, fill_map[tag])
            else:
                # 如果這格沒課，把標籤清空
                master_replace(doc, tag, "")

    return doc

# --- 側邊欄與主邏輯 (維持不變，僅微調呼叫部分) ---
with st.sidebar:
    st.header("⚙️ 系統資料管理")
    if st.button("🧹 重置所有資料"):
        st.session_state.clear()
        st.rerun()
    
    st.divider()
    st.info("請依序上傳三個檔案")
    f_assign = st.file_uploader("1. 配課表", type=["xlsx", "csv"])
    f_time = st.file_uploader("2. 課表", type=["xlsx", "csv"])
    f_sort = st.file_uploader("3. 教師排序表", type=["xlsx", "csv"])
    
    if f_assign and f_time and st.button("🚀 執行系統整合"):
        with st.spinner("正在讀取資料與樣板..."):
            try:
                # 1. 讀取 Excel
                df_assign = pd.read_excel(f_assign) if f_assign.name.endswith('xlsx') else pd.read_csv(f_assign)
                df_time = pd.read_excel(f_time) if f_time.name.endswith('xlsx') else pd.read_csv(f_time)
                
                # 2. 讀取 GitHub 內建樣板
                try:
                    with open("班級樣板.docx", "rb") as f: st.session_state.class_template = f.read()
                    with open("教師樣板.docx", "rb") as f: st.session_state.teacher_template = f.read()
                    with open("代調課通知單.docx", "rb") as f: st.session_state.sub_template = f.read()
                except FileNotFoundError:
                    st.warning("⚠️ 部分 Word 樣板未找到，請確認 GitHub 檔案名稱是否正確。")

                # 3. 解析資料 (標準邏輯)
                assign_lookup = []
                all_teachers_db = set()
                tutors = {}
                
                for _, row in df_assign.iterrows():
                    c, s, t_raw = str(row['班級']).strip(), str(row['科目']).strip(), str(row['教師']).strip()
                    t_list = [name.strip() for name in t_raw.split('/')]
                    for t in t_list:
                        if t and t != "nan":
                            assign_lookup.append({'c': c, 's': s, 't': t})
                            all_teachers_db.add(t)
                    if s == "班級": tutors[c] = t_raw

                ordered_teachers = sorted(list(all_teachers_db)) # 簡化排序邏輯以防錯誤
                if f_sort:
                    try:
                        df_s = pd.read_excel(f_sort) if f_sort.name.endswith('xlsx') else pd.read_csv(f_sort)
                        # 簡單處理排序
                        s_list = [str(x).strip() for x in df_s.iloc[:,0].tolist()]
                        ordered_teachers = [t for t in s_list if t in all_teachers_db] + [t for t in ordered_teachers if t not in s_list]
                    except: pass

                # 解析課表
                class_data = {}
                teacher_data = {}
                day_map = {"一":1,"二":2,"三":3,"四":4,"五":5,"週一":1,"週二":2,"週三":3,"週四":4,"週五":5}
                
                for _, row in df_time.iterrows():
                    c_raw, s_raw = str(row['班級']).strip(), str(row['科目']).strip()
                    d = day_map.get(str(row['星期']).strip(), 0)
                    p_match = re.search(r'\d+', str(row['節次']))
                    
                    if p_match and d > 0:
                        p = int(p_match.group())
                        curr_t_list = [x['t'] for x in assign_lookup if x['c'] == c_raw and x['s'] == s_raw]
                        display_t = "/".join(curr_t_list) if curr_t_list else "未知"
                        
                        if c_raw not in class_data: class_data[c_raw] = {}
                        class_data[c_raw][(d, p)] = {"subj": s_raw, "teacher": display_t}
                        
                        for t in curr_t_list:
                            if t not in teacher_data: teacher_data[t] = {}
                            teacher_data[t][(d, p)] = {"subj": s_raw, "class": c_raw}

                st.session_state.update({
                    "class_data": class_data,
                    "teacher_data": teacher_data,
                    "ordered_teachers": ordered_teachers,
                    "data_ready": True
                })
                st.success("✅ 資料整合完畢！")
                st.rerun()

            except Exception as e:
                st.error(f"❌ 解析發生錯誤：{e}")

# --- 主畫面 ---
if st.session_state.get("data_ready"):
    tab1, tab2, tab3 = st.tabs(["🏫 班級課表", "👩‍🏫 教師課表", "📅 調代課管理(NEW)"])
    
    with tab1:
        st.info("班級課表預覽功能區 (已省略以節省篇幅)")

    with tab2:
        st.info("教師課表預覽功能區 (已省略以節省篇幅)")

    with tab3:
        st.header("🔄 調代課通知單產製")
        
        col1, col2 = st.columns(2)
        with col1:
            target_date = st.date_input("選擇代課日期", datetime.now())
            week_num = target_date.weekday() + 1
            week_dates = get_week_dates(target_date)
            st.caption(f"通知單日期區間：{week_dates[0]} ~ {week_dates[4]}")

        with col2:
            absent_teacher = st.selectbox("請假/被代課教師", st.session_state.ordered_teachers)
        
        st.subheader("1. 選擇要代課的節次")
        day_lessons = []
        info_t = st.session_state.teacher_data.get(absent_teacher, {})
        
        for p in range(1, 9):
            info = info_t.get((week_num, p))
            if info:
                day_lessons.append({
                    "節次": p, 
                    "班級": info['class'], 
                    "科目": info['subj'],
                    "desc": f"第{p}節 - {info['class']}{info['subj']}"
                })
        
        if not day_lessons:
            st.warning(f"{absent_teacher} 老師在 {target_date} (週{week_num}) 沒有課程。")
        else:
            selected_lesson = st.radio("請勾選課程：", day_lessons, format_func=lambda x: x['desc'])
            
            st.divider()
            st.subheader("2. 選擇代課教師")
            
            available_teachers = []
            for t in st.session_state.ordered_teachers:
                if (week_num, selected_lesson['節次']) not in st.session_state.teacher_data.get(t, {}):
                    available_teachers.append(t)
            
            sub_teacher = st.selectbox("選擇代課教師", available_teachers)
            
            if st.button("🖨️ 產生代課通知單 (Word)"):
                if not st.session_state.get('sub_template'):
                    st.error("❌ 找不到樣板，請確認 GitHub 已上傳【代調課通知單.docx】")
                else:
                    # 準備寫入資料
                    change_info = {
                        'day': week_num,
                        'period': selected_lesson['節次'],
                        'text': f"代{selected_lesson['班級']} {selected_lesson['科目']}"
                    }
                    
                    doc_sub = fill_sub_notice(
                        st.session_state.sub_template,
                        sub_teacher, 
                        [change_info],
                        week_dates
                    )
                    
                    buf = BytesIO()
                    doc_sub.save(buf)
                    file_name = f"{target_date.strftime('%m%d')}_{sub_teacher}_代課單.docx"
                    st.download_button(f"⬇️ 下載 {sub_teacher} 的通知單", buf.getvalue(), file_name)
                    st.success(f"✅ 已生成！請打開檔案確認 {sub_teacher} 的名字與 {selected_lesson['班級']} 的代課內容是否正確填入。")

else:
    st.info("👋 請於左側上傳 3 個資料檔並執行整合。")
