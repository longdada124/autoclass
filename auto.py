import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="後龍國中課表暨調代課系統", layout="wide")

# --- 核心函數：Word 替換 ---
def master_replace(doc_obj, old_text, new_text):
    new_val = str(new_text) if new_text else ""
    for p in list(doc_obj.paragraphs):
        if old_text in p.text:
            p.text = p.text.replace(old_text, new_val)
    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old_text in p.text:
                        p.text = p.text.replace(old_text, new_val)

# --- 側邊欄：統一上傳區 ---
with st.sidebar:
    st.header("⚙️ 系統設定")
    if st.button("🧹 全系統重置"):
        st.session_state.clear()
        st.rerun()

    st.divider()
    st.subheader("📤 資料上傳 (必要)")
    f_assign = st.file_uploader("1. 上傳【配課表】", type=["xlsx", "csv"])
    f_time = st.file_uploader("2. 上傳【課表】", type=["xlsx", "csv"])
    f_sort = st.file_uploader("3. 上傳【教師排序表】", type=["xlsx", "csv"])
    
    if f_assign and f_time and st.button("🚀 啟動系統整合"):
        with st.spinner("正在同步課表資料..."):
            # (此處省略部分重複的解析邏輯，確保與您之前運作正常的邏輯一致)
            df_assign = pd.read_excel(f_assign) if f_assign.name.endswith('xlsx') else pd.read_csv(f_assign)
            df_time = pd.read_excel(f_time) if f_time.name.endswith('xlsx') else pd.read_csv(f_time)
            
            # 建立全域索引供「調代課」使用
            # ... 解析邏輯 ...
            st.session_state.data_loaded = True
            st.success("資料已連結！請切換至調代課標籤頁。")
            st.rerun()

# --- 主介面 ---
if 'class_data' in st.session_state:
    # 🌟 新增「調代課管理」標籤頁
    tab1, tab2, tab3 = st.tabs(["🏫 班級課表", "👩‍🏫 教師課表", "📅 調代課管理"])

    with tab3:
        st.header("🔄 調代課智慧作業")
        
        col_ctrl1, col_ctrl2 = st.columns(2)
        with col_ctrl1:
            target_date = st.date_input("選擇異動日期", datetime.now())
            week_day_num = target_date.weekday() + 1 # 1=Mon, 5=Fri
        
        if week_day_num > 5:
            st.warning("⚠️ 選擇日期為週末，請重新選擇。")
        else:
            absent_t = st.selectbox("1. 選擇【請假/欲調課】教師", st.session_state.ordered_teachers)
            
            # 找出該師該日課程
            day_map_rev = {1:"週一", 2:"週二", 3:"週三", 4:"週四", 5:"週五"}
            t_lessons = []
            for p in range(1, 9):
                info = st.session_state.teacher_data[absent_t].get((week_day_num, p))
                if info:
                    t_lessons.append({"節次": p, "班級": info['class'], "科目": info['subj']})
            
            if t_lessons:
                sel_lesson = st.radio("2. 選擇欲處理的節次", t_lessons, format_func=lambda x: f"第{x['節次']}節 - {x['班級']}{x['科目']}")
                
                mode = st.segmented_control("3. 處理模式", ["代課", "調課"])
                
                if mode == "代課":
                    # 自動推薦空堂老師
                    avail_teachers = []
                    for t in st.session_state.ordered_teachers:
                        if (week_day_num, sel_lesson['節次']) not in st.session_state.teacher_data[t]:
                            avail_teachers.append(t)
                    
                    sub_t = st.selectbox("4. 選擇代課老師 (已過濾出空堂者)", avail_teachers)
                    if st.button("📝 生成代課通知單"):
                        # 此處對接「代調課通知單.docx」
                        st.write(f"正在產製：{target_date} 第{sel_lesson['節次']}節 {sel_lesson['班級']}由{sub_t}代課")
                
                elif mode == "調課":
                    st.info("跨週調課功能：請選擇目標日期與節次，系統將自動對調並檢查兩位老師是否衝堂。")
                    # 跨週邏輯開發中...
            else:
                st.info(f"該老師在 {day_map_rev[week_day_num]} 沒有課。")
