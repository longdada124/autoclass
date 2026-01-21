import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
from datetime import datetime, timedelta

# 設定網頁標題與寬度
st.set_page_config(page_title="後龍國中課表暨調代課系統", layout="wide")

# --- 核心工具函數 ---

def master_replace(doc_obj, old_text, new_text):
    """
    進階替換函數：支援換行符號 \n，並盡可能保留原有的字體格式。
    """
    new_val = str(new_text) if new_text is not None else ""
    
    # 1. 替換段落文字 (主要用於標題名字、日期)
    for p in doc_obj.paragraphs:
        if old_text in p.text:
            for run in p.runs:
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_val)

    # 2. 替換表格文字 (主要用於課表格子，支援換行)
    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old_text in p.text:
                        for run in p.runs:
                            if old_text in run.text:
                                if "\n" in new_val:
                                    # 處理換行需求
                                    parts = new_val.split("\n")
                                    run.text = run.text.replace(old_text, parts[0])
                                    for part in parts[1:]:
                                        run.add_break() # 插入 Word 的換行符
                                        run.add_text(part)
                                else:
                                    run.text = run.text.replace(old_text, new_val)

def get_week_dates(base_date):
    """計算該週週一至週五的民國年日期"""
    start_of_week = base_date - timedelta(days=base_date.weekday())
    dates = []
    for i in range(5):
        d = start_of_week + timedelta(days=i)
        roc_year = d.year - 1911
        dates.append(f"{roc_year}.{d.month:02d}.{d.day:02d}")
    return dates

def fill_sub_notice(template_bytes, teacher_name, changes, week_dates):
    """產製代課通知單核心邏輯"""
    doc = Document(BytesIO(template_bytes))
    
    # 填寫抬頭老師名字與五天的日期標籤
    master_replace(doc, "{{TEACHER}}", teacher_name)
    for i, d_str in enumerate(week_dates):
        master_replace(doc, f"{{{{D{i+1}}}}}", d_str)
    
    # 準備填寫課表的資料 Map
    fill_map = {}
    for chg in changes:
        tag = f"{{{{{chg['day']}_{chg['period']}}}}}"
        fill_map[tag] = chg['text']

    # 掃描並填寫 1_1 到 5_8 的所有格子
    for d in range(1, 6):
        for p in range(1, 9):
            tag = f"{{{{{d}_{p}}}}}"
            if tag in fill_map:
                master_replace(doc, tag, fill_map[tag])
            else:
                master_replace(doc, tag, "") # 沒課的格子標籤清空
    return doc

# --- 側邊欄：資料上傳 ---

with st.sidebar:
    st.header("⚙️ 系統資料管理")
    if st.button("🧹 清空所有資料"):
        st.session_state.clear()
        st.rerun()
    
    st.divider()
    f_assign = st.file_uploader("1. 上傳配課表 (xlsx/csv)", type=["xlsx", "csv"])
    f_time = st.file_uploader("2. 上傳課表 (xlsx/csv)", type=["xlsx", "csv"])
    f_sort = st.file_uploader("3. 上傳教師排序表 (xlsx/csv)", type=["xlsx", "csv"])
    
    if f_assign and f_time and st.button("🚀 執行資料整合"):
        try:
            # 讀取 Excel/CSV
            df_assign = pd.read_excel(f_assign) if f_assign.name.endswith('xlsx') else pd.read_csv(f_assign)
            df_time = pd.read_excel(f_time) if f_time.name.endswith('xlsx') else pd.read_csv(f_time)
            
            # 載入內建樣板
            try:
                with open("班級樣板.docx", "rb") as f: st.session_state.class_template = f.read()
                with open("教師樣板.docx", "rb") as f: st.session_state.teacher_template = f.read()
                with open("代調課通知單.docx", "rb") as f: st.session_state.sub_template = f.read()
            except:
                st.warning("⚠️ 提醒：GitHub 內缺少部分 .docx 樣板檔案。")

            # 資料處理邏輯
            assign_lookup = []
            all_teachers = set()
            for _, row in df_assign.iterrows():
                c, s, t_raw = str(row['班級']).strip(), str(row['科目']).strip(), str(row['教師']).strip()
                for t in [x.strip() for x in t_raw.split('/')]:
                    if t and t != "nan":
                        assign_lookup.append({'c': c, 's': s, 't': t})
                        all_teachers.add(t)

            # 課表解析
            class_db, teacher_db = {}, {}
            day_map = {"一":1,"二":2,"三":3,"四":4,"五":5,"週一":1,"週二":2,"週三":3,"週四":4,"週五":5}
            for _, row in df_time.iterrows():
                c, s, d_str = str(row['班級']).strip(), str(row['科目']).strip(), str(row['星期']).strip()
                d = day_map.get(d_str, 0)
                p_match = re.search(r'\d+', str(row['節次']))
                if p_match and d > 0:
                    p = int(p_match.group())
                    ts = [x['t'] for x in assign_lookup if x['c'] == c and x['s'] == s]
                    t_disp = "/".join(ts)
                    # 班級視角
                    if c not in class_db: class_db[c] = {}
                    class_db[c][(d, p)] = {"s": s, "t": t_disp}
                    # 教師視角
                    for t in ts:
                        if t not in teacher_db: teacher_db[t] = {}
                        teacher_db[t][(d, p)] = {"s": s, "c": c}

            st.session_state.update({
                "class_data": class_db, "teacher_data": teacher_db,
                "ordered_teachers": sorted(list(all_teachers)), "data_ready": True
            })
            st.success("✅ 整合成功！")
            st.rerun()
        except Exception as e:
            st.error(f"解析失敗: {e}")

# --- 主畫面：功能分頁 ---

if st.session_state.get("data_ready"):
    tab1, tab2, tab3 = st.tabs(["🏫 班級課表", "👩‍🏫 教師課表", "📅 調代課管理"])

    with tab1:
        c_list = sorted(list(st.session_state.class_data.keys()))
        sel_c = st.selectbox("選擇班級", c_list)
        df_c = pd.DataFrame(index=range(1,9), columns=["週一","週二","週三","週四","週五"])
        for (d, p), val in st.session_state.class_data[sel_c].items():
            df_c.iloc[p-1, d-1] = f"{val['s']}\n{val['t']}"
        st.table(df_c.fillna(""))

    with tab2:
        t_list = st.session_state.ordered_teachers
        sel_t = st.selectbox("選擇教師", t_list)
        df_t = pd.DataFrame(index=range(1,9), columns=["週一","週二","週三","週四","週五"])
        for (d, p), val in st.session_state.teacher_data.get(sel_t, {}).items():
            df_t.iloc[p-1, d-1] = f"{val['c']}\n{val['s']}"
        st.table(df_t.fillna(""))

    with tab3:
        st.header("🔄 產製調代課通知單")
        c1, c2 = st.columns(2)
        with c1:
            date_val = st.date_input("代課日期", datetime.now())
            w_idx = date_val.weekday() + 1
            w_dates = get_week_dates(date_val)
        with c2:
            absent_t = st.selectbox("請假老師", st.session_state.ordered_teachers)

        # 找出該老師當天的課
        lessons = []
        for (d, p), v in st.session_state.teacher_data.get(absent_t, {}).items():
            if d == w_idx:
                lessons.append({"p": p, "c": v['c'], "s": v['s'], "txt": f"第{p}節: {v['c']} {v['s']}"})
        
        if lessons:
            sel_l = st.radio("選擇要代課的節次", lessons, format_func=lambda x: x['txt'])
            
            # 過濾空堂老師
            avail_ts = [t for t in st.session_state.ordered_teachers if (w_idx, sel_l['p']) not in st.session_state.teacher_data.get(t, {})]
            sub_t = st.selectbox("選擇代課老師 (已過濾空堂)", avail_ts)
            
            if st.button("🖨️ 產生 Word 通知單"):
                if "sub_template" not in st.session_state:
                    st.error("找不到樣板檔，請確認 GitHub 有『代調課通知單.docx』")
                else:
                    # 構建換行內容 
                    change = {
                        'day': w_idx, 'period': sel_l['p'],
                        'text': f"代{sel_l['c']}\n{sel_l['s']}" 
                    }
                    out_doc = fill_sub_notice(st.session_state.sub_template, sub_t, [change], w_dates)
                    
                    buf = BytesIO()
                    out_doc.save(buf)
                    st.download_button(f"⬇️ 下載 {sub_t} 代課單", buf.getvalue(), f"{sub_t}_代課單.docx")
        else:
            st.warning("該位老師當天沒有課程。")
else:
    st.info("👋 請在左側上傳 Excel 檔案並點擊「執行資料整合」。")
