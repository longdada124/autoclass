import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
from datetime import datetime, timedelta

# --- 1. 基礎設定 ---
st.set_page_config(page_title="後龍國中全方位課務系統", layout="wide")

# --- 2. 核心工具函數 ---

def master_replace(doc_obj, old_text, new_text):
    """
    強力替換函數：
    1. 支援換行符號 \n
    2. 保留 Word 樣板原本的字體 (標楷體)
    3. 適用於段落與表格
    """
    if new_text is None: new_text = ""
    new_val = str(new_text)

    # 內部函數：處理單個 run 的替換
    def replace_run(run):
        if old_text in run.text:
            if "\n" in new_val:
                # 處理換行：切割文字 -> 插入換行符 -> 插入第二段
                parts = new_val.split("\n")
                # 替換掉標籤，換成第一行文字
                run.text = run.text.replace(old_text, parts[0])
                # 依序加入後面的文字
                for part in parts[1:]:
                    run.add_break() 
                    run.add_text(part)
            else:
                # 一般替換 (包含替換成空字串)
                run.text = run.text.replace(old_text, new_val)

    # 1. 掃描文件段落 (如抬頭、日期)
    for p in doc_obj.paragraphs:
        if old_text in p.text:
            for run in p.runs:
                replace_run(run)

    # 2. 掃描所有表格 (如課表格子)
    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old_text in p.text:
                        for run in p.runs:
                            replace_run(run)

def get_week_dates(base_date):
    """取得該週週一至週五的民國日期字串 (格式: 115.02.09)"""
    start_of_week = base_date - timedelta(days=base_date.weekday())
    dates = []
    for i in range(5):
        d = start_of_week + timedelta(days=i)
        roc_year = d.year - 1911
        dates.append(f"{roc_year}.{d.month:02d}.{d.day:02d}")
    return dates

def generate_doc(template_bytes, teacher_name, target_data, week_dates):
    """
    產製通知單主程序
    target_data: {'day': 1, 'period': 2, 'content': '代701\n國文'}
    """
    doc = Document(BytesIO(template_bytes))
    
    # 1. 填寫基本資料
    master_replace(doc, "{{TEACHER}}", teacher_name)
    for i, d_str in enumerate(week_dates):
        master_replace(doc, f"{{{{D{i+1}}}}}", d_str)
    
    # 2. 建立「要填寫」的標籤清單
    # 格式: {{1_2}} -> "代701\n國文"
    fill_map = {}
    target_tag = f"{{{{{target_data['day']}_{target_data['period']}}}}}"
    fill_map[target_tag] = target_data['content']
    
    # 3. 強力掃描：填入目標內容，並清空其餘所有格子
    # 假設一天最多9節，週一到週五
    for d in range(1, 6):
        for p in range(1, 10): 
            tag = f"{{{{{d}_{p}}}}}"
            if tag in fill_map:
                # 這是要填寫的格子
                master_replace(doc, tag, fill_map[tag])
            else:
                # 這是沒用到的格子，確實替換為「空字串」
                master_replace(doc, tag, "")
                
    return doc

# --- 3. 側邊欄：資料讀取 ---
with st.sidebar:
    st.header("📂 系統資料上傳")
    if st.button("🗑️ 清除資料重來"):
        st.session_state.clear()
        st.rerun()
    
    st.info("請依序上傳三個檔案以啟動系統")
    f_assign = st.file_uploader("1. 配課表 (xlsx/csv)", type=["xlsx", "csv"])
    f_time = st.file_uploader("2. 課表 (xlsx/csv)", type=["xlsx", "csv"])
    f_sort = st.file_uploader("3. 教師排序表 (xlsx/csv)", type=["xlsx", "csv"])
    
    if f_assign and f_time and st.button("🚀 執行資料整合"):
        try:
            with st.spinner("正在分析資料..."):
                # 讀取檔案
                df_assign = pd.read_excel(f_assign) if f_assign.name.endswith('xlsx') else pd.read_csv(f_assign)
                df_time = pd.read_excel(f_time) if f_time.name.endswith('xlsx') else pd.read_csv(f_time)
                
                # 嘗試讀取內建 Word 樣板 (需預先上傳到 GitHub)
                try:
                    with open("代調課通知單.docx", "rb") as f: 
                        st.session_state.sub_template = f.read()
                except:
                    st.warning("⚠️ 尚未找到【代調課通知單.docx】，請確認 GitHub 檔案是否存在。")

                # 解析配課表 (建立老師名單)
                assign_lookup = []
                all_teachers = set()
                for _, row in df_assign.iterrows():
                    c, s, t_raw = str(row['班級']).strip(), str(row['科目']).strip(), str(row['教師']).strip()
                    t_list = [t.strip() for t in t_raw.split('/') if t.strip() and t != "nan"]
                    for t in t_list:
                        assign_lookup.append({'c': c, 's': s, 't': t})
                        all_teachers.add(t)

                # 解析課表 (建立查詢索引)
                class_db = {}   # 班級視角
                teacher_db = {} # 老師視角
                day_map = {"一":1, "二":2, "三":3, "四":4, "五":5, "週一":1, "週二":2, "週三":3, "週四":4, "週五":5}
                
                for _, row in df_time.iterrows():
                    c = str(row['班級']).strip()
                    s = str(row['科目']).strip()
                    d_str = str(row['星期']).strip()
                    d = day_map.get(d_str, 0)
                    
                    # 抓取節次數字
                    p_match = re.search(r'\d+', str(row['節次']))
                    
                    if p_match and d > 0:
                        p = int(p_match.group())
                        # 找出這堂課的老師
                        matches = [x['t'] for x in assign_lookup if x['c'] == c and x['s'] == s]
                        t_disp = "/".join(matches) if matches else "未知"
                        
                        # 存入班級資料
                        if c not in class_db: class_db[c] = {}
                        class_db[c][(d, p)] = {"s": s, "t": t_disp}
                        
                        # 存入教師資料
                        for t in matches:
                            if t not in teacher_db: teacher_db[t] = {}
                            teacher_db[t][(d, p)] = {"c": c, "s": s}

                # 處理教師排序
                ordered_teachers = sorted(list(all_teachers))
                if f_sort:
                    try:
                        df_s = pd.read_excel(f_sort) if f_sort.name.endswith('xlsx') else pd.read_csv(f_sort)
                        s_list = [str(x).strip() for x in df_s.iloc[:,0].tolist()]
                        # 排序邏輯: 在清單內的優先，不在的放後面
                        ordered_teachers = [t for t in s_list if t in all_teachers] + [t for t in ordered_teachers if t not in s_list]
                    except: pass

                st.session_state.class_data = class_db
                st.session_state.teacher_data = teacher_db
                st.session_state.ordered_teachers = ordered_teachers
                st.session_state.data_ready = True
                
                st.success(f"✅ 資料整合完畢！共 {len(all_teachers)} 位教師。")
                st.rerun()

        except Exception as e:
            st.error(f"❌ 資料解析失敗: {e}")

# --- 4. 主介面邏輯 ---

if st.session_state.get("data_ready"):
    tab1, tab2, tab3 = st.tabs(["🏫 班級課表", "👩‍🏫 教師課表", "🔄 調代課通知單"])
    
    # === 分頁 1: 班級課表 ===
    with tab1:
        if st.session_state.class_data:
            c_list = sorted(list(st.session_state.class_data.keys()))
            sel_c = st.selectbox("請選擇班級", c_list)
            
            # 建立課表 Grid
            data_grid = {d: [""] * 8 for d in ["週一", "週二", "週三", "週四", "週五"]}
            for (d, p), info in st.session_state.class_data.get(sel_c, {}).items():
                if 1 <= p <= 8:
                    data_grid[list(data_grid.keys())[d-1]][p-1] = f"{info['s']}\n{info['t']}"
            
            df_display = pd.DataFrame(data_grid)
            df_display.index = [f"第{i}節" for i in range(1, 9)]
            st.table(df_display)
        else:
            st.info("尚無班級資料")

    # === 分頁 2: 教師課表 ===
    with tab2:
        sel_t = st.selectbox("請選擇教師", st.session_state.ordered_teachers)
        data_grid_t = {d: [""] * 8 for d in ["週一", "週二", "週三", "週四", "週五"]}
        
        info_map = st.session_state.teacher_data.get(sel_t, {})
        for (d, p), info in info_map.items():
            if 1 <= p <= 8:
                data_grid_t[list(data_grid_t.keys())[d-1]][p-1] = f"{info['c']}\n{info['s']}"
        
        df_display_t = pd.DataFrame(data_grid_t)
        df_display_t.index = [f"第{i}節" for i in range(1, 9)]
        st.table(df_display_t)

    # === 分頁 3: 調代課通知單 (重點功能) ===
    with tab3:
        st.markdown("### 步驟 1: 選擇原課程 (請假/調動方)")
        
        col1, col2 = st.columns(2)
        with col1:
            target_date = st.date_input("選擇日期", datetime.now())
            week_idx = target_date.weekday() + 1 # 1=Mon, 5=Fri
            week_dates = get_week_dates(target_date)
            st.caption(f"本週區間：{week_dates[0]} ~ {week_dates[4]}")
            
        with col2:
            orig_teacher = st.selectbox("原任課教師", st.session_state.ordered_teachers, index=0)

        # 搜尋該老師當日的課
        lessons = []
        t_schedule = st.session_state.teacher_data.get(orig_teacher, {})
        for p in range(1, 10):
            if (week_idx, p) in t_schedule:
                info = t_schedule[(week_idx, p)]
                lessons.append({
                    "p": p, 
                    "c": info['c'], 
                    "s": info['s'], 
                    "label": f"第 {p} 節 - {info['c']} {info['s']}"
                })
        
        if not lessons:
            st.warning(f"⚠️ {orig_teacher} 老師在 {target_date} 沒有課程。")
        else:
            # 選擇課程
            selected_lesson = st.radio("請勾選要處理的課程：", lessons, format_func=lambda x: x['label'])
            
            st.divider()
            st.markdown("### 步驟 2: 設定變動方式與接收教師")
            
            c3, c4 = st.columns(2)
            with c3:
                # 新增功能：選擇是代課還是調課
                change_type = st.radio("變動類型", ["代課 (Substitute)", "調課 (Swap)"], horizontal=True)
                type_prefix = "代" if "代課" in change_type else "調"
            
            with c4:
                # 智慧過濾：預設濾掉該節次已經有課的老師
                st.write("選擇新任課教師 (已過濾衝堂)")
                available_ts = []
                for t in st.session_state.ordered_teachers:
                    # 檢查該老師當天該節次是否有課
                    if (week_idx, selected_lesson['p']) not in st.session_state.teacher_data.get(t, {}):
                        available_ts.append(t)
                
                new_teacher = st.selectbox("新任課教師", available_ts)

            st.divider()
            
            # 預覽輸出結果
            preview_text = f"{type_prefix}{selected_lesson['c']}\n{selected_lesson['s']}"
            st.info(f"📄 預覽格子內容：\n\n{preview_text}\n\n(將填入 {new_teacher} 的通知單週{week_idx}第{selected_lesson['p']}節)")

            if st.button("🖨️ 產生 Word 通知單"):
                if "sub_template" not in st.session_state:
                    st.error("❌ 錯誤：找不到樣板檔，請確認已上傳【代調課通知單.docx】")
                else:
                    # 準備資料包
                    data_packet = {
                        'day': week_idx,
                        'period': selected_lesson['p'],
                        'content': preview_text
                    }
                    
                    # 呼叫產製函數
                    final_doc = generate_doc(
                        st.session_state.sub_template,
                        new_teacher,
                        data_packet,
                        week_dates
                    )
                    
                    # 存檔並提供下載
                    buf = BytesIO()
                    final_doc.save(buf)
                    file_name = f"{target_date.strftime('%m%d')}_{new_teacher}_通知單.docx"
                    
                    st.success("✅ 產製成功！")
                    st.download_button(
                        label=f"⬇️ 下載 {new_teacher} 的通知單",
                        data=buf.getvalue(),
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

else:
    st.info("👋 歡迎使用！請查看左側側邊欄，依序上傳資料以開始使用。")
