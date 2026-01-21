import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from io import BytesIO
import re
import requests
from datetime import datetime, timedelta

# --- 1. 配置與遠端樣板讀取 ---
st.set_page_config(page_title="後龍國中智慧代調課系統", layout="wide")

# 直接連結您的 GitHub 樣板庫
GITHUB_URL = "https://raw.githubusercontent.com/longdada124/autoclass/main/%E4%BB%A3%E8%AA%BF%E8%AA%B2%E9%80%9A%E7%9F%A5%E5%96%AE%E6%A8%A3%E6%9D%BF.docx"

def get_template():
    try:
        resp = requests.get(GITHUB_URL)
        resp.raise_for_status()
        return resp.content
    except:
        st.error("⚠️ 無法連線至 GitHub 抓取樣板，請檢查網路或檔案路徑。")
        return None

# --- 2. Word 處理核心 (標楷體 + 標籤清理) ---
def set_font_style(run, font_name="標楷體"):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def master_replace(doc, old, new):
    val = str(new) if new else ""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old in p.text:
                        for run in p.runs:
                            if old in run.text:
                                run.text = run.text.replace(old, val)
                                set_font_style(run)

# --- 3. 資料庫初始化 (持久化儲存) ---
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
    st.session_state.t_db = {}     # 教師課表索引
    st.session_state.c_db = {}     # 班級課表索引
    st.session_state.all_t = []
    st.session_state.all_c = []

# 側邊欄：僅供管理員上傳資料庫
with st.sidebar:
    st.header("⚙️ 系統資料更新")
    f_assign = st.file_uploader("1. 更新配課表 (Excel)", type=["xlsx"])
    f_time = st.file_uploader("2. 更新全校課表 (Excel)", type=["xlsx"])
    
    if f_assign and f_time:
        if st.button("🔄 重新載入資料庫"):
            df_a = pd.read_excel(f_assign).astype(str).apply(lambda x: x.str.strip())
            df_t = pd.read_excel(f_time).astype(str).apply(lambda x: x.str.strip())
            
            t_db, c_db = {}, {}
            all_t, all_c = set(), set()
            day_map = {"一":1, "二":2, "三":3, "四":4, "五":5}

            for _, r in df_t.iterrows():
                d_m = re.search(r'[一二三四五]', r['星期'])
                p_m = re.search(r'\d+', r['節次'])
                if d_m and p_m:
                    d, p = day_map[d_m.group()], int(p_m.group())
                    cls, sub = r['班級'], r['科目']
                    all_c.add(cls)
                    
                    # 建立班級課表索引
                    if cls not in c_db: c_db[cls] = {}
                    
                    # 搜尋配課老師
                    match = df_a[(df_a['班級'] == cls) & (df_a['科目'] == sub)]
                    if not match.empty:
                        ts = [x.strip() for x in str(match.iloc[0]['教師']).split('/')]
                        c_db[cls][(d, p)] = f"{sub}\n({', '.join(ts)})"
                        for t in ts:
                            all_t.add(t)
                            if t not in t_db: t_db[t] = {}
                            t_db[t][(d, p)] = {"c": cls, "s": sub}
            
            st.session_state.update({
                "t_db": t_db, "c_db": c_db, "all_t": sorted(list(all_t)),
                "all_c": sorted(list(all_c)), "data_loaded": True,
                "template": get_template()
            })
            st.success("✅ 資料庫更新成功！")

# --- 4. 主介面分頁 ---
if st.session_state.data_loaded:
    tab1, tab2, tab3 = st.tabs(["🏫 班級課表", "👨‍🏫 教師課表", "📝 代調課作業"])

    with tab1:
        st.subheader("班級課表查詢")
        sel_c = st.selectbox("請選擇班級", st.session_state.all_c)
        df_view = pd.DataFrame(index=[f"第{i}節" for i in range(1, 9)], columns=["週一", "週二", "週三", "週四", "週五"])
        for d in range(1, 6):
            for p in range(1, 9):
                df_view.iloc[p-1, d-1] = st.session_state.c_db.get(sel_c, {}).get((d, p), "")
        st.table(df_view)

    with tab2:
        st.subheader("教師個人課表")
        sel_t = st.selectbox("請選擇教師", st.session_state.all_t, key="view_t")
        df_t_view = pd.DataFrame(index=[f"第{i}節" for i in range(1, 9)], columns=["週一", "週二", "週三", "週四", "週五"])
        for d in range(1, 6):
            for p in range(1, 9):
                item = st.session_state.t_db.get(sel_t, {}).get((d, p))
                df_t_view.iloc[p-1, d-1] = f"{item['c']}\n{item['s']}" if item else ""
        st.table(df_t_view)

    with tab3:
        st.subheader("智慧代調課管理")
        leave_t = st.selectbox("🔍 1. 選擇請假教師", st.session_state.all_t, key="leave_t")
        
        # 視覺化互動網格
        st.write("📌 請點擊欲變動之課程：")
        grid = st.columns(5)
        for d in range(5):
            with grid[d]:
                st.button(["週一","週二","週三","週四","週五"][d], disabled=True, use_container_width=True)
                for p in range(1, 9):
                    info = st.session_state.t_db.get(leave_t, {}).get((d + 1, p))
                    if info:
                        if st.button(f"第{p}節\n{info['c']}\n{info['s']}", key=f"job_{d}_{p}", use_container_width=True, type="primary"):
                            st.session_state.selected_lesson = {'day': d+1, 'period': p, 'c': info['c'], 's': info['s']}
                    else:
                        st.button(f"第{p}節", key=f"mt_{d}_{p}", disabled=True, use_container_width=True)

        if st.session_state.get('selected_lesson'):
            l = st.session_state.selected_lesson
            st.info(f"📍 已選取：週{l['day']} 第{l['period']}節 ({l['c']} {l['s']})")
            
            c1, c2 = st.columns(2)
            with c1:
                v_date = st.date_input("🗓️ 變動日期", datetime.now())
                v_mode = st.radio("🔄 性質", ["代課", "調課"], horizontal=True)
            with c2:
                # 智慧衝堂檢索
                avail = [t for t in st.session_state.all_t if (l['day'], l['period']) not in st.session_state.t_db.get(t, {})]
                to_t = st.selectbox("👤 2. 選擇接收教師 (自動過濾衝堂)", avail)
            
            if st.button("🚀 生成通知單"):
                doc = Document(BytesIO(st.session_state.template))
                master_replace(doc, "{{TEACHER}}", to_t)
                
                # 計算該週日期 D1-D5
                mon = v_date - timedelta(days=v_date.weekday())
                for i in range(5):
                    d_str = f"{mon.year-1911}.{(mon+timedelta(days=i)).month:02d}.{(mon+timedelta(days=i)).day:02d}"
                    master_replace(doc, f"{{{{D{i+1}}}}}", d_str)
                
                # 填寫選中格子，其餘清除
                tag_target = f"{{{{{l['day']}_{l['period']}}}}}"
                content = f"{v_mode[:1]}{l['c']}\n{l['s']}"
                for d_ in range(1, 6):
                    for p_ in range(1, 9):
                        tag = f"{{{{{d_}_{p_}}}}}"
                        master_replace(doc, tag, content if tag == tag_target else "")
                
                out = BytesIO(); doc.save(out)
                st.download_button(f"⬇️ 下載 {to_t} 的通知單", out.getvalue(), f"{to_t}_通知單.docx")
else:
    st.info("👋 歡迎使用！請先於左側側邊欄上傳「配課表」與「全校課表」Excel 以初始化系統。")
