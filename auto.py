import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from io import BytesIO
import re
import requests
from datetime import datetime, timedelta

# --- 1. GitHub 檔案路徑設定 ---
RAW_URL = "https://raw.githubusercontent.com/longdada124/autoclass/main/"
FILES = {
    "assign": "配課表.xlsx",
    "timetable": "課表.xlsx",
    "template": "代調課通知單樣板.docx"
}

# 增加 Cache 機制，提高加載速度並減少 GitHub API 調用
@st.cache_data(ttl=600)
def load_cloud_files():
    data = {}
    try:
        for key, filename in FILES.items():
            r = requests.get(RAW_URL + filename, timeout=10)
            r.raise_for_status()
            data[key] = r.content
        return data
    except Exception as e:
        st.error(f"❌ 雲端抓取失敗：{e}。請檢查 GitHub 檔名是否正確。")
        return None

# --- 2. Word 格式處理 (標楷體控制) ---
def set_font_style(run, font_name="標楷體"):
    """強制設定標楷體，解決輸出字體不一問題"""
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def safe_replace(doc, old_txt, new_txt):
    """安全替換標籤並套用格式"""
    val = str(new_txt) if new_txt else ""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old_txt in p.text:
                        for run in p.runs:
                            if old_txt in run.text:
                                run.text = run.text.replace(old_txt, val)
                                set_font_style(run)

# --- 3. 系統核心邏輯 ---
st.set_page_config(page_title="後龍國中教務系統", layout="wide")

# 初始化資料庫
if 'db' not in st.session_state:
    files = load_cloud_files()
    if files:
        try:
            # 讀取 Excel 並清理空白字元
            df_a = pd.read_excel(BytesIO(files['assign'])).astype(str).apply(lambda x: x.str.strip())
            df_t = pd.read_excel(BytesIO(files['timetable'])).astype(str).apply(lambda x: x.str.strip())
            
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
                    if cls not in c_db: c_db[cls] = {}
                    
                    # 解決 IndexError：檢查配課表是否有該班級科目
                    match = df_a[(df_a['班級'] == cls) & (df_a['科目'] == sub)]
                    if not match.empty:
                        ts = [x.strip() for x in str(match.iloc[0]['教師']).split('/')]
                        c_db[cls][(d, p)] = f"{sub}\n({', '.join(ts)})"
                        for t in ts:
                            all_t.add(t)
                            if t not in t_db: t_db[t] = {}
                            t_db[t][(d, p)] = {"c": cls, "s": sub}
            
            st.session_state.db = {
                "t_db": t_db, "c_db": c_db, 
                "all_t": sorted(list(all_t)), "all_c": sorted(list(all_c)),
                "template": files['template']
            }
        except Exception as e:
            st.error(f"📊 資料解析出錯：{e}")

# --- 4. 介面顯示 ---
if 'db' in st.session_state:
    db = st.session_state.db
    tab1, tab2, tab3 = st.tabs(["🏫 班級課表預覽", "👨‍🏫 教師課表預覽", "📝 代調課系統"])

    with tab1:
        c_sel = st.selectbox("請選擇班級", db['all_c'])
        df = pd.DataFrame(index=[f"第{i}節" for i in range(1, 9)], columns=["週一", "週二", "週三", "週四", "週五"])
        for d in range(1, 6):
            for p in range(1, 9):
                df.iloc[p-1, d-1] = db['c_db'].get(c_sel, {}).get((d, p), "")
        st.table(df)

    with tab2:
        t_sel = st.selectbox("請選擇教師", db['all_t'])
        df_t = pd.DataFrame(index=[f"第{i}節" for i in range(1, 9)], columns=["週一", "週二", "週三", "週四", "週五"])
        for d in range(1, 6):
            for p in range(1, 9):
                item = db['t_db'].get(t_sel, {}).get((d, p))
                df_t.iloc[p-1, d-1] = f"{item['c']} {item['s']}" if item else ""
        st.table(df_t)

    with tab3:
        st.subheader("智慧代調課作業")
        l_teacher = st.selectbox("1. 選擇請假教師", db['all_t'], key="lt")
        
        # 互動課表網格
        cols = st.columns(5)
        for d in range(5):
            with cols[d]:
                st.button(["一","二","三","四","五"][d], disabled=True, use_container_width=True)
                for p in range(1, 9):
                    info = db['t_db'].get(l_teacher, {}).get((d + 1, p))
                    if info:
                        if st.button(f"第{p}節\n{info['c']}\n{info['s']}", key=f"btn_{d}_{p}", use_container_width=True, type="primary"):
                            st.session_state.active = {'day': d+1, 'period': p, 'c': info['c'], 's': info['s']}
                    else:
                        st.button(f"第{p}節", key=f"emp_{d}_{p}", disabled=True, use_container_width=True)

        if 'active' in st.session_state:
            act = st.session_state.active
            st.divider()
            c1, c2 = st.columns(2)
            with c1:
                v_date = st.date_input("變動日期", datetime.now())
                v_mode = st.radio("性質", ["代課", "調課"], horizontal=True)
            with c2:
                # 智慧排除衝堂
                no_conflict = [t for t in db['all_t'] if (act['day'], act['period']) not in db['t_db'].get(t, {})]
                to_teacher = st.selectbox("2. 選擇接收教師 (已排除衝堂)", no_conflict)
            
            if st.button("🚀 生成通知單"):
                doc = Document(BytesIO(db['template']))
                safe_replace(doc, "{{TEACHER}}", to_teacher)
                
                # 更新日期 D1-D5 
                monday = v_date - timedelta(days=v_date.weekday())
                for i in range(5):
                    d_s = f"{monday.year-1911}.{(monday+timedelta(days=i)).month:02d}.{(monday+timedelta(days=i)).day:02d}"
                    safe_replace(doc, f"{{{{D{i+1}}}}}", d_s)
                
                # 填充格子與清理 [cite: 4, 6]
                target = f"{{{{{act['day']}_{act['period']}}}}}"
                content = f"{v_mode[:1]}{act['c']}\n{act['s']}"
                for d_ in range(1, 6):
                    for p_ in range(1, 9):
                        tag = f"{{{{{d_}_{p_}}}}}"
                        safe_replace(doc, tag, content if tag == target else "")
                
                output = BytesIO(); doc.save(output)
                st.download_button(f"⬇️ 下載 {to_teacher} 通知單", output.getvalue(), f"{to_teacher}_通知單.docx")
else:
    st.warning("⚠️ 系統正在讀取雲端資料，若長時間未出現請確認 GitHub 檔案是否存在且公開。")
