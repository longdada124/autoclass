import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from io import BytesIO
import re
import requests
from datetime import datetime, timedelta

# --- 1. GitHub 雲端檔案配置 ---
# 請確認您的 GitHub 檔案名稱是否正確
RAW_URL_BASE = "https://raw.githubusercontent.com/longdada124/autoclass/main/"
FILES = {
    "assign": "配課表.xlsx", # 請確保 GitHub 上檔名一致
    "timetable": "課表.xlsx", 
    "template": "代調課通知單樣板.docx"
}

@st.cache_data(ttl=3600) # 快取一小時，避免頻繁抓取
def fetch_github_data():
    data = {}
    try:
        for key, name in FILES.items():
            url = RAW_URL_BASE + name
            resp = requests.get(url)
            resp.raise_for_status()
            data[key] = resp.content
        return data
    except Exception as e:
        st.error(f"❌ 無法從 GitHub 抓取資料，請檢查路徑。錯誤：{e}")
        return None

# --- 2. Word 格式核心 (強制標楷體) ---
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

# --- 3. 資料庫預載入邏輯 ---
st.set_page_config(page_title="後龍國中智慧教務系統", layout="wide")

cloud_data = fetch_github_data()

if cloud_data and 'initialized' not in st.session_state:
    # 解析 Excel
    df_a = pd.read_excel(BytesIO(cloud_data['assign'])).astype(str).apply(lambda x: x.str.strip())
    df_t = pd.read_excel(BytesIO(cloud_data['timetable'])).astype(str).apply(lambda x: x.str.strip())
    
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
            
            match = df_a[(df_a['班級'] == cls) & (df_a['科目'] == sub)]
            if not match.empty:
                ts = [x.strip() for x in str(match.iloc[0]['教師']).split('/')]
                c_db[cls][(d, p)] = f"{sub}\n({', '.join(ts)})"
                for t in ts:
                    all_t.add(t)
                    if t not in t_db: t_db[t] = {}
                    t_db[t][(d, p)] = {"c": cls, "s": sub}
    
    st.session_state.update({
        "t_db": t_db, "c_db": c_db, 
        "all_t": sorted(list(all_t)), "all_c": sorted(list(all_c)),
        "template": cloud_data['template'], "initialized": True
    })

# --- 4. 主介面預覽與功能 ---
if st.session_state.get("initialized"):
    tab1, tab2, tab3 = st.tabs(["🏫 班級課表預覽", "👨‍🏫 教師課表預覽", "📝 代調課系統"])

    with tab1:
        sel_c = st.selectbox("選擇班級", st.session_state.all_c)
        df_view = pd.DataFrame(index=[f"第{i}節" for i in range(1, 9)], columns=["週一", "週二", "週三", "週四", "週五"])
        for d in range(1, 6):
            for p in range(1, 9):
                df_view.iloc[p-1, d-1] = st.session_state.c_db.get(sel_c, {}).get((d, p), "")
        st.table(df_view)

    with tab2:
        sel_t = st.selectbox("選擇教師", st.session_state.all_t)
        df_t_view = pd.DataFrame(index=[f"第{i}節" for i in range(1, 9)], columns=["週一", "週二", "週三", "週四", "週五"])
        for d in range(1, 6):
            for p in range(1, 9):
                item = st.session_state.t_db.get(sel_t, {}).get((d, p))
                df_t_view.iloc[p-1, d-1] = f"{item['c']} {item['s']}" if item else ""
        st.table(df_t_view)

    with tab3:
        st.subheader("智慧代調課生成")
        leave_t = st.selectbox("1. 選擇請假教師", st.session_state.all_t, key="lt")
        
        # 課表點擊區
        st.caption("👇 點擊下方課程格子以選取：")
        grid = st.columns(5)
        for d in range(5):
            with grid[d]:
                st.button(["週一","週二","週三","週四","週五"][d], disabled=True, use_container_width=True)
                for p in range(1, 9):
                    info = st.session_state.t_db.get(leave_t, {}).get((d + 1, p))
                    if info:
                        if st.button(f"第{p}節\n{info['c']}\n{info['s']}", key=f"job_{d}_{p}", use_container_width=True, type="primary"):
                            st.session_state.selected = {'day': d+1, 'period': p, 'c': info['c'], 's': info['s']}
                    else:
                        st.button(f"第{p}節", key=f"mt_{d}_{p}", disabled=True, use_container_width=True)

        if st.session_state.get('selected'):
            l = st.session_state.selected
            st.divider()
            c1, c2 = st.columns(2)
            with c1:
                v_date = st.date_input("🗓️ 變動日期", datetime.now())
                v_mode = st.radio("🔄 性質", ["代課", "調課"], horizontal=True)
            with c2:
                # 智慧排除衝堂
                avail = [t for t in st.session_state.all_t if (l['day'], l['period']) not in st.session_state.t_db.get(t, {})]
                to_t = st.selectbox("👤 2. 選擇接收教師 (自動過濾衝堂)", avail)
            
            if st.button("🚀 生成代調課通知單", use_container_width=True):
                doc = Document(BytesIO(st.session_state.template))
                master_replace(doc, "{{TEACHER}}", to_t)
                
                # 日期標籤 (D1-D5)
                mon = v_date - timedelta(days=v_date.weekday())
                for i in range(5):
                    d_str = f"{mon.year-1911}.{(mon+timedelta(days=i)).month:02d}.{(mon+timedelta(days=i)).day:02d}"
                    master_replace(doc, f"{{{{D{i+1}}}}}", d_str)
                
                # 格子填充與清理 (標楷體)
                tag_target = f"{{{{{l['day']}_{l['period']}}}}}"
                content = f"{v_mode[:1]}{l['c']}\n{l['s']}"
                for d_ in range(1, 6):
                    for p_ in range(1, 9):
                        tag = f"{{{{{d_}_{p_}}}}}"
                        master_replace(doc, tag, content if tag == tag_target else "")
                
                out = BytesIO(); doc.save(out)
                st.download_button(f"⬇️ 下載 {to_t} 的通知單", out.getvalue(), f"{to_t}_代調課通知單.docx")

else:
    st.warning("🔄 正在從 GitHub 同步雲端資料庫，請稍候...")
