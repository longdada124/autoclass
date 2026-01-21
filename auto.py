import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from io import BytesIO
import re
import requests
from datetime import datetime, timedelta

# --- 1. 從 GitHub 自動抓取樣板 ---
GITHUB_TEMPLATE_URL = "https://raw.githubusercontent.com/longdada124/autoclass/main/%E4%BB%A3%E8%AA%BF%E8%AA%B2%E9%80%9A%E7%9F%A5%E5%96%AE%E6%A8%A3%E6%9D%BF.docx"

def get_remote_template():
    try:
        response = requests.get(GITHUB_TEMPLATE_URL)
        response.raise_for_status()
        return response.content
    except Exception as e:
        st.error(f"無法從 GitHub 取得樣板：{e}")
        return None

# --- 2. Word 格式控制核心 (標楷體 + 標籤清理) ---
def set_font_style(run, font_name="標楷體"):
    """確保輸出內容強制為標楷體"""
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def master_replace(doc_obj, old_text, new_text):
    """替換文字並套用標楷體 """
    new_val = str(new_text) if new_text is not None else ""
    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old_text in p.text:
                        for run in p.runs:
                            if old_text in run.text:
                                run.text = run.text.replace(old_text, new_val)
                                set_font_style(run)

# --- 3. 系統初始化與資料保存 ---
st.set_page_config(page_title="後龍國中智慧代調課系統", layout="wide")

if 'db_ready' not in st.session_state: st.session_state.db_ready = False
if 'selected_lesson' not in st.session_state: st.session_state.selected_lesson = None

with st.sidebar:
    st.header("📂 資料更新區")
    f_assign = st.file_uploader("1. 上傳配課表 (Excel)", type=["xlsx"])
    f_time = st.file_uploader("2. 上傳課表 (Excel)", type=["xlsx"])
    
    if f_assign and f_time:
        if st.button("🚀 更新資料庫"):
            try:
                # 處理 Excel 並加入防錯
                df_a = pd.read_excel(f_assign).astype(str).apply(lambda x: x.str.strip())
                df_t = pd.read_excel(f_time).astype(str).apply(lambda x: x.str.strip())
                
                t_db = {}
                all_t = set()
                day_map = {"一":1, "二":2, "三":3, "四":4, "五":5}
                
                for _, r in df_t.iterrows():
                    d_m = re.search(r'[一二三四五]', r['星期'])
                    p_m = re.search(r'\d+', r['節次'])
                    if d_m and p_m:
                        d, p = day_map[d_m.group()], int(p_m.group())
                        c, s = r['班級'], r['科目']
                        # 避免 IndexError：先檢查是否有匹配資料
                        match = df_a[(df_a['班級'] == c) & (df_a['科目'] == s)]
                        if not match.empty:
                            teachers = str(match.iloc[0]['教師']).split('/')
                            for t in [x.strip() for x in teachers]:
                                all_t.add(t)
                                if t not in t_db: t_db[t] = {}
                                t_db[t][(d, p)] = {"c": c, "s": s}
                
                # 同步抓取遠端樣板 
                template_data = get_remote_template()
                if template_data:
                    st.session_state.update({
                        "t_db": t_db, "all_t": sorted(list(all_t)), 
                        "template": template_data, "db_ready": True
                    })
                    st.success("✅ 遠端樣板與資料庫已就緒")
            except Exception as e:
                st.error(f"資料整合錯誤：{e}")

# --- 4. 主畫面：互動式操作 (模仿 DM 功能) ---
if st.session_state.db_ready:
    st.title("📑 智慧代調課作業系統")
    sel_teacher = st.selectbox("🔍 選擇請假教師", st.session_state.all_t)
    
    # 視覺化課表網格
    st.write(f"### 📅 {sel_teacher} 老師課表")
    cols = st.columns(5)
    for d_idx in range(5):
        with cols[d_idx]:
            st.button(f"週{['一','二','三','四','五'][d_idx]}", disabled=True, use_container_width=True)
            for p in range(1, 9):
                info = st.session_state.t_db.get(sel_teacher, {}).get((d_idx + 1, p))
                if info:
                    if st.button(f"第{p}節\n{info['c']}\n{info['s']}", key=f"b_{d_idx}_{p}", use_container_width=True, type="primary"):
                        st.session_state.selected_lesson = {'day': d_idx+1, 'period': p, 'c': info['c'], 's': info['s']}
                else:
                    st.button(f"第{p}節", key=f"e_{d_idx}_{p}", disabled=True, use_container_width=True)

    if st.session_state.selected_lesson:
        l = st.session_state.selected_lesson
        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            date_sel = st.date_input("🗓️ 變動日期", datetime.now())
            mode = st.radio("🔄 性質", ["代課", "調課"], horizontal=True)
        with c2:
            # 衝堂提示
            avail_ts = [t for t in st.session_state.all_t if (l['day'], l['period']) not in st.session_state.t_db.get(t, {})]
            to_t = st.selectbox("👤 接收教師 (已過濾衝堂)", avail_ts)
        
        if st.button("🚀 生成並下載通知單"):
            doc = Document(BytesIO(st.session_state.template))
            master_replace(doc, "{{TEACHER}}", to_t)
            
            # 計算該週日期並替換 D1~D5 [cite: 5, 6]
            mon = date_sel - timedelta(days=date_sel.weekday())
            for i in range(5):
                d_str = f"{mon.year-1911}.{(mon+timedelta(days=i)).month:02d}.{(mon+timedelta(days=i)).day:02d}"
                master_replace(doc, f"{{{{D{i+1}}}}}", d_str)
            
            # 填寫內容並清理剩餘標籤 
            target_tag = f"{{{{{l['day']}_{l['period']}}}}}"
            content = f"{mode[:1]}{l['c']}\n{l['s']}"
            for d in range(1, 6):
                for p in range(1, 9):
                    tag = f"{{{{{d}_{p}}}}}"
                    master_replace(doc, tag, content if tag == target_tag else "")
            
            buf = BytesIO(); doc.save(buf)
            st.download_button(f"⬇️ 下載 {to_t} 老師通知單", buf.getvalue(), f"通知單_{to_t}.docx")
else:
    st.info("👋 您好！系統已連線至 GitHub 樣板庫。請上傳 Excel 課表以開始作業。")
