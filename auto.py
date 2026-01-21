import streamlit as st

# --- 模擬拖曳邏輯 ---
st.title("💡 視覺化代課排程 (模擬拖曳)")

# 1. 建立 Session State 存儲「抓取」中的對象
if "dragging_teacher" not in st.session_state:
    st.session_state.dragging_teacher = None
if "schedule_data" not in st.session_state:
    st.session_state.schedule_data = {} # 格式: {(day, period): "教師名稱"}

# --- 第一步：選取區域 (Source) ---
st.subheader("🏃 第一步：選取請假/代課人員")
teachers = ["王小明", "李大華", "張老師", "林老師"]
cols = st.columns(len(teachers))

for i, t in enumerate(teachers):
    with cols[i]:
        # 如果這個老師被選中了，按鈕顏色會變深 (type="primary")
        is_selected = (st.session_state.dragging_teacher == t)
        if st.button(t, key=f"src_{t}", type="primary" if is_selected else "secondary", use_container_width=True):
            st.session_state.dragging_teacher = t
            st.rerun()

if st.session_state.dragging_teacher:
    st.info(f"✨ 目前已抓取：**{st.session_state.dragging_teacher}**。請點擊下方課表位置進行「投放」。")
else:
    st.write("請先點擊上方一位教師。")

# --- 第二步：投放區域 (Target) ---
st.subheader("📅 第二步：投放到課表位置")

days = ["週一", "週二", "週三", "週四", "週五"]
grid_cols = st.columns(5)

for d_idx, day_name in enumerate(days):
    with grid_cols[d_idx]:
        st.button(day_name, disabled=True, use_container_width=True)
        for p in range(1, 9):
            pos = (d_idx + 1, p)
            # 取得該格子目前的內容
            current_val = st.session_state.schedule_data.get(pos, "＋")
            
            if st.button(current_val, key=f"cell_{d_idx}_{p}", use_container_width=True):
                if st.session_state.dragging_teacher:
                    # 執行「投放」動作
                    st.session_state.schedule_data[pos] = st.session_state.dragging_teacher
                    # 投放完畢後清除抓取狀態 (或保留，看您的習慣)
                    # st.session_state.dragging_teacher = None 
                    st.rerun()

# --- 清除按鈕 ---
if st.button("🗑️ 清空重新排程"):
    st.session_state.schedule_data = {}
    st.session_state.dragging_teacher = None
    st.rerun()
