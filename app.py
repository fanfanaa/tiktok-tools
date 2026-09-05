import streamlit as st

from config import APP_TITLE
from services.auth_service import init_auth_session, render_login_sidebar, render_sidebar_history

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container { max-width: 1180px; padding-top: 1.2rem; padding-bottom: 3rem; }
    h1 { font-size: 1.9rem !important; margin-bottom: .8rem !important; }
    h2 { font-size: 1.35rem !important; }
    h3 { font-size: 1.10rem !important; margin-top: 1rem !important; margin-bottom: .55rem !important; }
    div[data-testid="stButton"] button, div[data-testid="stDownloadButton"] button {
        border-radius: 8px; font-weight: 600; min-height: 42px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

init_auth_session()
render_login_sidebar()
st.title(APP_TITLE)

if not st.session_state.get("authenticated"):
    st.info("请先在左侧登录。")
    st.stop()

render_sidebar_history()

pages = [
    st.Page("pages/sop1_breakdown.py", title="爆款拆解", icon="🎬", default=True),
    st.Page("pages/sop2_compare.py", title="爆款对比", icon="🆚"),
    st.Page("pages/review.py", title="数据复盘", icon="📊"),
    st.Page("pages/history.py", title="历史记录", icon="🗂️"),
]

pg = st.navigation(pages, position="top")
pg.run()
