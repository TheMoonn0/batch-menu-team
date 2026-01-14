# -*- coding: utf-8 -*-
import streamlit as st

# Import apps
from apps import gen_gl, gen_database, database, defect_status

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Menu",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============================================================
# CUSTOM CSS - Beautiful Animations + Dark Theme
# ============================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    
    * { font-family: 'Inter', sans-serif; }
    
    #MainMenu, footer { visibility: hidden; }
    
    /* Centered main content */
    .main .block-container {
        max-width: 1000px;
        margin: 0 auto;
        padding: 2rem;
    }
    
    /* ============================================ */
    /* KEYFRAME ANIMATIONS */
    /* ============================================ */
    @keyframes fadeInUp {
        from { 
            opacity: 0; 
            transform: translateY(30px); 
        }
        to { 
            opacity: 1; 
            transform: translateY(0); 
        }
    }
    
    @keyframes fadeInDown {
        from { 
            opacity: 0; 
            transform: translateY(-20px); 
        }
        to { 
            opacity: 1; 
            transform: translateY(0); 
        }
    }
    
    @keyframes fadeInLeft {
        from { 
            opacity: 0; 
            transform: translateX(-30px); 
        }
        to { 
            opacity: 1; 
            transform: translateX(0); 
        }
    }
    
    @keyframes fadeInRight {
        from { 
            opacity: 0; 
            transform: translateX(30px); 
        }
        to { 
            opacity: 1; 
            transform: translateX(0); 
        }
    }
    
    @keyframes scaleIn {
        from { 
            opacity: 0; 
            transform: scale(0.9); 
        }
        to { 
            opacity: 1; 
            transform: scale(1); 
        }
    }
    
    @keyframes shimmer {
        0% { background-position: -200% 0; }
        100% { background-position: 200% 0; }
    }
    
    @keyframes pulse {
        0%, 100% { transform: scale(1); }
        50% { transform: scale(1.05); }
    }
    
    @keyframes float {
        0%, 100% { transform: translateY(0px); }
        50% { transform: translateY(-8px); }
    }
    
    @keyframes glow {
        0%, 100% { box-shadow: 0 0 5px rgba(88, 166, 255, 0.3); }
        50% { box-shadow: 0 0 20px rgba(88, 166, 255, 0.6); }
    }
    
    @keyframes slideInFromLeft {
        0% { transform: translateX(-100%); opacity: 0; }
        100% { transform: translateX(0); opacity: 1; }
    }
    
    @keyframes bounce {
        0%, 20%, 50%, 80%, 100% { transform: translateY(0); }
        40% { transform: translateY(-10px); }
        60% { transform: translateY(-5px); }
    }
    
    @keyframes rotateIn {
        from { 
            opacity: 0; 
            transform: rotate(-10deg) scale(0.9); 
        }
        to { 
            opacity: 1; 
            transform: rotate(0) scale(1); 
        }
    }
    
    @keyframes gradientShift {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    
    /* ============================================ */
    /* SIDEBAR - Animated */
    /* ============================================ */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0d1117 0%, #161b22 50%, #1a1f26 100%) !important;
        border-right: 1px solid rgba(88, 166, 255, 0.1);
        animation: slideInFromLeft 0.4s ease-out;
    }
    
    [data-testid="stSidebar"] > div:first-child {
        background: transparent !important;
        padding-top: 1rem;
    }
    
    /* Sidebar buttons with stagger animation */
    [data-testid="stSidebar"] .stButton > button {
        background: rgba(255,255,255,0.03);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 10px;
        margin: 0.3rem 0;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        animation: fadeInLeft 0.5s ease-out backwards;
    }
    
    [data-testid="stSidebar"] .stButton:nth-child(1) > button { animation-delay: 0.1s; }
    [data-testid="stSidebar"] .stButton:nth-child(2) > button { animation-delay: 0.15s; }
    [data-testid="stSidebar"] .stButton:nth-child(3) > button { animation-delay: 0.2s; }
    [data-testid="stSidebar"] .stButton:nth-child(4) > button { animation-delay: 0.25s; }
    [data-testid="stSidebar"] .stButton:nth-child(5) > button { animation-delay: 0.3s; }
    [data-testid="stSidebar"] .stButton:nth-child(6) > button { animation-delay: 0.35s; }
    
    [data-testid="stSidebar"] .stButton > button:hover {
        background: linear-gradient(135deg, rgba(88, 166, 255, 0.15), rgba(163, 113, 247, 0.1));
        border-color: rgba(88, 166, 255, 0.4);
        transform: translateX(8px);
        box-shadow: 0 4px 15px rgba(88, 166, 255, 0.2);
    }
    
    /* ============================================ */
    /* HOME PAGE */
    /* ============================================ */
    .home-header {
        text-align: center;
        padding: 2rem 0 3rem 0;
        animation: fadeInDown 0.6s ease-out;
    }
    
    .home-icon {
        font-size: 4rem;
        display: block;
        margin-bottom: 0.5rem;
        animation: float 3s ease-in-out infinite;
    }
    
    .home-title {
        font-size: 3.5rem;
        font-weight: 800;
        background: linear-gradient(135deg, #58a6ff, #a371f7, #f778ba, #58a6ff);
        background-size: 300% 300%;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        animation: gradientShift 4s ease infinite;
        display: inline;
    }
    
    .home-subtitle {
        color: #8b949e;
        font-size: 1.2rem;
        margin-top: 1rem;
        animation: fadeInUp 0.8s ease-out 0.2s backwards;
    }
    
    /* ============================================ */
    /* PAGE TITLE */
    /* ============================================ */
    .page-title {
        color: #c9d1d9;
        font-size: 1.75rem;
        font-weight: 700;
        margin-bottom: 1.5rem;
        padding-bottom: 0.75rem;
        border-bottom: 2px solid transparent;
        border-image: linear-gradient(90deg, #58a6ff, #a371f7, transparent) 1;
        animation: fadeInLeft 0.5s ease-out;
    }
    
    /* ============================================ */
    /* MENU CARDS - Home Page */
    /* ============================================ */
    .main .stButton > button {
        min-height: 110px;
        background: linear-gradient(145deg, rgba(22, 27, 34, 0.95), rgba(13, 17, 23, 0.98));
        border: 1px solid rgba(255, 255, 255, 0.06);
        border-radius: 16px;
        transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
        animation: scaleIn 0.5s ease-out backwards;
        position: relative;
        overflow: hidden;
    }
    
    /* Stagger animation for cards */
    .main .stButton:nth-of-type(1) > button { animation-delay: 0.1s; }
    .main .stButton:nth-of-type(2) > button { animation-delay: 0.2s; }
    .main .stButton:nth-of-type(3) > button { animation-delay: 0.3s; }
    .main .stButton:nth-of-type(4) > button { animation-delay: 0.4s; }
    
    /* Card hover shine effect */
    .main .stButton > button::before {
        content: '';
        position: absolute;
        top: 0;
        left: -100%;
        width: 100%;
        height: 100%;
        background: linear-gradient(
            90deg, 
            transparent, 
            rgba(255, 255, 255, 0.1), 
            transparent
        );
        transition: left 0.6s ease;
    }
    
    .main .stButton > button:hover::before {
        left: 100%;
    }
    
    .main .stButton > button:hover {
        transform: translateY(-8px) scale(1.02);
        border-color: rgba(88, 166, 255, 0.5);
        box-shadow: 
            0 20px 40px rgba(88, 166, 255, 0.15),
            0 0 30px rgba(88, 166, 255, 0.1),
            inset 0 1px 0 rgba(255, 255, 255, 0.1);
    }
    
    .main .stButton > button:active {
        transform: translateY(-4px) scale(0.98);
    }
    
    /* ============================================ */
    /* PRIMARY BUTTON (Run, Process, etc) */
    /* ============================================ */
    .stButton > button[kind="primary"] {
        min-height: auto !important;
        background: linear-gradient(135deg, #238636, #2ea043, #3fb950);
        background-size: 200% 200%;
        border: none;
        border-radius: 10px;
        animation: fadeInUp 0.4s ease-out, gradientShift 3s ease infinite;
    }
    
    .stButton > button[kind="primary"]:hover {
        transform: translateY(-3px);
        box-shadow: 
            0 10px 30px rgba(46, 160, 67, 0.4),
            0 0 20px rgba(46, 160, 67, 0.3);
    }
    
    .stButton > button[kind="primary"]:active {
        transform: translateY(-1px);
    }
    
    /* ============================================ */
    /* FORM ELEMENTS */
    /* ============================================ */
    .stTextArea textarea, 
    .stTextInput input {
        background: rgba(13, 17, 23, 0.8);
        border: 1px solid #30363d;
        border-radius: 10px;
        color: #c9d1d9;
        transition: all 0.3s ease;
        animation: fadeInUp 0.5s ease-out;
    }
    
    .stTextArea textarea:focus, 
    .stTextInput input:focus {
        border-color: #58a6ff;
        box-shadow: 
            0 0 0 3px rgba(88, 166, 255, 0.15),
            0 0 20px rgba(88, 166, 255, 0.1);
        animation: glow 2s ease-in-out infinite;
    }
    
    /* ============================================ */
    /* FILE UPLOADER */
    /* ============================================ */
    .stFileUploader {
        animation: fadeInUp 0.5s ease-out;
    }
    
    .stFileUploader > div {
        border-radius: 12px;
        border: 2px dashed rgba(88, 166, 255, 0.3);
        transition: all 0.3s ease;
    }
    
    .stFileUploader > div:hover {
        border-color: rgba(88, 166, 255, 0.6);
        background: rgba(88, 166, 255, 0.05);
    }
    
    /* ============================================ */
    /* ALERTS & INFO BOXES */
    /* ============================================ */
    .stAlert {
        border-radius: 12px;
        animation: fadeInUp 0.5s ease-out;
        border-left: 4px solid;
    }
    
    .stSuccess {
        animation: pulse 0.5s ease-out;
    }
    
    /* ============================================ */
    /* CODE BLOCKS */
    /* ============================================ */
    .stCodeBlock {
        border-radius: 12px;
        animation: fadeInUp 0.5s ease-out;
    }
    
    /* ============================================ */
    /* DOWNLOAD BUTTON */
    /* ============================================ */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #1f6feb, #388bfd);
        border: none;
        border-radius: 10px;
        animation: fadeInUp 0.5s ease-out, pulse 2s ease-in-out infinite;
    }
    
    .stDownloadButton > button:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 30px rgba(31, 111, 235, 0.4);
        animation: none;
    }
    
    /* ============================================ */
    /* SPINNER */
    /* ============================================ */
    .stSpinner > div {
        animation: rotateIn 0.5s ease-out;
    }
    
    /* ============================================ */
    /* EXPANDER */
    /* ============================================ */
    .streamlit-expanderHeader {
        border-radius: 10px;
        transition: all 0.3s ease;
        animation: fadeInUp 0.5s ease-out;
    }
    
    .streamlit-expanderHeader:hover {
        background: rgba(88, 166, 255, 0.1);
    }
    
    /* ============================================ */
    /* DATAFRAME */
    /* ============================================ */
    .stDataFrame {
        animation: fadeInUp 0.6s ease-out;
        border-radius: 12px;
        overflow: hidden;
    }
    
    /* ============================================ */
    /* RADIO BUTTONS */
    /* ============================================ */
    .stRadio > div {
        animation: fadeInUp 0.5s ease-out;
    }
    
    .stRadio label {
        transition: all 0.2s ease;
    }
    
    .stRadio label:hover {
        color: #58a6ff;
    }
    
    /* ============================================ */
    /* CHECKBOX */
    /* ============================================ */
    .stCheckbox {
        animation: fadeInUp 0.5s ease-out;
    }
    
    .stCheckbox label {
        transition: all 0.2s ease;
    }
    
    .stCheckbox label:hover {
        color: #58a6ff;
    }
    
    /* ============================================ */
    /* METRIC */
    /* ============================================ */
    [data-testid="stMetricValue"] {
        animation: fadeInUp 0.5s ease-out;
    }
    
    /* ============================================ */
    /* DIVIDER */
    /* ============================================ */
    hr {
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(88, 166, 255, 0.3), transparent);
        animation: fadeInUp 0.5s ease-out;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# SESSION STATE
# ============================================================
if 'current_page' not in st.session_state:
    st.session_state.current_page = "🏠 Home"

# ============================================================
# SIDEBAR NAVIGATION
# ============================================================
with st.sidebar:
    st.markdown("### 📋 เมนู")
    
    if st.button("🏠 Home", key="nav_home", use_container_width=True):
        st.session_state.current_page = "🏠 Home"
        st.rerun()
    
    st.markdown("---")
    
    if st.button("📄 GEN File GL", key="nav_gl", use_container_width=True):
        st.session_state.current_page = "📄 GEN File GL"
        st.rerun()
    
    if st.button("💾 GEN File Database", key="nav_gendb", use_container_width=True):
        st.session_state.current_page = "💾 GEN File Database"
        st.rerun()
    
    if st.button("🗄️ Database", key="nav_db", use_container_width=True):
        st.session_state.current_page = "🗄️ Database"
        st.rerun()
    
    if st.button("🐛 Defect Status", key="nav_defect", use_container_width=True):
        st.session_state.current_page = "🐛 Defect Status"
        st.rerun()

# ============================================================
# PAGE: Home
# ============================================================
def page_home():
    st.markdown("""
        <div class="home-header">
            <span class="home-icon">🏠</span>
            <h1 class="home-title">Home</h1>
            <p class="home-subtitle">เลือกเครื่องมือที่ต้องการใช้งาน</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Centered grid
    col_spacer1, col1, col2, col_spacer2 = st.columns([0.5, 2, 2, 0.5])
    
    with col1:
        if st.button("📄\n\n**GEN File GL**", key="card_gl", use_container_width=True):
            st.session_state.current_page = "📄 GEN File GL"
            st.rerun()
        
        st.write("")
        
        if st.button("🗄️\n\n**Database**", key="card_db", use_container_width=True):
            st.session_state.current_page = "🗄️ Database"
            st.rerun()
    
    with col2:
        if st.button("💾\n\n**GEN File Database**", key="card_gendb", use_container_width=True):
            st.session_state.current_page = "💾 GEN File Database"
            st.rerun()
        
        st.write("")
        
        if st.button("🐛\n\n**Defect Status**", key="card_defect", use_container_width=True):
            st.session_state.current_page = "🐛 Defect Status"
            st.rerun()

# ============================================================
# ROUTING
# ============================================================
if st.session_state.current_page == "🏠 Home":
    page_home()
elif st.session_state.current_page == "📄 GEN File GL":
    st.markdown('<div class="page-title">📄 GEN File GL</div>', unsafe_allow_html=True)
    gen_gl.render()
elif st.session_state.current_page == "💾 GEN File Database":
    st.markdown('<div class="page-title">💾 GEN File Database</div>', unsafe_allow_html=True)
    gen_database.render()
elif st.session_state.current_page == "🗄️ Database":
    st.markdown('<div class="page-title">🗄️ Database</div>', unsafe_allow_html=True)
    database.render()
elif st.session_state.current_page == "🐛 Defect Status":
    st.markdown('<div class="page-title">🐛 Defect Status</div>', unsafe_allow_html=True)
    defect_status.render()
