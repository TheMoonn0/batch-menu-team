# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import duckdb
import uuid
import time

try:
    from streamlit_ace import st_ace
    ACE_AVAILABLE = True
except Exception:
    ACE_AVAILABLE = False


def render():
    """Render Database (SQL Query) page"""
    
    # =====================
    # SESSION STATE
    # =====================
    if "data_frame" not in st.session_state:
        st.session_state.data_frame = None

    if "query_pages" not in st.session_state:
        first_id = str(uuid.uuid4())
        st.session_state.query_pages = [{
            "id": first_id,
            "title": "Query 1",
            "query": "SELECT * FROM database",
            "last_result": None
        }]

    if "active_page_id" not in st.session_state:
        st.session_state.active_page_id = st.session_state.query_pages[0]["id"]

    if "page_selector_id" not in st.session_state:
        st.session_state.page_selector_id = st.session_state.active_page_id

    if "is_switching_page" not in st.session_state:
        st.session_state.is_switching_page = False

    if "run_id" not in st.session_state:
        st.session_state.run_id = 0

    # =====================
    # DATA LOADER
    # =====================
    @st.cache_data(show_spinner="Loading & Cleaning data...")
    def load_uploaded_file(uploaded_file):
        if uploaded_file.name.endswith(".parquet"):
            df = pd.read_parquet(uploaded_file).astype(str)
        else:
            df = pd.read_excel(uploaded_file, dtype=str)

        df = df.fillna("").replace("nan", "")
        df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
        df.columns = df.columns.str.strip().str.replace(" ", "_")
        return df

    # =====================
    # PAGE HELPERS
    # =====================
    def add_new_page():
        if len(st.session_state.query_pages) >= 10:
            st.toast("⚠️ สร้างได้สูงสุด 10 Pages เท่านั้น", icon="⚠️")
            return

        new_id = str(uuid.uuid4())
        st.session_state.query_pages.append({
            "id": new_id,
            "title": f"Query {len(st.session_state.query_pages) + 1}",
            "query": "SELECT * FROM database",
            "last_result": None
        })
        st.session_state.active_page_id = new_id
        st.session_state.is_switching_page = True

    def remove_page(page_id):
        pages = st.session_state.query_pages
        if len(pages) <= 1:
            return

        idx = next(i for i, p in enumerate(pages) if p["id"] == page_id)
        pages.pop(idx)

        if st.session_state.active_page_id == page_id:
            st.session_state.active_page_id = pages[max(0, idx - 1)]["id"]
            st.session_state.is_switching_page = True

    def _get_page_title_by_id(pid):
        for p in st.session_state.query_pages:
            if p["id"] == pid:
                return p["title"]
        return "Unknown"

    def _on_page_change():
        st.session_state.active_page_id = st.session_state.page_selector_id
        st.session_state.is_switching_page = True

    # =====================
    # MAIN CONTENT
    # =====================
    
    # Upload section
    if st.session_state.data_frame is None:
        uploaded = st.file_uploader("📂 Upload File (.xlsx / .parquet)", type=["xlsx", "parquet"])
        if uploaded:
            st.session_state.data_frame = load_uploaded_file(uploaded)
            st.rerun()
        return

    df = st.session_state.data_frame

    # Sidebar info
    with st.sidebar:
        st.markdown("---")
        if st.button("📤 Upload New File", use_container_width=True):
            st.session_state.data_frame = None
            st.session_state.query_pages = [{
                "id": str(uuid.uuid4()),
                "title": "Query 1",
                "query": "SELECT * FROM database",
                "last_result": None
            }]
            st.rerun()
        st.metric("📦 Total Rows", f"{len(df):,}")

    # Page selector
    page_ids = [p["id"] for p in st.session_state.query_pages]
    st.session_state.page_selector_id = st.session_state.active_page_id

    st.radio(
        "Pages",
        page_ids,
        horizontal=True,
        label_visibility="collapsed",
        key="page_selector_id",
        format_func=_get_page_title_by_id,
        on_change=_on_page_change,
    )

    page_id = st.session_state.active_page_id
    page = next(p for p in st.session_state.query_pages if p["id"] == page_id)

    # Header: rename + add + delete
    c1, c2, c3 = st.columns([0.65, 0.2, 0.15], vertical_alignment="bottom")

    with c1:
        new_title = st.text_input(
            "Page name",
            value=page["title"],
            label_visibility="collapsed",
            key=f"title_{page_id}",
        )
        if new_title != page["title"]:
            page["title"] = new_title
            st.rerun()

    with c2:
        if st.button("➕ เพิ่ม Page", use_container_width=True):
            add_new_page()
            st.rerun()

    with c3:
        if st.button("🗑️", use_container_width=True, disabled=len(st.session_state.query_pages) == 1):
            remove_page(page_id)
            st.rerun()

    st.caption("👆 เปลี่ยนชื่อ / เพิ่ม / ลบ Page ได้จากตรงนี้")

    # SQL Editor
    editor_key = f"q_{page_id}"
    if editor_key not in st.session_state:
        st.session_state[editor_key] = page["query"]

    with st.form(f"sql_form_{page_id}", clear_on_submit=False):
        if ACE_AVAILABLE:
            query = st_ace(
                value=st.session_state[editor_key],
                language="sql",
                theme="chrome",
                height=360,
                font_size=14,
                wrap=True,
                show_gutter=True,
                show_print_margin=False,
                auto_update=True,
                key=editor_key,
            )
        else:
            query = st.text_area(
                "SQL",
                key=editor_key,
                height=360,
                label_visibility="collapsed",
            )

        submitted = st.form_submit_button("▶️ Run SQL", type="primary")

    page["query"] = query or ""

    # Run SQL
    if submitted:
        with st.spinner("⏳ Running SQL..."):
            time.sleep(0.25)
            try:
                q = (page["query"] or "").strip().rstrip(";")
                if not q:
                    raise ValueError("SQL ว่างอยู่ครับ")

                if "limit" not in q.lower():
                    q += " LIMIT 300"

                result = duckdb.query(q.replace("database", "df")).to_df()
                page["last_result"] = result
                st.session_state.run_id += 1
                st.toast("✅ Results updated", icon="✅")
            except Exception as e:
                st.error(f"❌ SQL Error:\n{e}")
                page["last_result"] = None

    # Result
    if page["last_result"] is not None:
        res = page["last_result"]
        st.success(f"Query Success: {len(res):,} rows")

        with st.expander("📋 Copy Data for Excel (All rows in result)"):
            st.code(res.to_csv(index=False, sep="\t"), language="text")

        st.divider()
        st.subheader("Result Table")

        st.dataframe(
            res,
            use_container_width=True,
            key=f"result_{page_id}_{st.session_state.run_id}",
        )
