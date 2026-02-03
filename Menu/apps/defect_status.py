# -*- coding: utf-8 -*-
import streamlit as st
import base64
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# --- CONFIG ---
JIRA_DOMAIN    = "scbjira.atlassian.net"
JIRA_EMAIL     = "t_pattanaphon.onrodprai@scb.co.th"
JIRA_API_TOKEN = "ATATT3xFfGF02t9xxsl-kce6l3cj19ZFX6bm1H7fTlVrjs67b1MzjOlHuq0fROrEUWZopFFxRNF5oOyL7QABOERjMK4Oko6szJXPRCCNmuVZDBLjfyJ6364zGtBRYg27PhhIhrkOFQyf3VEHgEz1SWFbAyC8CIA24zCHV_19tWV6r4A3t0J-vW8=E5AB51E8"

MAX_WORKERS = 20
TIMEOUT = 30

sess = requests.Session()
retry = Retry(total=3, backoff_factor=1, status_forcelist=(429, 500, 502, 503, 504), allowed_methods=["GET"])
adapter = HTTPAdapter(max_retries=retry, pool_connections=50, pool_maxsize=50)
sess.mount("https://", adapter)

def _jira_headers():
    tok = base64.b64encode(f"{JIRA_EMAIL}:{JIRA_API_TOKEN}".encode()).decode()
    return {"Authorization": f"Basic {tok}", "Accept": "application/json"}

def fetch_single_issue_data(key):
    key = key.strip()
    if not key:
        return None, None, None, None
    url = f"https://{JIRA_DOMAIN}/rest/api/3/issue/{key}"
    try:
        r = sess.get(url, headers=_jira_headers(), params={"fields": "status,summary,assignee"}, timeout=TIMEOUT)
        if r.status_code == 200:
            fields = r.json().get("fields", {})
            status = fields.get("status", {}).get("name", "Unknown")
            summary = fields.get("summary", "")
            assignee_obj = fields.get("assignee")
            assignee = assignee_obj.get("displayName") or assignee_obj.get("emailAddress") or "Assigned" if assignee_obj else "Unassigned"
            return key, status, summary, assignee
        elif r.status_code == 404:
            return key, "Not Found", "Not Found", "Not Found"
        else:
            return key, f"Error {r.status_code}", f"Error {r.status_code}", f"Error {r.status_code}"
    except Exception:
        return key, "Error", "Error", "Error"


def render():
    """Render Defect Status page"""
    
    if "YOUR_JIRA" in JIRA_API_TOKEN or JIRA_API_TOKEN.strip() == "":
        st.error("⛔ กรุณาใส่ JIRA_API_TOKEN ใน Code ก่อนครับ")
        return
    
    input_keys = st.text_area(
        "Issue Keys",
        height=120,
        placeholder="วาง Issue Key ที่นี่ (หนึ่ง Key ต่อบรรทัด)...",
        label_visibility="collapsed"
    )
    
    if st.button("🚀 ตรวจสอบ", type="primary"):
        keys = [k.strip() for k in input_keys.split('\n') if k.strip()]
        
        if not keys:
            st.warning("⚠️ กรุณาใส่ Issue Key ก่อนครับ")
        else:
            results_map = {}
            with st.spinner("กำลังดึงข้อมูล..."):
                with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                    future_to_key = {executor.submit(fetch_single_issue_data, key): key for key in keys}
                    for future in as_completed(future_to_key):
                        key, status, summary, assignee = future.result()
                        if key:
                            results_map[key] = {"status": status, "summary": summary, "assignee": assignee}
            
            final_keys, final_statuses, final_summaries, final_assignees = [], [], [], []
            for k in keys:
                final_keys.append(k)
                data = results_map.get(k, {"status": "Error", "summary": "Error", "assignee": "Error"})
                final_statuses.append(data["status"])
                final_summaries.append(data["summary"])
                final_assignees.append(data["assignee"])
            
            st.markdown(f"### 📊 ผลลัพธ์ ({len(keys)} รายการ)")
            
            col1, col2, col3, col4 = st.columns([1, 3, 1, 1.5])
            with col1:
                st.markdown("**🔑 Key**")
                st.code("\n".join(final_keys), language="text")
            with col2:
                st.markdown("**📝 Summary**")
                st.code("\n".join(final_summaries), language="text")
            with col3:
                st.markdown("**📌 Status**")
                st.code("\n".join(final_statuses), language="text")
            with col4:
                st.markdown("**👤 Assignee**")
                st.code("\n".join(final_assignees), language="text")
