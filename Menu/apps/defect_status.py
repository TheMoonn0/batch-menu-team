# -*- coding: utf-8 -*-
import streamlit as st
import base64
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# -------------------------
# CONFIG (อ่านจาก Streamlit Secrets)
# -------------------------
# ตั้งค่าใน Streamlit Cloud > App settings > Secrets (TOML)
# JIRA_API_TOKEN = "xxxxx"
# JIRA_EMAIL = "xxxxx@xxx.com"
# JIRA_DOMAIN = "xxxx.atlassian.net"

def _secret(key: str, default: str = "") -> str:
    """อ่านค่า secrets แบบกันพัง (มี default)"""
    try:
        return str(st.secrets.get(key, default)).strip()
    except Exception:
        return default

JIRA_DOMAIN = _secret("JIRA_DOMAIN")
JIRA_EMAIL  = _secret("JIRA_EMAIL")

MAX_WORKERS = 20
TIMEOUT = 30

# -------------------------
# HTTP session + retry
# -------------------------
sess = requests.Session()
retry = Retry(
    total=3,
    backoff_factor=1,
    status_forcelist=(429, 500, 502, 503, 504),
    allowed_methods=["GET"],
)
adapter = HTTPAdapter(max_retries=retry, pool_connections=50, pool_maxsize=50)
sess.mount("https://", adapter)

# -------------------------
# JIRA helpers
# -------------------------
def _get_jira_token() -> str:
    return _secret("JIRA_API_TOKEN")

def _validate_config() -> str:
    """คืนข้อความ error ถ้าค่าที่จำเป็นไม่ครบ"""
    missing = []
    if not JIRA_DOMAIN:
        missing.append("JIRA_DOMAIN")
    if not JIRA_EMAIL:
        missing.append("JIRA_EMAIL")
    if not _get_jira_token():
        missing.append("JIRA_API_TOKEN")
    if missing:
        return "ขาดค่าใน Streamlit Secrets: " + ", ".join(missing)
    return ""

def _jira_headers():
    jira_token = _get_jira_token()
    basic = base64.b64encode(f"{JIRA_EMAIL}:{jira_token}".encode("utf-8")).decode("utf-8")
    return {
        "Authorization": f"Basic {basic}",
        "Accept": "application/json",
    }

def fetch_single_issue_data(key: str):
    key = (key or "").strip()
    if not key:
        return None, None, None, None

    url = f"https://{JIRA_DOMAIN}/rest/api/3/issue/{key}"
    try:
        r = sess.get(
            url,
            headers=_jira_headers(),
            params={"fields": "status,summary,assignee"},
            timeout=TIMEOUT,
        )

        if r.status_code == 200:
            fields = r.json().get("fields", {})
            status = fields.get("status", {}).get("name", "Unknown")
            summary = fields.get("summary", "") or ""
            assignee_obj = fields.get("assignee")
            if assignee_obj:
                assignee = assignee_obj.get("displayName") or assignee_obj.get("emailAddress") or "Assigned"
            else:
                assignee = "Unassigned"
            return key, status, summary, assignee

        if r.status_code == 404:
            return key, "Not Found", "Not Found", "Not Found"

        return key, f"Error {r.status_code}", f"Error {r.status_code}", f"Error {r.status_code}"

    except Exception:
        return key, "Error", "Error", "Error"

# -------------------------
# UI
# -------------------------
def render():
    """Render Defect Status page"""

    err = _validate_config()
    if err:
        st.error(f"⛔ {err}\n\nไปที่ App settings → Secrets แล้วใส่ค่าให้ครบ")
        return

    st.subheader("Defect Status Checker")

    input_keys = st.text_area(
        "Issue Keys",
        height=120,
        placeholder="วาง Issue Key ที่นี่ (หนึ่ง Key ต่อบรรทัด)...",
        label_visibility="collapsed",
    )

    if st.button("🚀 ตรวจสอบ", type="primary"):
        keys = [k.strip() for k in (input_keys or "").split("\n") if k.strip()]

        if not keys:
            st.warning("⚠️ กรุณาใส่ Issue Key ก่อนครับ")
            return

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


# ให้รันได้ทั้งแบบ import และแบบ run ตรง ๆ
if __name__ == "__main__":
    render()
