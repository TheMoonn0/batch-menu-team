# -*- coding: utf-8 -*-
import base64
import os
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import streamlit as st
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# --- CONFIG ---
JIRA_DOMAIN = "scbjira.atlassian.net"
JIRA_EMAIL = "t_pattanaphon.onrodprai@scb.co.th"

# --- PERFORMANCE TUNING ---
MAX_WORKERS = 20
TIMEOUT = 30

# --- SESSION SETUP ---
sess = requests.Session()
retry = Retry(
    total=3,
    backoff_factor=1,
    status_forcelist=(429, 500, 502, 503, 504),
    allowed_methods=["GET"],
)
adapter = HTTPAdapter(
    max_retries=retry,
    pool_connections=50,
    pool_maxsize=50,
)
sess.mount("https://", adapter)


def _get_jira_api_token():
    token = os.getenv("JIRA_API_TOKEN", "").strip()
    if token:
        return token

    try:
        secret_token = str(st.secrets.get("JIRA_API_TOKEN", "")).strip()
        if secret_token:
            return secret_token

        jira_config = st.secrets.get("jira", {})
        if jira_config:
            return str(jira_config.get("api_token", "")).strip()
    except Exception:
        return ""

    return ""


def _jira_headers():
    token = _get_jira_api_token()
    encoded_token = base64.b64encode(f"{JIRA_EMAIL}:{token}".encode()).decode()
    return {
        "Authorization": f"Basic {encoded_token}",
        "Accept": "application/json",
    }


def fetch_single_issue_data(key):
    key = key.strip()

    if not key:
        return None, None, None, None

    url = f"https://{JIRA_DOMAIN}/rest/api/3/issue/{key}"

    try:
        response = sess.get(
            url,
            headers=_jira_headers(),
            params={
                "fields": "summary,status",
                "expand": "changelog",
            },
            timeout=TIMEOUT,
        )

        if response.status_code == 200:
            data = response.json()
            fields = data.get("fields", {})

            summary = fields.get("summary", "")
            status = fields.get("status", {}).get("name", "Unknown")

            histories = data.get("changelog", {}).get("histories", [])
            reopen_count = 0

            for history in histories:
                for item in history.get("items", []):
                    if (item.get("field") or "").lower() == "status":
                        to_status = (item.get("toString") or "").strip().lower()
                        if to_status == "reopened":
                            reopen_count += 1

            return key, summary, status, reopen_count

        if response.status_code == 404:
            return key, "Not Found", "Not Found", 0

        error = f"Error {response.status_code}"
        return key, error, error, 0

    except Exception:
        return key, "Error", "Error", 0


def render():
    """Render Defect Status page."""

    st.markdown("### Jira Status + Reopen Counter")

    jira_api_token = _get_jira_api_token()
    if not jira_api_token:
        st.error("Set JIRA_API_TOKEN in Streamlit secrets or environment before running.")
        return

    c_input, c_btn = st.columns([3, 1], vertical_alignment="bottom")

    with c_input:
        input_keys = st.text_area(
            "Input",
            value="",
            height=150,
            label_visibility="collapsed",
            placeholder="Paste issue keys here, one per line...",
        )

    with c_btn:
        run_btn = st.button(
            "Check",
            use_container_width=True,
            type="primary",
        )

    if not run_btn:
        return

    keys = [key.strip() for key in input_keys.split("\n") if key.strip()]

    if not keys:
        st.warning("Enter at least one issue key.")
        return

    results_map = {}

    with st.spinner("Fetching Jira data..."):
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            future_to_key = {
                executor.submit(fetch_single_issue_data, key): key
                for key in keys
            }

            for future in as_completed(future_to_key):
                key, summary, status, reopen_count = future.result()

                if key:
                    results_map[key] = {
                        "summary": summary,
                        "status": status,
                        "reopen_count": reopen_count,
                    }

    final_keys = []
    final_summaries = []
    final_statuses = []
    final_reopen_counts = []

    for key in keys:
        data = results_map.get(
            key,
            {"summary": "Error", "status": "Error", "reopen_count": 0},
        )

        final_keys.append(key)
        final_summaries.append(data["summary"])
        final_statuses.append(data["status"])
        final_reopen_counts.append(str(data["reopen_count"]))

    st.divider()

    col1, col2, col3, col4 = st.columns([1, 4, 1, 1])

    with col1:
        st.caption("Keys")
        st.code("\n".join(final_keys), language="text")

    with col2:
        st.caption("Summary")
        st.code("\n".join(final_summaries), language="text")

    with col3:
        st.caption("Status")
        st.code("\n".join(final_statuses), language="text")

    with col4:
        st.caption("Reopen")
        st.code("\n".join(final_reopen_counts), language="text")
