#!/usr/bin/env python3
"""
Cost Logger — Persistent API usage tracking
=============================================
Logs every API call (image audit, future tools) to a Google Sheet for
persistent cost tracking across sessions and deploys.

Falls back gracefully to session-only logging if Google Sheets isn't
configured — the sidebar display still works, you just lose persistence.

Setup:
  1. Create a Google Sheet with headers in row 1:
     Timestamp | Tool | Filename | Images | Input Tokens | Output Tokens | Cost USD | User
  2. Share the sheet with the service account email
  3. Add to Streamlit secrets (.streamlit/secrets.toml or Cloud dashboard):
     [gsheets]
     spreadsheet_url = "https://docs.google.com/spreadsheets/d/..."
     credentials = '{"type": "service_account", ...}'
"""

import json
from datetime import datetime

import streamlit as st

# ─── Google Sheets Persistence ────────────────────────────────────────

_SHEET_CLIENT = None
_SHEET_CONFIGURED = None  # None = not checked yet


def _get_sheet():
    """Get the Google Sheet worksheet, or None if not configured."""
    global _SHEET_CLIENT, _SHEET_CONFIGURED

    if _SHEET_CONFIGURED is False:
        return None
    if _SHEET_CLIENT is not None:
        return _SHEET_CLIENT

    try:
        import gspread
        from google.oauth2.service_account import Credentials

        creds_json = st.secrets["gsheets"]["credentials"]
        if isinstance(creds_json, str):
            creds_dict = json.loads(creds_json)
        else:
            creds_dict = dict(creds_json)

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        gc = gspread.authorize(creds)

        sheet_url = st.secrets["gsheets"]["spreadsheet_url"]
        sheet = gc.open_by_url(sheet_url).sheet1

        _SHEET_CLIENT = sheet
        _SHEET_CONFIGURED = True
        return sheet

    except Exception:
        _SHEET_CONFIGURED = False
        return None


def is_sheets_configured():
    """Check if Google Sheets logging is available."""
    _get_sheet()
    return _SHEET_CONFIGURED is True


# ─── Session State Log ────────────────────────────────────────────────

def _init_session_log():
    """Ensure session state has a cost log list."""
    if "cost_log" not in st.session_state:
        st.session_state["cost_log"] = []


def log_cost(tool, filename, num_images, input_tokens, output_tokens,
             cost_usd, user="Sean"):
    """Log a cost entry to both session state and Google Sheets.

    Args:
        tool: Which tool generated the cost (e.g. "Image Audit", "Combined Pipeline")
        filename: Name of the file being processed
        num_images: Number of images classified
        input_tokens: Total input tokens used
        output_tokens: Total output tokens used
        cost_usd: Total cost in USD
        user: Who ran it (for multi-user tracking)
    """
    _init_session_log()

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = {
        "timestamp": timestamp,
        "tool": tool,
        "filename": filename,
        "images": num_images,
        "input_tokens": input_tokens,
        "output_tokens": output_tokens,
        "cost_usd": round(cost_usd, 4),
        "user": user,
    }

    # Add to session state
    st.session_state["cost_log"].append(entry)

    # Persist to Google Sheets if configured
    sheet = _get_sheet()
    if sheet:
        try:
            sheet.append_row([
                timestamp,
                tool,
                filename,
                num_images,
                input_tokens,
                output_tokens,
                round(cost_usd, 4),
                user,
            ])
        except Exception as e:
            # Don't break the app if sheets logging fails
            st.session_state.setdefault("cost_log_errors", []).append(str(e))


def get_session_log():
    """Get the current session's cost log."""
    _init_session_log()
    return st.session_state["cost_log"]


def get_session_total():
    """Get the total cost for the current session."""
    log = get_session_log()
    return sum(entry["cost_usd"] for entry in log)


# ─── Sidebar Display ─────────────────────────────────────────────────

def render_sidebar_admin():
    """Render the admin/cost section in the sidebar.

    Call this inside `with st.sidebar:` in app.py.
    """
    _init_session_log()

    with st.expander("Admin", expanded=False):
        log = get_session_log()
        session_total = get_session_total()

        if is_sheets_configured():
            st.caption("📊 Logging to Google Sheets")
        else:
            st.caption("⚠️ Session-only (Google Sheets not configured)")

        if not log:
            st.markdown("*No API calls this session*")
        else:
            st.markdown(f"**Session total: USD ${session_total:.4f}**")
            st.markdown(f"*{len(log)} API call{'s' if len(log) != 1 else ''} this session*")

            st.markdown("---")
            for i, entry in enumerate(reversed(log), 1):
                st.markdown(
                    f"**{entry['tool']}** — `{entry['filename']}`  \n"
                    f"{entry['images']} images · "
                    f"{entry['input_tokens']:,}+{entry['output_tokens']:,} tokens · "
                    f"**${entry['cost_usd']:.4f}**  \n"
                    f"<small style='color: #999;'>{entry['timestamp']}</small>",
                    unsafe_allow_html=True,
                )

        # Show any logging errors
        errors = st.session_state.get("cost_log_errors", [])
        if errors:
            with st.expander("Logging errors"):
                for e in errors:
                    st.caption(e)
