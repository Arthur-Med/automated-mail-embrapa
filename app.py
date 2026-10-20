import os
import smtplib
import re
import imaplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
import random
from io import BytesIO

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from openpyxl.styles import Font, PatternFill

# Automatic dark-mode CSS injection for the Streamlit Quill rich-text component
try:
    import streamlit_quill
    css_path = os.path.join(os.path.dirname(streamlit_quill.__file__), "frontend", "build", "style.css")
    css_content = """
.ql-editor {
    min-height: 18em;
}

body, html {
    background-color: transparent !important;
    color: #fafafa !important;
}
.ql-toolbar {
    background-color: #1e293b !important;
    border: 1px solid rgba(250, 250, 250, 0.2) !important;
    border-top-left-radius: 8px !important;
    border-top-right-radius: 8px !important;
}
.ql-toolbar .ql-stroke {
    stroke: #fafafa !important;
}
.ql-toolbar .ql-fill {
    fill: #fafafa !important;
}
.ql-toolbar .ql-picker {
    color: #fafafa !important;
}
.ql-container {
    border: 1px solid rgba(250, 250, 250, 0.2) !important;
    border-top: none !important;
    background-color: #1e293b !important;
    border-bottom-left-radius: 8px !important;
    border-bottom-right-radius: 8px !important;
}
.ql-editor {
    background-color: transparent !important;
    color: #fafafa !important;
}
"""
    if os.path.exists(css_path):
        with open(css_path, "r", encoding="utf-8") as f:
            current_css = f.read()
        if current_css.strip() != css_content.strip():
            with open(css_path, "w", encoding="utf-8") as f:
                f.write(css_content)
except Exception:
    pass

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))


def convert_github_raw_url(url: str) -> str:
    """Converts standard GitHub URLs (including /blob/ and /raw/) into direct raw.githubusercontent.com links."""
    if not url:
        return ""
    url = url.strip().strip('"\'')
    if "github.com/" in url:
        if "/blob/" in url:
            url = url.replace("github.com/", "raw.githubusercontent.com/").replace("/blob/", "/")
        elif "/raw/" in url:
            url = url.replace("github.com/", "raw.githubusercontent.com/").replace("/raw/", "/")
    return url


def get_blocklist_filename() -> str:
    return "blocked_emails.txt"


def get_sent_filename() -> str:
    return "sent_emails.txt"


def remove_sent_email(email_addr: str) -> None:
    """Removes an email address from sent_emails.txt if a bounce or delivery failure is detected later."""
    if not email_addr or "@" not in email_addr:
        return
    filename = get_sent_filename()
    if not os.path.exists(filename):
        return
    try:
        with open(filename, "r", encoding="utf-8") as f:
            emails = set(line.strip().lower() for line in f if line.strip())
        target = email_addr.lower().strip()
        if target in emails:
            emails.remove(target)
            with open(filename, "w", encoding="utf-8") as f:
                for em in sorted(emails):
                    f.write(f"{em}\n")
    except Exception:
        pass


def purge_bounces_from_sent() -> int:
    """Ensures no address listed in blocked_emails.txt remains in sent_emails.txt."""
    blocklist_file = get_blocklist_filename()
    sent_file = get_sent_filename()
    if not os.path.exists(blocklist_file) or not os.path.exists(sent_file):
        return 0
    try:
        with open(blocklist_file, "r", encoding="utf-8") as f:
            blocked = set(line.strip().lower().replace("*", "") for line in f if line.strip())
        with open(sent_file, "r", encoding="utf-8") as f:
            sent = set(line.strip().lower() for line in f if line.strip())

        invalid = sent.intersection(blocked)
        if invalid:
            clean_sent = sent - invalid
            with open(sent_file, "w", encoding="utf-8") as f:
                for em in sorted(clean_sent):
                    f.write(f"{em}\n")
            return len(invalid)
    except Exception:
        pass
    return 0


def register_sent_email(email_addr: str) -> None:
    """Logs a recipient into sent_emails.txt only if it is not currently in the blocklist."""
    if not email_addr or "@" not in email_addr:
        return
    email_clean = email_addr.lower().strip()

    blocklist_file = get_blocklist_filename()
    if os.path.exists(blocklist_file):
        try:
            with open(blocklist_file, "r", encoding="utf-8") as f:
                blocked = set(line.strip().lower().replace("*", "") for line in f if line.strip())
            if email_clean in blocked:
                remove_sent_email(email_clean)
                return
        except Exception:
            pass

    filename = get_sent_filename()
    emails = set()
    if os.path.exists(filename):
        try:
            with open(filename, "r", encoding="utf-8") as f:
                emails = set(line.strip().lower() for line in f if line.strip())
        except Exception:
            pass
    if email_clean not in emails:
        emails.add(email_clean)
        with open(filename, "w", encoding="utf-8") as f:
            for em in sorted(emails):
                f.write(f"{em}\n")


def register_blocked_email(email_addr: str) -> None:
    """Adds a failed recipient to blocked_emails.txt and immediately purges it from sent_emails.txt."""
    if not email_addr or "@" not in email_addr:
        return
    email_clean = email_addr.lower().strip()
    remove_sent_email(email_clean)

    filename = get_blocklist_filename()
    emails = set()
    if os.path.exists(filename):
        try:
            with open(filename, "r", encoding="utf-8") as f:
                emails = set(line.strip().lower() for line in f if line.strip())
        except Exception:
            pass
    if email_clean not in emails:
        emails.add(email_clean)
        with open(filename, "w", encoding="utf-8") as f:
            for em in sorted(emails):
                f.write(f"{em}\n")


def extract_bounced_emails(limit: int = 300, update_progress: bool = True) -> int:
    """Scans the IMAP mailbox (INBOX and Spam) for Delivery Status Notifications (bounces),
    adds failed addresses to the blocklist, purges them from sent_emails.txt, and updates the progress spreadsheet."""
    email_addr = os.getenv("GMAIL_USER")
    password = os.getenv("GMAIL_APP_PASSWORD")
    if not email_addr or not password:
        return 0

    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(email_addr, password)
    except Exception:
        return 0

    bounced_emails = set()
    blocklist_file = get_blocklist_filename()
    if os.path.exists(blocklist_file):
        try:
            with open(blocklist_file, "r", encoding="utf-8") as f:
                for line in f:
                    if line.strip():
                        bounced_emails.add(line.strip().lower().replace("*", ""))
        except Exception:
            pass

    folders = ["INBOX", "[Gmail]/Spam"]
    new_bounces = 0
    found_in_run = set()

    for folder in folders:
        try:
            status, _ = mail.select(f'"{folder}"', readonly=True)
            if status != "OK":
                continue

            queries = [
                '(FROM "mailer-daemon")',
                '(FROM "postmaster")',
                '(SUBJECT "Delivery Status")',
                '(SUBJECT "Undeliverable")',
                '(SUBJECT "Falha na entrega")',
                '(SUBJECT "Failure")',
            ]

            all_ids = set()
            for q in queries:
                try:
                    _, ids = mail.search(None, q)
                    if ids[0]:
                        for x in ids[0].split():
                            all_ids.add(x)
                except Exception:
                    pass

            numeric_ids = sorted(list(set(int(x) for x in all_ids if x.isdigit())))
            if not numeric_ids:
                continue

            ids_to_check = [str(x) for x in numeric_ids[-limit:]]

            for num in ids_to_check:
                try:
                    _, data = mail.fetch(num, "(RFC822)")
                    if not data or not data[0] or not isinstance(data[0], tuple):
                        continue
                    raw_email = data[0][1]
                    msg = email.message_from_bytes(raw_email)
                    invalid_email = None

                    if msg.is_multipart():
                        for part in msg.walk():
                            ctype = part.get_content_type()
                            if ctype == "message/delivery-status":
                                payload = part.get_payload()
                                if not isinstance(payload, list):
                                    payload = [payload]
                                for p in payload:
                                    text_p = p.as_string() if hasattr(p, "as_string") else str(p)
                                    m = re.search(
                                        r"(?:Final-Recipient|Original-Recipient):\s*rfc822;\s*([^\s;<>]+)",
                                        text_p,
                                        re.IGNORECASE,
                                    )
                                    if m:
                                        invalid_email = m.group(1).strip()
                                        break
                            elif ctype == "message/rfc822" and not invalid_email:
                                try:
                                    payload = part.get_payload()
                                    if isinstance(payload, list) and payload:
                                        orig_msg = payload[0]
                                        to_hdr = orig_msg.get("To", "")
                                        m_to = re.search(r"([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)", to_hdr)
                                        if m_to:
                                            invalid_email = m_to.group(1).strip()
                                except Exception:
                                    pass
                            elif ctype == "text/plain" and not invalid_email:
                                text = part.get_payload(decode=True).decode(errors="ignore")
                                m_gmail = re.search(
                                    r"(?:não foi entregue a|wasn\'t delivered to|failed to deliver to)\s+<?([^\s<>]+@[^\s<>]+)>?",
                                    text,
                                    re.IGNORECASE,
                                )
                                m_ms = re.search(
                                    r"Delivery has failed to these recipients or groups:\s*<?([^\s<>\r\n]+@[^\s<>\r\n]+)>?",
                                    text,
                                    re.IGNORECASE,
                                )
                                m_rej = re.search(
                                    r"Recipient address rejected:.*?<([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)>",
                                    text,
                                    re.IGNORECASE,
                                )
                                if m_gmail:
                                    invalid_email = m_gmail.group(1).strip()
                                elif m_ms:
                                    invalid_email = m_ms.group(1).strip()
                                elif m_rej:
                                    invalid_email = m_rej.group(1).strip()
                    else:
                        text = msg.get_payload(decode=True).decode(errors="ignore")
                        m_gmail = re.search(
                            r"(?:não foi entregue a|wasn\'t delivered to|failed to deliver to)\s+<?([^\s<>]+@[^\s<>]+)>?",
                            text,
                            re.IGNORECASE,
                        )
                        if m_gmail:
                            invalid_email = m_gmail.group(1).strip()

                    if not invalid_email:
                        raw_str = raw_email.decode(errors="ignore")
                        m_final = re.search(
                            r"Final-Recipient:\s*rfc822;\s*<?([^\s;<>]+)>?", raw_str, re.IGNORECASE
                        )
                        if m_final:
                            invalid_email = m_final.group(1).strip()

                    if invalid_email:
                        em_clean = invalid_email.lower().strip().strip("<>;")
                        if (
                            "@" in em_clean
                            and "." in em_clean.split("@")[1]
                            and not em_clean.startswith("mailer-daemon")
                        ):
                            found_in_run.add(em_clean)
                            if em_clean not in bounced_emails:
                                bounced_emails.add(em_clean)
                                new_bounces += 1
                except Exception:
                    continue
        except Exception:
            continue

    try:
        mail.logout()
    except Exception:
        pass

    # 1. Persist all detected bounces to the blocklist file
    if new_bounces > 0 or found_in_run:
        try:
            with open(blocklist_file, "w", encoding="utf-8") as f:
                for em in sorted(bounced_emails):
                    f.write(f"{em}\n")
        except Exception:
            pass

    # 2. Purge all bounced addresses from sent_emails.txt
    purge_bounces_from_sent()
    for em in found_in_run:
        remove_sent_email(em)

    # 3. Update status in active session DataFrame and local checkpoint spreadsheet
    if update_progress and (new_bounces > 0 or found_in_run):
        try:
            if "dataframe" in st.session_state and st.session_state.dataframe is not None:
                df_sess = st.session_state.dataframe
                if "Email" in df_sess.columns and "Status" in df_sess.columns:
                    mask = df_sess["Email"].astype(str).str.strip().str.lower().isin(bounced_emails)
                    df_sess.loc[mask, "Status"] = "Error (Bounce)"
                    st.session_state.dataframe = df_sess
        except Exception:
            pass

        if os.path.exists("dispatch_progress.xlsx"):
            try:
                df_disk = pd.read_excel("dispatch_progress.xlsx")
                if "Email" in df_disk.columns and "Status" in df_disk.columns:
                    mask = df_disk["Email"].astype(str).str.strip().str.lower().isin(bounced_emails)
                    mask_affected = mask & (
                        df_disk["Status"].astype(str).str.strip().str.lower().isin(["sent", "enviado"])
                    )
                    if mask_affected.any():
                        df_disk.loc[mask_affected, "Status"] = "Error (Bounce)"
                        df_disk.to_excel("dispatch_progress.xlsx", index=False)
            except Exception:
                pass

    return len(found_in_run)


def darken_hex(hex_color: str, factor: float = 0.30) -> str:
    """Darkens a hexadecimal color by a given factor for 3D button borders."""
    try:
        h = str(hex_color).lstrip("#")
        if len(h) == 6:
            r = int(h[0:2], 16)
            g = int(h[2:4], 16)
            b = int(h[4:6], 16)
            r = max(0, int(r * (1 - factor)))
            g = max(0, int(g * (1 - factor)))
            b = max(0, int(b * (1 - factor)))
            return f"#{r:02x}{g:02x}{b:02x}"
    except Exception:
        pass
    return "#002040"


def lighten_hex(hex_color: str, factor: float = 0.15) -> str:
    """Lightens a hexadecimal color by a given factor for top highlights."""
    try:
        h = str(hex_color).lstrip("#")
        if len(h) == 6:
            r = int(h[0:2], 16)
            g = int(h[2:4], 16)
            b = int(h[4:6], 16)
            r = min(255, int(r + (255 - r) * factor))
            g = min(255, int(g + (255 - g) * factor))
            b = min(255, int(b + (255 - b) * factor))
            return f"#{r:02x}{g:02x}{b:02x}"
    except Exception:
        pass
    return "#3370a6"


def generate_button_html(
    lnk: str,
    safe_btn_text: str,
    btn_color: str,
    style: str = "Modern Elevated (Soft Shadow & Depth)",
) -> str:
    """Generates inline-styled HTML call-to-action buttons compatible with major email clients."""
    dark_color = darken_hex(btn_color, 0.30)
    slight_dark = darken_hex(btn_color, 0.15)
    light_color = lighten_hex(btn_color, 0.15)

    common_base = (
        f"background-color: {btn_color}; "
        f"color: #ffffff; "
        f"text-decoration: none; "
        f"display: inline-block; "
        f"font-weight: bold; "
        f"font-family: Arial, sans-serif; "
        f"font-size: 16px; "
        f"width: 270px; "
        f"box-sizing: border-box; "
        f"white-space: normal; "
        f"line-height: 1.4; "
        f"word-wrap: break-word; "
        f"letter-spacing: 0.3px; "
    )

    style_str = str(style)
    if "Pill" in style_str:
        style_btn = (
            common_base
            + f"background-image: linear-gradient(180deg, {light_color} 0%, {btn_color} 100%); "
            f"padding: 16px 24px; "
            f"border-radius: 50px; "
            f"border-top: 1px solid {light_color}; "
            f"border-left: 1px solid {slight_dark}; "
            f"border-right: 1px solid {slight_dark}; "
            f"border-bottom: 2px solid {dark_color}; "
            f"box-shadow: 0 8px 20px -2px rgba(0, 0, 0, 0.25), 0 3px 6px -1px rgba(0, 0, 0, 0.12), inset 0 1px 0 rgba(255, 255, 255, 0.3); "
            f"text-shadow: 0 1px 2px rgba(0, 0, 0, 0.35);"
        )
    elif "Beveled" in style_str:
        style_btn = (
            common_base
            + f"background-image: linear-gradient(180deg, {light_color} 0%, {btn_color} 100%); "
            f"padding: 15px 22px 13px 22px; "
            f"border-radius: 8px; "
            f"border-top: 1px solid {light_color}; "
            f"border-left: 1px solid {slight_dark}; "
            f"border-right: 1px solid {slight_dark}; "
            f"border-bottom: 3px solid {dark_color}; "
            f"box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2), inset 0 1px 0 rgba(255, 255, 255, 0.2); "
            f"text-shadow: 0 1px 2px rgba(0, 0, 0, 0.35);"
        )
    elif "Gradient" in style_str:
        strong_light = lighten_hex(btn_color, 0.28)
        style_btn = (
            common_base
            + f"background-image: linear-gradient(180deg, {strong_light} 0%, {btn_color} 60%, {slight_dark} 100%); "
            f"padding: 16px 22px; "
            f"border-radius: 8px; "
            f"border: 1px solid {dark_color}; "
            f"border-top: 1px solid {strong_light}; "
            f"box-shadow: 0 6px 16px rgba(0, 0, 0, 0.28), inset 0 1px 1px rgba(255, 255, 255, 0.35); "
            f"text-shadow: 0 1px 2px rgba(0, 0, 0, 0.5);"
        )
    else:
        style_btn = (
            common_base
            + f"background-image: linear-gradient(180deg, {light_color} 0%, {btn_color} 100%); "
            f"padding: 16px 22px; "
            f"border-radius: 10px; "
            f"border-top: 1px solid {light_color}; "
            f"border-left: 1px solid {slight_dark}; "
            f"border-right: 1px solid {slight_dark}; "
            f"border-bottom: 2px solid {dark_color}; "
            f"box-shadow: 0 6px 18px -2px rgba(0, 0, 0, 0.25), 0 2px 5px -1px rgba(0, 0, 0, 0.12), inset 0 1px 0 rgba(255, 255, 255, 0.25); "
            f"text-shadow: 0 1px 2px rgba(0, 0, 0, 0.3);"
        )
    return f'<div style="text-align: center; margin: 25px 0;"><a href="{lnk}" class="btn-cta-relevo" style="{style_btn}">{safe_btn_text}</a></div>'


favicon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "favicon.ico")
if not os.path.exists(favicon_path):
    favicon_path = "assets/favicon.ico"

st.set_page_config(
    page_title="Automated Mail Sender",
    page_icon=favicon_path if os.path.exists(favicon_path) else "",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
/* =================================================================
   GLOBAL STREAMLIT RESET & LAYOUT CONFIGURATION
   ================================================================= */
#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
header { visibility: hidden; }

.stApp {
    width: 100% !important;
}

[data-testid="stAppViewContainer"] {
    width: 100% !important;
}

section.main {
    width: 100% !important;
}

.block-container { 
    padding-top: 1.5rem !important; 
    padding-bottom: 5.5rem !important; 
    padding-left: 2.5rem !important;
    padding-right: 2.5rem !important;
    max-width: 100% !important;
    width: 100% !important;
}

/* =================================================================
   HIDE DEFAULT STREAMLIT HEADER ANCHOR LINKS (#)
   ================================================================= */
[data-testid="stHeaderActionElements"],
[data-testid="stHeadingWithActionElements"] a,
a.anchor-link,
a.header-anchor,
.stMarkdown a[href^="#"],
.stMarkdown a.header-anchor,
h1 a, h2 a, h3 a, h4 a, h5 a, h6 a,
.st-emotion-cache-15zrgzn,
.st-emotion-cache-16idsys a {
    display: none !important;
    visibility: hidden !important;
    pointer-events: none !important;
    opacity: 0 !important;
    width: 0 !important;
    height: 0 !important;
    margin: 0 !important;
    padding: 0 !important;
    position: absolute !important;
}

/* =================================================================
   DARK THEME VARIABLES & DESIGN SYSTEM
   ================================================================= */
:root {
    --brand-primary: #0284c7;
    --brand-gradient: linear-gradient(135deg, #1d4ed8 0%, #0284c7 100%);
    --brand-light: #38bdf8;
    --bg-main: #0b0f19;
    --card-bg: rgba(30, 41, 59, 0.7);
    --card-border: rgba(255, 255, 255, 0.08);
    --card-hover-border: rgba(56, 189, 248, 0.35);
    --text-title: #f8fafc;
    --text-body: #cbd5e1;
    --text-muted: #94a3b8;
}

body, html, .stApp {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif !important;
    color: #f1f5f9 !important;
}

/* =================================================================
   MODERN TOP NAVBAR
   ================================================================= */
.app-navbar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    background: linear-gradient(135deg, rgba(30, 41, 59, 0.85) 0%, rgba(15, 23, 42, 0.95) 100%);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    border: 1px solid rgba(255, 255, 255, 0.1);
    border-radius: 16px;
    padding: 16px 28px;
    margin-bottom: 24px;
    box-shadow: 0 10px 30px -5px rgba(0, 0, 0, 0.4), 0 0 0 1px rgba(255, 255, 255, 0.05);
}

.nav-brand {
    display: flex;
    align-items: center;
    gap: 18px;
}

.nav-logo {
    height: 48px;
    max-width: 155px;
    object-fit: contain;
    filter: brightness(0) invert(1) opacity(0.95);
    transition: transform 0.2s ease, opacity 0.2s ease;
}
.nav-logo:hover {
    transform: scale(1.02);
    opacity: 1;
}

.nav-divider {
    width: 1px;
    height: 38px;
    background: rgba(255, 255, 255, 0.15);
}

.nav-text-group {
    display: flex;
    flex-direction: column;
    gap: 3px;
}

.nav-title {
    font-size: 1.55rem;
    font-weight: 800;
    letter-spacing: -0.5px;
    background: linear-gradient(135deg, #60a5fa 0%, #38bdf8 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    line-height: 1.2;
    user-select: none;
}

.nav-badge-row {
    display: flex;
    align-items: center;
    gap: 8px;
}

.nav-sub {
    font-size: 0.82rem;
    color: #94a3b8;
    font-weight: 500;
}

.nav-status-pill {
    display: flex;
    align-items: center;
    gap: 10px;
    background: rgba(16, 185, 129, 0.1);
    border: 1px solid rgba(16, 185, 129, 0.25);
    padding: 8px 18px;
    border-radius: 9999px;
    font-size: 0.85rem;
    box-shadow: 0 2px 10px rgba(16, 185, 129, 0.1);
}

.status-pulse-dot {
    width: 8px;
    height: 8px;
    background-color: #10b981;
    border-radius: 50%;
    box-shadow: 0 0 10px #10b981;
    animation: statusPulse 2s infinite ease-in-out;
}

@keyframes statusPulse {
    0%, 100% { opacity: 1; transform: scale(1); }
    50% { opacity: 0.4; transform: scale(0.8); }
}

.status-label {
    color: #94a3b8;
    font-weight: 500;
}

.status-user {
    color: #f1f5f9;
    font-weight: 600;
}

/* =================================================================
   STEPPER TABS (STEP 1, 2, 3)
   ================================================================= */
.stTabs [data-baseweb="tab-list"] {
    display: flex !important;
    gap: 12px !important;
    background: rgba(15, 23, 42, 0.7) !important;
    backdrop-filter: blur(12px) !important;
    border: 1px solid rgba(255, 255, 255, 0.08) !important;
    border-radius: 14px !important;
    padding: 8px !important;
    margin-bottom: 24px !important;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.25) !important;
}

.stTabs [data-baseweb="tab"] {
    flex: 1 !important;
    height: 52px !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    border-radius: 10px !important;
    padding: 0 20px !important;
    background: transparent !important;
    border: 1px solid transparent !important;
    transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1) !important;
    color: #94a3b8 !important;
}

.stTabs [data-baseweb="tab"]:hover {
    background: rgba(255, 255, 255, 0.04) !important;
    color: #f1f5f9 !important;
    border-color: rgba(255, 255, 255, 0.08) !important;
}

.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #1d4ed8 0%, #0284c7 100%) !important;
    border-color: rgba(96, 165, 250, 0.4) !important;
    box-shadow: 0 4px 18px rgba(29, 78, 216, 0.4) !important;
}

.stTabs [aria-selected="true"] p, 
.stTabs [aria-selected="true"] span {
    color: #ffffff !important;
    font-weight: 700 !important;
    letter-spacing: 0.3px !important;
}

.stTabs [aria-selected="false"] p,
.stTabs [aria-selected="false"] span {
    color: #94a3b8 !important;
    font-weight: 500 !important;
}

[data-baseweb="tab-highlight"], [data-baseweb="tab-border"] {
    display: none !important;
}

/* =================================================================
   STEP HEADERS
   ================================================================= */
.step-card-header {
    background: linear-gradient(135deg, rgba(30, 41, 59, 0.5) 0%, rgba(15, 23, 42, 0.6) 100%);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 14px;
    padding: 22px 26px;
    margin-bottom: 24px;
}

.step-badge {
    display: inline-block;
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 1px;
    color: #38bdf8;
    background: rgba(56, 189, 248, 0.12);
    border: 1px solid rgba(56, 189, 248, 0.25);
    padding: 3px 10px;
    border-radius: 6px;
    margin-bottom: 8px;
    text-transform: uppercase;
}

.step-title {
    font-size: 1.45rem;
    font-weight: 700;
    color: #f8fafc;
    margin: 0 0 6px 0;
    letter-spacing: -0.3px;
}

.step-subtitle {
    font-size: 0.95rem;
    color: #94a3b8;
    line-height: 1.5;
    margin: 0;
}

.section-label {
    font-size: 0.95rem;
    font-weight: 600;
    color: #e2e8f0;
    margin-bottom: 10px;
    display: flex;
    align-items: center;
    gap: 8px;
}

/* =================================================================
   CONTAINERS, CARDS & BUTTONS
   ================================================================= */
[data-testid="stVerticalBlockBorderWrapper"] {
    background: rgba(30, 41, 59, 0.6) !important;
    border: 1px solid rgba(255, 255, 255, 0.08) !important;
    border-radius: 14px !important;
    box-shadow: 0 4px 20px -2px rgba(0, 0, 0, 0.3) !important;
    transition: border-color 0.2s ease, box-shadow 0.2s ease !important;
}

[data-testid="stVerticalBlockBorderWrapper"]:hover {
    border-color: rgba(56, 189, 248, 0.3) !important;
}

div.stButton > button[kind="primary"], div.stDownloadButton > button[kind="primary"] {
    background: linear-gradient(135deg, #1d4ed8 0%, #0284c7 100%) !important;
    color: white !important;
    border: 1px solid rgba(96, 165, 250, 0.3) !important;
    border-radius: 10px !important;
    padding: 0.65rem 1.6rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.3px !important;
    transition: all 0.25s ease !important;
    box-shadow: 0 4px 14px rgba(29, 78, 216, 0.35) !important;
}

div.stButton > button[kind="primary"]:hover, div.stDownloadButton > button[kind="primary"]:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 6px 20px rgba(29, 78, 216, 0.5) !important;
    border-color: rgba(96, 165, 250, 0.6) !important;
}

div.stButton > button:not([kind="primary"]) {
    background: rgba(30, 41, 59, 0.6) !important;
    border: 1px solid rgba(255, 255, 255, 0.1) !important;
    color: #f1f5f9 !important;
    border-radius: 10px !important;
    transition: all 0.2s ease !important;
}

div.stButton > button:not([kind="primary"]):hover {
    background: rgba(51, 65, 85, 0.7) !important;
    border-color: rgba(255, 255, 255, 0.2) !important;
}

/* Metrics */
div[data-testid="stMetric"] {
    background: rgba(15, 23, 42, 0.6) !important;
    border: 1px solid rgba(255, 255, 255, 0.08) !important;
    border-radius: 12px !important;
    padding: 16px 20px !important;
    transition: transform 0.2s ease, border-color 0.2s ease !important;
}

div[data-testid="stMetric"]:hover {
    border-color: rgba(56, 189, 248, 0.3) !important;
    transform: translateY(-2px) !important;
}

[data-testid="stMetricValue"] {
    color: #38bdf8 !important;
    font-weight: 800 !important;
    font-size: 2.1rem !important;
}

[data-testid="stMetricLabel"] {
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    color: #94a3b8 !important;
}

/* Inputs & File Uploader */
.stTextInput>div>div>input, .stTextArea>div>div>textarea {
    border-radius: 8px !important;
    padding: 12px !important;
    background-color: rgba(15, 23, 42, 0.7) !important;
    border: 1px solid rgba(255, 255, 255, 0.12) !important;
    color: #f8fafc !important;
    transition: all 0.2s ease !important;
}

.stTextInput>div>div>input:focus, .stTextArea>div>div>textarea:focus {
    border-color: #38bdf8 !important;
    box-shadow: 0 0 0 2px rgba(56, 189, 248, 0.2) !important;
}

[data-testid="stFileUploader"] {
    background: rgba(15, 23, 42, 0.5) !important;
    border-radius: 12px !important;
    padding: 10px !important;
    border: 1px dashed rgba(255, 255, 255, 0.15) !important;
}
</style>
""",
    unsafe_allow_html=True,
)

if "dispatch_completed" not in st.session_state:
    st.session_state.dispatch_completed = False
if "processing" not in st.session_state:
    st.session_state.processing = False
if "final_report" not in st.session_state:
    st.session_state.final_report = None
if "dataframe" not in st.session_state:
    st.session_state.dataframe = None
if "external_image_url" not in st.session_state:
    st.session_state.external_image_url = None
if "header_logo_url" not in st.session_state:
    st.session_state.header_logo_url = ""

email_user = os.getenv("GMAIL_USER") or "sender@example.org"
email_pass = os.getenv("GMAIL_APP_PASSWORD") or ""

import base64


def load_ui_logo() -> str:
    """Loads the local sample brand logo as Base64 for instant rendering in the UI and preview."""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "sample_logo.png")
    if os.path.exists(path):
        try:
            with open(path, "rb") as f_logo:
                return f"data:image/png;base64,{base64.b64encode(f_logo.read()).decode('utf-8')}"
        except Exception:
            pass
    return ""


ui_sample_logo = load_ui_logo()

# Top Navigation Bar
nav_logo_el = (
    f'<img src="{ui_sample_logo}" alt="Logo" class="nav-logo"><div class="nav-divider"></div>'
    if ui_sample_logo
    else ""
)
st.markdown(
    f"""<div class="app-navbar">
<div class="nav-brand">
{nav_logo_el}<div class="nav-text-group">
<div class="nav-title">Automated Mail Sender</div>
<div class="nav-badge-row">
<span class="nav-sub">Enterprise Bulk Email &amp; Campaign Automation Engine</span>
</div>
</div>
</div>
<div class="nav-right">
<div class="nav-status-pill">
<span class="status-pulse-dot"></span>
<span class="status-label">Connected:</span>
<span class="status-user">{email_user}</span>
</div>
</div>
</div>""",
    unsafe_allow_html=True,
)

st.write("")

tab1, tab2, tab3 = st.tabs(
    [
        "Step 1: Contact Database",
        "Step 2: Message & Layout",
        "Step 3: Dispatch & Analytics",
    ]
)

with tab1:
    st.write("")
    st.markdown(
        """
    <div class="step-card-header">
        <span class="step-badge">STEP 01</span>
        <div class="step-title">Contact Spreadsheet Import</div>
        <div class="step-subtitle">Upload an Excel file (<code>.xlsx</code>) containing your recipient list. The spreadsheet must include <b>Name</b> and <b>Email</b> columns (optional: <b>Course</b>, <b>Link</b>).</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader("", type=["xlsx"], key="uploader")

    if uploaded_file is not None:
        file_identifier = (uploaded_file.name, uploaded_file.size)
        if st.session_state.dataframe is None or st.session_state.get("last_file_id") != file_identifier:
            df_raw = pd.read_excel(uploaded_file)

            # Normalize Portuguese column names to English if uploaded from legacy sheets
            col_rename = {}
            for col in df_raw.columns:
                col_clean = str(col).strip().lower()
                if col_clean == "nome" and "Name" not in df_raw.columns:
                    col_rename[col] = "Name"
                elif col_clean == "curso" and "Course" not in df_raw.columns:
                    col_rename[col] = "Course"
            if col_rename:
                df_raw = df_raw.rename(columns=col_rename)

            if not st.session_state.get("bounced_synced_once", False):
                with st.spinner("Synchronizing delivery status and anti-spam blocklist with Gmail..."):
                    try:
                        purge_bounces_from_sent()
                        extract_bounced_emails(limit=300, update_progress=True)
                    except Exception:
                        pass
                st.session_state.bounced_synced_once = True

            if "Email" in df_raw.columns:
                df_raw = df_raw.dropna(subset=["Email"])
                df_raw = df_raw[df_raw["Email"].astype(str).str.contains("@")]
                df_raw["Email"] = df_raw["Email"].astype(str).str.strip()

                # Blocklist integration
                blocklist_file = get_blocklist_filename()
                if os.path.exists(blocklist_file):
                    try:
                        with open(blocklist_file, "r", encoding="utf-8") as f:
                            blocked_emails = set(line.strip().lower() for line in f if line.strip())

                        if blocked_emails:
                            size_before = len(df_raw)
                            df_raw = df_raw[~df_raw["Email"].str.lower().isin(blocked_emails)]
                            removed = size_before - len(df_raw)

                            if removed > 0:
                                st.warning(
                                    f"{removed} recipient(s) were automatically filtered out because they appear in `{blocklist_file}`."
                                )
                    except Exception as e:
                        st.error(f"Error reading {blocklist_file}: {e}")

            if "Status" not in df_raw.columns:
                df_raw["Status"] = ""
            df_raw["Status"] = df_raw["Status"].astype(object)

            st.session_state.dataframe = df_raw
            st.session_state.last_file_id = file_identifier
            st.session_state.dispatch_completed = False
            st.session_state.processing = False
            st.session_state.final_report = None

        df = st.session_state.dataframe

        if "Name" not in df.columns or "Email" not in df.columns:
            st.error("The uploaded spreadsheet must contain 'Name' and 'Email' columns.")
        else:
            total = len(df)
            status_clean = df["Status"].astype(str).str.strip().str.lower()
            sent_count = len(df[status_clean.isin(["sent", "enviado"])])
            bounces_count = len(
                df[status_clean.isin(["error", "error (bounce)", "failed", "erro", "erro (bounce)", "falha"])]
            )
            pending = total - sent_count - bounces_count

            st.write("")
            with st.container(border=True):
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Total Contacts", total)
                c2.metric("Pending in Queue", pending)
                c3.metric("Successfully Delivered", sent_count)
                c4.metric("Failures / Bounces", bounces_count)

            col_sync1, _ = st.columns([2, 3])
            with col_sync1:
                if st.button("Sync Gmail Bounces Now", use_container_width=True, key="btn_sync_tab1"):
                    with st.spinner("Scanning Gmail via IMAP and purging invalid addresses..."):
                        extract_bounced_emails(limit=300, update_progress=True)
                        st.success("Synchronization complete! Bounces updated and sent list sanitized.")
                        time.sleep(1)
                        st.rerun()

            st.write("")
            st.markdown('<div class="section-label">Loaded Dataset Preview</div>', unsafe_allow_html=True)
            st.dataframe(df, use_container_width=True, height=250)
            st.info("Spreadsheet successfully loaded and validated.")

with tab2:
    st.write("")
    st.markdown(
        """
    <div class="step-card-header">
        <span class="step-badge">STEP 02</span>
        <div class="step-title">Message Composition &amp; HTML Template</div>
        <div class="step-subtitle">Configure the email subject, compose the body using dynamic merge tags such as <code>{Name}</code>, <code>{Course}</code>, and <code>{Link}</code>, and customize branding, buttons, and media.</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    if st.session_state.dataframe is None:
        st.warning("Please upload a contact spreadsheet in Step 1 first.")
    else:
        df = st.session_state.dataframe
        col_txt, col_prev = st.columns([1.2, 1.8], gap="large")

        with col_txt:
            with st.container(border=True):
                st.markdown('<div class="section-label">Email Subject &amp; Template Branding</div>', unsafe_allow_html=True)
                default_subject = "Open Enrollment: {Course}"
                user_subject = st.text_input(
                    "Email Subject:",
                    value=default_subject,
                    key="subject_msg",
                )
                col_brand1, col_brand2 = st.columns(2)
                with col_brand1:
                    header_logo_raw = st.text_input(
                        "Header Logo URL (Optional):",
                        value=st.session_state.get("header_logo_url", ""),
                        placeholder="https://raw.githubusercontent.com/.../logo.png",
                        help="Optional public image URL displayed at the top of the email template.",
                        key="input_header_logo",
                    )
                    st.session_state.header_logo_url = convert_github_raw_url(header_logo_raw)
                with col_brand2:
                    org_footer_name = st.text_input(
                        "Organization Footer Text:",
                        value="Institutional Training & Outreach Division",
                        key="org_footer_name",
                    )

            st.markdown(
                """
            <style>
                iframe[title="streamlit_quill.st_quill"] {
                    min-height: 300px !important;
                    border: none !important;
                    background-color: transparent !important;
                }
            </style>
            """,
                unsafe_allow_html=True,
            )

            default_body = (
                "<p>Dear {Name},</p>"
                "<p>We are pleased to announce that registration is now open for the <strong>{Course}</strong> training program!</p>"
                "<p>Click the button below to view the full syllabus and secure your spot:</p>"
                "<p>{Link}</p>"
                "<p>Best regards,<br>Training &amp; Outreach Team</p>"
            )
            from streamlit_quill import st_quill

            custom_toolbar = [
                ["bold", "italic", "underline"],
                [{"size": ["small", False, "large", "huge"]}],
                [{"align": []}],
            ]

            with st.container(border=True):
                st.markdown('<div class="section-label">Main Email Body</div>', unsafe_allow_html=True)
                user_text = st_quill(value=default_body, html=True, toolbar=custom_toolbar, key="text_msg_top")
                if not user_text:
                    user_text = default_body

            with st.container(border=True):
                st.markdown(
                    '<div class="section-label">Embedded Banner / Flyer (Optional - GitHub / Raw URL)</div>',
                    unsafe_allow_html=True,
                )

                image_url_input = st.text_input(
                    "Direct Image URL (GitHub Raw / Public URL):",
                    value=st.session_state.get("external_image_url", "") or "",
                    placeholder="https://raw.githubusercontent.com/username/repository/main/banner.png",
                    help="Paste a GitHub image link or public URL. Standard GitHub '/blob/' links are automatically converted to direct Raw URLs.",
                    key="input_link_image_raw",
                )

                clean_image_url = convert_github_raw_url(image_url_input)
                bottom_text = ""

                if clean_image_url:
                    st.session_state.external_image_url = clean_image_url
                    image_url = clean_image_url

                    try:
                        st.image(image_url, caption="Hosted Banner Preview", width=300)
                    except Exception:
                        st.warning(
                            "Could not render a local preview for this URL. Ensure the repository and file are public."
                        )

                    st.session_state.img_size_pct = st.slider(
                        "Image width inside email (%)",
                        min_value=10,
                        max_value=100,
                        value=st.session_state.get("img_size_pct", 100),
                        step=5,
                    )

                    st.markdown('<div class="section-label">Text Below Image (Optional)</div>', unsafe_allow_html=True)
                    bottom_text = st_quill(value="", html=True, toolbar=custom_toolbar, key="text_msg_bottom")
                    if not bottom_text:
                        bottom_text = ""
                else:
                    st.session_state.external_image_url = None
                    image_url = None

            full_text = user_text
            if st.session_state.get("external_image_url"):
                full_text += "\n{Image}\n" + bottom_text

            st.session_state.text_msg = full_text

            with st.container(border=True):
                st.markdown('<div class="section-label">Call-to-Action Button Customization</div>', unsafe_allow_html=True)
                found_links = re.findall(r"(https?://[^\s<>\"'\,]+|\{Link\})", full_text)
                unique_links = list(dict.fromkeys(found_links))

                style_options = [
                    "Modern Elevated (Soft Shadow & Depth)",
                    "Floating Pill (Premium Rounded)",
                    "Soft Beveled 3D (Classic Embossed)",
                    "High-Contrast 3D Gradient (Vibrant)",
                ]

                if unique_links:
                    for idx, lnk in enumerate(unique_links):
                        col_b1, col_b2 = st.columns([3, 1])
                        with col_b1:
                            st.text_input(
                                f"Button {idx+1} Label",
                                value="Click here to access",
                                key=f"btn_txt_{idx}_{lnk}",
                                help=f"Target URL: {lnk}",
                            )
                        with col_b2:
                            st.color_picker(
                                f"Color {idx+1}",
                                value="#004B87",
                                key=f"btn_color_{idx}_{lnk}",
                                label_visibility="hidden",
                            )

                        st.selectbox(
                            f"Button {idx+1} Style",
                            options=style_options,
                            index=0,
                            key=f"btn_style_{idx}_{lnk}",
                            help="Choose the visual depth and surface finish of the CTA button.",
                        )
                else:
                    st.info("No URLs or `{Link}` tags detected in the message body. Add a link to generate a CTA button.")

        with col_prev:
            st.markdown('<div class="section-label">Real-Time HTML Email Preview</div>', unsafe_allow_html=True)
            try:
                sample = df.iloc[0]

                def clean_str(val, default):
                    if pd.isna(val) or str(val).strip().lower() in ["", "nan"]:
                        return default
                    return str(val).strip()

                sample_name = clean_str(sample.get("Name"), "Recipient Name")
                sample_course = clean_str(sample.get("Course"), "[COURSE NAME]")
                sample_link = clean_str(sample.get("Link"), "https://example.org")

                preview_subject = user_subject.format(
                    Name=sample_name, Course=sample_course, Link=sample_link,
                    Nome=sample_name, Curso=sample_course
                )

                modified_text = full_text
                for idx, lnk in enumerate(unique_links):
                    btn_text = st.session_state.get(f"btn_txt_{idx}_{lnk}", "Click here to access")
                    btn_color = st.session_state.get(f"btn_color_{idx}_{lnk}", "#004B87")
                    btn_style = st.session_state.get(
                        f"btn_style_{idx}_{lnk}", "Modern Elevated (Soft Shadow & Depth)"
                    )
                    safe_btn_text = btn_text.replace("{", "{{").replace("}", "}}")
                    button_html = generate_button_html(lnk, safe_btn_text, btn_color, btn_style)
                    modified_text = modified_text.replace(lnk, button_html)

                preview_text = modified_text.format(
                    Name=sample_name,
                    Course=sample_course,
                    Link=sample_link,
                    Image="{Image}",
                    Nome=sample_name,
                    Curso=sample_course,
                    Imagem="{Image}",
                )
                preview_html = preview_text

                if st.session_state.get("external_image_url"):
                    img_url = st.session_state.external_image_url
                    size_pct = st.session_state.get("img_size_pct", 100)
                    img_tag = f'<div style="text-align: center; margin: 15px 0;"><img src="{img_url}" style="max-width: {size_pct}%; height: auto; display: block; margin: 0 auto; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);"></div>'
                    if "{Image}" in preview_html:
                        preview_html = preview_html.replace("{Image}", img_tag)
                    else:
                        preview_html += img_tag
                else:
                    preview_html = preview_html.replace("{Image}", "")

                preview_logo_src = st.session_state.get("header_logo_url") or ui_sample_logo
                if preview_logo_src:
                    header_logo_html = f'<div style="text-align: center; padding: 40px 20px 30px 20px; border-top: 6px solid #004B87; border-bottom: 1px solid #f1f3f5;"><img src="{preview_logo_src}" alt="Logo" style="height: 85px; border: 0; display: block; margin: 0 auto; max-height: 85px;"></div>'
                else:
                    header_logo_html = '<div style="border-top: 6px solid #004B87;"></div>'

                footer_label = st.session_state.get("org_footer_name", "Institutional Training & Outreach Division")

                st.markdown(
                    f"""
<div style="margin-bottom: 20px; padding: 20px; background-color: #1e293b; border: 1px solid #334155; border-radius: 12px; box-shadow: inset 0 2px 4px rgba(0,0,0,0.2);">
<p style="margin: 4px 0; font-size: 14px; color: #94a3b8;"><b>From:</b> <span style="color: #f8fafc;">{email_user}</span></p>
<p style="margin: 4px 0; font-size: 14px; color: #94a3b8;"><b>To:</b> <span style="color: #f8fafc;">{sample.get('Email', 'recipient@example.com')}</span></p>
<p style="margin: 4px 0; font-size: 14px; color: #3b82f6; font-weight: 600;">{preview_subject}</p>
</div>

<div style="background-color: #e9ecef; padding: 40px 0; border-radius: 8px;">
<div style="max-width: 800px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; border: 1px solid #dee2e6; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.05);">
{header_logo_html}
<style>
    .email-content p {{ margin: 0 0 5px 0 !important; }}
    .btn-cta-relevo {{
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
        cursor: pointer !important;
    }}
    .btn-cta-relevo:hover {{
        filter: brightness(1.08) !important;
        box-shadow: 0 10px 25px -2px rgba(0, 0, 0, 0.35), 0 4px 8px -1px rgba(0, 0, 0, 0.15) !important;
        transform: translateY(-2px) !important;
    }}
    .btn-cta-relevo:active {{
        transform: translateY(1px) !important;
        box-shadow: 0 4px 10px -1px rgba(0, 0, 0, 0.2) !important;
    }}
</style>
<div class="email-content" style="padding: 40px 50px; color: #343a40; font-size: 16px; line-height: 1.5; font-family: 'Segoe UI', Arial, sans-serif;">
{preview_html}
</div>
<div style="text-align: center; background-color: #f8f9fa; padding: 30px 40px; border-top: 1px solid #e9ecef;">
<p style="margin: 0; font-size: 13px; color: #6c757d; line-height: 1.6; font-family: 'Segoe UI', Arial, sans-serif;">
<strong>{footer_label}</strong><br>
Automated Outreach &amp; Communications<br>
</p>
</div>
</div>
</div>
                    """,
                    unsafe_allow_html=True,
                )
            except Exception:
                st.error("Preview rendering error. Please verify that your dynamic merge tags are properly formatted.")

with tab3:
    st.write("")
    st.markdown(
        """
    <div class="step-card-header">
        <span class="step-badge">STEP 03</span>
        <div class="step-title">Batch Dispatch &amp; Delivery Reports</div>
        <div class="step-subtitle">Monitor real-time SMTP transmission with human-like anti-spam throttling and export consolidated Excel delivery reports.</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    if st.session_state.dataframe is None:
        st.info("Upload a contact spreadsheet in Step 1 to enable batch dispatching.")
    elif not st.session_state.get("text_msg", "").strip():
        st.info("Configure your email subject and body in Step 2 before starting.")
    else:
        df = st.session_state.dataframe
        total_all = len(df)
        status_clean = df["Status"].astype(str).str.strip().str.lower()
        sent_count = len(df[status_clean.isin(["sent", "enviado"])])
        errors_count = len(
            df[status_clean.isin(["error", "error (bounce)", "failed", "erro", "erro (bounce)", "falha"])]
        )
        pending_count = total_all - sent_count - errors_count
        df_filtered = df[~status_clean.isin(["sent", "enviado"])]

        is_running = st.session_state.get("processing", False)

        # Top Metrics Panel
        st.write("")
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            st.metric("Total Contacts", total_all)
        with col_m2:
            st.metric("Successfully Delivered", sent_count)
        with col_m3:
            st.metric("Pending in Queue", pending_count)
        with col_m4:
            rate = int((sent_count / total_all) * 100) if total_all > 0 else 0
            st.metric("Queue Completion", f"{rate}%")

        st.write("")

        # Main Action Card
        with st.container(border=True):
            if df_filtered.empty:
                st.markdown(
                    """
                <div style="padding: 14px 18px; background: rgba(34, 197, 94, 0.1); border: 1px solid rgba(34, 197, 94, 0.25); border-radius: 8px; color: #4ade80; font-size: 14px; font-weight: 500; margin-bottom: 12px;">
                    All recipients in this spreadsheet have been processed and marked as delivered.
                </div>
                """,
                    unsafe_allow_html=True,
                )

                col_end1, _ = st.columns([1.5, 3.5])
                with col_end1:
                    if st.button("Reset Spreadsheet Queue", use_container_width=True, type="secondary"):
                        df["Status"] = ""
                        st.session_state.dataframe = df
                        st.session_state.dispatch_completed = False
                        st.session_state.processing = False
                        st.session_state.final_report = None
                        if os.path.exists("dispatch_progress.xlsx"):
                            try:
                                os.remove("dispatch_progress.xlsx")
                            except Exception:
                                pass
                        st.rerun()
            else:
                total_pending = len(df_filtered)
                total_already_sent = total_all - total_pending

                if total_already_sent > 0:
                    btn_label = f"RESUME DISPATCH ({total_pending} pending)"
                    status_text = (
                        f"Current status: <b>{total_already_sent}</b> emails sent | <b>{total_pending}</b> remaining in queue."
                    )
                else:
                    btn_label = f"START DISPATCH ({total_pending} contacts)"
                    status_text = f"Queue ready: <b>{total_pending}</b> recipients prepared for transmission."

                st.markdown(
                    f"""
                <div style="padding: 10px 14px; background: rgba(59, 130, 246, 0.1); border: 1px solid rgba(59, 130, 246, 0.25); border-radius: 8px; color: #93c5fd; font-size: 14px; margin-bottom: 14px;">
                    {status_text}
                </div>
                """,
                    unsafe_allow_html=True,
                )

                if is_running:
                    col_act1, col_act2 = st.columns([2.5, 1.2])
                    with col_act1:
                        st.button(
                            "Dispatch in Progress...",
                            type="primary",
                            use_container_width=True,
                            disabled=True,
                        )
                    with col_act2:
                        if st.button("Pause Dispatch", type="secondary", use_container_width=True, key="btn_pause"):
                            st.session_state.processing = False
                            st.rerun()
                else:
                    col_act1, col_act2 = st.columns([2.5, 1.2])
                    with col_act1:
                        if st.button(btn_label, type="primary", use_container_width=True):
                            st.session_state.processing = True
                            st.session_state.dispatch_completed = False
                            st.rerun()
                    with col_act2:
                        if st.button(
                            "Reset Queue",
                            type="secondary",
                            use_container_width=True,
                            help="Clears current delivery progress and restarts the queue from the beginning.",
                        ):
                            df["Status"] = ""
                            st.session_state.dataframe = df
                            st.session_state.dispatch_completed = False
                            st.session_state.processing = False
                            st.session_state.final_report = None
                            if os.path.exists("dispatch_progress.xlsx"):
                                try:
                                    os.remove("dispatch_progress.xlsx")
                                except Exception:
                                    pass
                            st.rerun()

        if is_running and not df_filtered.empty:
            status_box = st.status("Initializing connection to SMTP server...", expanded=True)
            progress_bar = st.progress(0)
            st.markdown(
                """
    <div class="step-card-header" style="margin-top: 25px; margin-bottom: 15px;">
        <div class="step-title" style="font-size: 1.15rem;">Live Transmission &amp; Throttling Logs</div>
        <div class="step-subtitle">Real-time SMTP server responses and anti-spam cooldown intervals.</div>
    </div>
    """,
                unsafe_allow_html=True,
            )
            log_container = st.container(height=350)

            try:
                def connect_smtp():
                    """Establishes an SMTP connection with optional corporate HTTP proxy tunneling (PySocks)
                    and automatic multi-port fallback (587 STARTTLS -> 25 STARTTLS -> 465 SSL)."""
                    import socket
                    import socks

                    original_socket = socket.socket
                    proxy_host = os.getenv("PROXY_HOST", "").strip()
                    proxy_port = int(os.getenv("PROXY_PORT", "3128"))
                    proxy_user = os.getenv("PROXY_USER", "").strip() or None
                    proxy_pass = os.getenv("PROXY_PASS", "").strip() or None

                    try:
                        if proxy_host:
                            socks.setdefaultproxy(
                                socks.PROXY_TYPE_HTTP,
                                proxy_host,
                                proxy_port,
                                True,
                                proxy_user,
                                proxy_pass,
                            )
                            socket.socket = socks.socksocket

                        try:
                            status_box.update(label="Attempting connection via Port 587 (STARTTLS)...", state="running")
                            s = smtplib.SMTP("smtp.gmail.com", 587, timeout=60)
                            s.starttls()
                            s.login(email_user, email_pass)
                            return s
                        except Exception:
                            try:
                                status_box.update(
                                    label="Port 587 unavailable. Falling back to Port 25...", state="running"
                                )
                                s = smtplib.SMTP("smtp.gmail.com", 25, timeout=60)
                                s.starttls()
                                s.login(email_user, email_pass)
                                return s
                            except Exception:
                                status_box.update(
                                    label="Port 25 unavailable. Falling back to Port 465 (SSL)...", state="running"
                                )
                                s = smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60)
                                s.login(email_user, email_pass)
                                return s
                    finally:
                        socket.socket = original_socket

                server = connect_smtp()
                status_box.update(label="SMTP connection established. Starting batch dispatch...", state="running")

                total_rows = len(df)
                sent_in_batch = 0
                pause_counter = 0

                found_links = re.findall(r"(https?://[^\s<>\"'\,]+|\{Link\})", st.session_state.text_msg)
                unique_links = list(dict.fromkeys(found_links))

                for i, (idx, row) in enumerate(df.iterrows()):
                    if not st.session_state.get("processing", False):
                        status_box.update(label="Dispatch paused by user.", state="complete", expanded=False)
                        break

                    if str(row.get("Status", "")).strip().lower() in ["sent", "enviado"]:
                        continue

                    recipient = str(row.get("Email", "")).strip()
                    if not recipient or "@" not in recipient:
                        continue

                    name_raw = row.get("Name", row.get("Nome", ""))
                    name = (
                        ""
                        if pd.isna(name_raw) or str(name_raw).strip().lower() in ["", "nan"]
                        else str(name_raw).strip()
                    )

                    course_raw = row.get("Course", row.get("Curso", "Extension Course"))
                    course = (
                        "Extension Course"
                        if pd.isna(course_raw) or str(course_raw).strip().lower() in ["", "nan"]
                        else str(course_raw).strip()
                    )

                    link_raw = row.get("Link", "https://example.org")
                    link = (
                        "https://example.org"
                        if pd.isna(link_raw) or str(link_raw).strip().lower() in ["", "nan"]
                        else str(link_raw).strip()
                    )

                    processed_in_batch = sent_in_batch + 1
                    status_box.update(
                        label=f"Sending email ({processed_in_batch}/{total_pending}): {recipient}...",
                        state="running",
                    )

                    try:
                        if pause_counter >= 50:
                            status_box.update(
                                label="2-minute anti-spam cooldown pause in progress...", state="running"
                            )
                            log_container.markdown(
                                """<div style="background-color: #0f172a; padding: 10px 14px; border-radius: 8px; margin-bottom: 8px; border: 1px solid #f59e0b; color: #f59e0b; font-size: 13px;">
                                     <b>[ANTI-SPAM COOLDOWN]</b> Holding for 2 minutes to preserve sender reputation and checking IMAP bounces...
                                </div>""",
                                unsafe_allow_html=True,
                            )
                            for _ in range(240):
                                if not st.session_state.get("processing", False):
                                    break
                                time.sleep(0.5)
                            if not st.session_state.get("processing", False):
                                status_box.update(label="Dispatch paused by user.", state="complete", expanded=False)
                                break
                            pause_counter = 0
                            server.quit()
                            try:
                                extract_bounced_emails(limit=150, update_progress=True)
                            except Exception:
                                pass
                            server = connect_smtp()

                        formatted_subject = st.session_state.subject_msg.format(
                            Name=name, Course=course, Link=link, Nome=name, Curso=course
                        )

                        modified_txt = st.session_state.text_msg
                        for idx_lnk, lnk in enumerate(unique_links):
                            btn_text = st.session_state.get(f"btn_txt_{idx_lnk}_{lnk}", "Click here to access")
                            btn_color = st.session_state.get(f"btn_color_{idx_lnk}_{lnk}", "#004B87")
                            btn_style = st.session_state.get(
                                f"btn_style_{idx_lnk}_{lnk}", "Modern Elevated (Soft Shadow & Depth)"
                            )
                            safe_btn_text = btn_text.replace("{", "{{").replace("}", "}}")
                            button_html = generate_button_html(lnk, safe_btn_text, btn_color, btn_style)
                            modified_txt = modified_txt.replace(lnk, button_html)

                        formatted_txt = modified_txt.format(
                            Name=name,
                            Course=course,
                            Link=link,
                            Image="{Image}",
                            Nome=name,
                            Curso=course,
                            Imagem="{Image}",
                        )
                        txt_html = formatted_txt

                        if st.session_state.get("external_image_url"):
                            img_url = st.session_state.external_image_url
                            size_pct = st.session_state.get("img_size_pct", 100)
                            img_tag = f'<div style="text-align: center; margin: 15px 0;"><img src="{img_url}" style="max-width: {size_pct}%; height: auto; display: block; margin: 0 auto; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); border: none;"></div>'
                            if "{Image}" in txt_html:
                                txt_html = txt_html.replace("{Image}", img_tag)
                            else:
                                txt_html += img_tag
                        else:
                            txt_html = txt_html.replace("{Image}", "")

                        msg_root = MIMEMultipart("related")
                        msg_root["From"] = email_user
                        msg_root["To"] = recipient
                        msg_root["Subject"] = formatted_subject

                        m = MIMEMultipart("alternative")
                        msg_root.attach(m)

                        m.attach(
                            MIMEText("Please enable HTML viewing in your email client to read this message.", "plain")
                        )

                        header_row_html = ""
                        if st.session_state.get("header_logo_url"):
                            header_row_html = f"""
                                            <tr>
                                                <td align="center" style="padding: 40px 20px 30px 20px; border-top: 6px solid #004B87; border-bottom: 1px solid #f1f3f5;">
                                                    <img src="{st.session_state.header_logo_url}" alt="Logo" height="85" style="display: block; max-height: 85px; border: 0; margin: 0 auto;">
                                                </td>
                                            </tr>
                            """
                        else:
                            header_row_html = '<tr><td style="border-top: 6px solid #004B87;"></td></tr>'

                        footer_label = st.session_state.get(
                            "org_footer_name", "Institutional Training & Outreach Division"
                        )

                        body_html = f"""
                        <!DOCTYPE html>
                        <html>
                        <head>
                            <meta charset="UTF-8">
                            <style>
                                .email-content p {{ margin: 0 0 5px 0 !important; }}
                                .btn-cta-relevo {{
                                    transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
                                }}
                                .btn-cta-relevo:hover {{
                                    filter: brightness(1.08);
                                    box-shadow: 0 10px 25px -2px rgba(0, 0, 0, 0.35);
                                    transform: translateY(-2px);
                                }}
                                .btn-cta-relevo:active {{
                                    transform: translateY(1px);
                                    box-shadow: 0 4px 10px -1px rgba(0, 0, 0, 0.2);
                                }}
                            </style>
                        </head>
                        <body style="margin: 0; padding: 0; background-color: #e9ecef; font-family: 'Segoe UI', Arial, sans-serif;">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #e9ecef; padding: 40px 0;">
                                <tr>
                                    <td align="center">
                                        <table border="0" cellpadding="0" cellspacing="0" width="800" style="background-color: #ffffff; border-radius: 8px; border: 1px solid #dee2e6; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.05);">
                                            {header_row_html}
                                            <tr>
                                                <td class="email-content" style="padding: 40px 50px; color: #343a40; font-size: 16px; line-height: 1.6;">
                                                    {txt_html}
                                                </td>
                                            </tr>
                                            <tr>
                                                <td align="center" style="background-color: #f8f9fa; padding: 30px 40px; border-top: 1px solid #e9ecef;">
                                                    <p style="margin: 0; font-size: 13px; color: #6c757d; line-height: 1.6;">
                                                        <strong>{footer_label}</strong><br>
                                                        Automated Outreach &amp; Communications<br>
                                                    </p>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </body>
                        </html>
                        """

                        m.attach(MIMEText(body_html, "html"))

                        server.sendmail(email_user, recipient, msg_root.as_string())
                        sent_in_batch += 1
                        pause_counter += 1

                        wait_seconds = random.randint(5, 12)
                        for _ in range(wait_seconds * 2):
                            if not st.session_state.get("processing", False):
                                break
                            time.sleep(0.5)

                        df.at[idx, "Status"] = "Sent"
                        st.session_state.dataframe = df
                        register_sent_email(recipient)

                        log_name = f"<b>{name}</b> " if name else ""
                        log_container.markdown(
                            f"""<div style="background-color: #0f172a; padding: 10px 14px; border-radius: 8px; margin-bottom: 8px; border: 1px solid #334155; color: #f8fafc; font-size: 13px;">
                                <span style="color: #4ade80; font-weight: 600;">[DELIVERED] [{sent_in_batch}/{total_pending}]:</span> {log_name}<span style="color: #60a5fa;">({recipient})</span>
                            </div>""",
                            unsafe_allow_html=True,
                        )

                    except smtplib.SMTPServerDisconnected:
                        status_box.update(label="Connection dropped. Reconnecting to SMTP...", state="running")
                        server = connect_smtp()
                        server.sendmail(email_user, recipient, msg_root.as_string())

                        df.at[idx, "Status"] = "Sent"
                        st.session_state.dataframe = df
                        register_sent_email(recipient)
                        sent_in_batch += 1
                        pause_counter += 1

                        log_name = f"<b>{name}</b> " if name else ""
                        log_container.markdown(
                            f"""<div style="background-color: #0f172a; padding: 10px 14px; border-radius: 8px; margin-bottom: 8px; border: 1px solid #334155; color: #f8fafc; font-size: 13px;">
                                <span style="color: #4ade80; font-weight: 600;">[DELIVERED - RECONNECTED] [{sent_in_batch}/{total_pending}]:</span> {log_name}<span style="color: #60a5fa;">({recipient})</span>
                            </div>""",
                            unsafe_allow_html=True,
                        )
                        wait_seconds = random.randint(5, 12)
                        for _ in range(wait_seconds * 2):
                            if not st.session_state.get("processing", False):
                                break
                            time.sleep(0.5)

                    except Exception as send_err:
                        df.at[idx, "Status"] = "Error"
                        st.session_state.dataframe = df
                        register_blocked_email(recipient)
                        pause_counter += 1
                        log_container.markdown(
                            f"""<div style="background-color: #0f172a; padding: 10px 14px; border-radius: 8px; margin-bottom: 8px; border: 1px solid #f87171; color: #f8fafc; font-size: 13px;">
                                <span style="color: #f87171; font-weight: 600;">[FAILED] [{processed_in_batch}/{total_pending}]:</span> <span style="color: #60a5fa;">({recipient})</span> - {send_err}
                            </div>""",
                            unsafe_allow_html=True,
                        )

                    # Save checkpoint after each recipient for fault-tolerant recovery
                    df.to_excel("dispatch_progress.xlsx", index=False)
                    progress_bar.progress((i + 1) / total_rows)

                if server:
                    try:
                        server.quit()
                    except Exception:
                        pass

                status_box.update(
                    label="Batch completed. Verifying delivery receipts via Gmail IMAP...", state="running"
                )
                time.sleep(4)
                try:
                    extract_bounced_emails(limit=300, update_progress=True)
                except Exception:
                    pass
                status_box.update(label="Queue processing completed.", state="complete", expanded=False)

                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Delivery_Report")
                    worksheet = writer.sheets["Delivery_Report"]

                    header_font = Font(bold=True, color="FFFFFF")
                    header_fill = PatternFill(start_color="004B87", end_color="004B87", fill_type="solid")
                    for cell in worksheet[1]:
                        cell.font = header_font
                        cell.fill = header_fill

                st.session_state.final_report = output.getvalue()
                st.session_state.dataframe = df
                st.session_state.dispatch_completed = True

            except Exception as e:
                status_box.update(label="Critical Error", state="error")
                st.error(f"Operation failed: {e}")

            finally:
                st.session_state.processing = False
                st.rerun()

        # Bottom Panel: Report Export & Bounce Synchronization
        st.write("")
        with st.container(border=True):
            st.markdown('<div class="section-label">Report Export &amp; Mailbox Maintenance</div>', unsafe_allow_html=True)

            col_exp1, col_exp2 = st.columns(2, gap="large")

            with col_exp1:
                st.markdown(
                    '<p style="font-size: 13px; color: #94a3b8; margin-bottom: 8px;">Consolidated Delivery Spreadsheet</p>',
                    unsafe_allow_html=True,
                )
                if st.session_state.final_report:
                    st.download_button(
                        label="Download Final Report (Excel)",
                        data=st.session_state.final_report,
                        file_name="delivery_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
                elif os.path.exists("dispatch_progress.xlsx"):
                    with open("dispatch_progress.xlsx", "rb") as f:
                        progress_bytes = f.read()
                    st.download_button(
                        label="Download Queue Checkpoint (Excel)",
                        data=progress_bytes,
                        file_name="dispatch_progress.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
                else:
                    st.button("Report Unavailable (No dispatches yet)", disabled=True, use_container_width=True)

            with col_exp2:
                st.markdown(
                    '<p style="font-size: 13px; color: #94a3b8; margin-bottom: 8px;">Automated Bounce / NDR Verification</p>',
                    unsafe_allow_html=True,
                )
                if st.button("Sync Gmail Bounces", use_container_width=True, key="btn_sync_tab3"):
                    with st.spinner("Connecting to Gmail via IMAP and scanning mailboxes..."):
                        new_bounces = extract_bounced_emails(limit=300, update_progress=True)
                        st.success(f"Synchronization complete. {new_bounces} bounce(s) detected and filtered.")
                        time.sleep(1)
                        st.rerun()

st.markdown(
    """
    <style>
        .block-container {
            padding-bottom: 60px !important;
        }
        .custom-footer {
            position: fixed;
            bottom: 0;
            left: 0;
            width: 100%;
            background: rgba(15, 23, 42, 0.65);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            color: #64748b;
            text-align: center;
            padding: 8px 0;
            border-top: 1px solid rgba(51, 65, 85, 0.3);
            font-size: 12px;
            z-index: 9999;
        }
        .custom-footer p {
            margin: 0;
            font-weight: 400;
            letter-spacing: 0.2px;
        }
        .custom-footer span {
            color: #3b82f6;
            font-weight: 500;
            opacity: 0.9;
        }
    </style>
    <div class="custom-footer">
        <p>Developed by <span>Arthur Medeiros</span> | 2026</p>
    </div>
    """,
    unsafe_allow_html=True,
)
