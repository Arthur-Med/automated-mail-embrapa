<div align="center">
  <img src="https://readme-typing-svg.demolab.com?font=Fira+Code&weight=700&size=30&duration=3000&pause=1200&color=38BDF8&center=true&vCenter=true&width=650&lines=Automated+Mail+Sender;Bulk+Email+%26+Campaign+Automation" alt="Automated Mail Sender" />

  <p>
    <img src="https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"/>
    <img src="https://img.shields.io/badge/Streamlit-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white" alt="Streamlit"/>
    <img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas"/>
    <img src="https://img.shields.io/badge/SMTP%20%2F%20IMAP-Automation-004B87?style=for-the-badge" alt="SMTP/IMAP"/>
  </p>
</div>

---

A web-based bulk email dispatch system built with Python and Streamlit. Originally developed at **Embrapa Dairy Cattle** (*Brazilian Agricultural Research Corporation*) to automate institutional training outreach and event communications across restricted corporate networks.

This repository contains the white-label version of the application, stripped of internal institutional assets and mailing lists.

## Features

- **Anti-Spam Throttling:** Randomized delays (5–12 seconds) between dispatches and an automatic 2-minute cooldown every 50 emails with SMTP session renewal to avoid rate limits.
- **IMAP Bounce Synchronization:** Scans Gmail (`INBOX` and `Spam`) for `mailer-daemon` failure reports (RFC 822), automatically moving invalid addresses to a local blocklist and updating the delivery spreadsheet.
- **Corporate Proxy Support:** Built-in HTTP proxy tunneling (`PySocks`) and automatic SMTP port fallback (`587 STARTTLS` -> `25 STARTTLS` -> `465 SSL`) for firewalled networks.
- **HTML Template Editor:** Integrated Quill WYSIWYG editor with live preview, customizable call-to-action buttons, and dynamic spreadsheet placeholders (`{Name}`, `{Course}`, `{Link}`).
- **Fault-Tolerant Queue:** Saves delivery state (`dispatch_progress.xlsx`) after every email, allowing campaigns to be paused, resumed, and exported as formatted Excel reports.

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Create a `.env` file in the root directory (see `.env.example`):
```env
GMAIL_USER=your.email@example.com
GMAIL_APP_PASSWORD=your_app_password

# Optional: fill only if running behind a corporate HTTP proxy
PROXY_HOST=
PROXY_PORT=3128
PROXY_USER=
PROXY_PASS=
```

3. Run the application:
```bash
streamlit run app.py
```

A fictional contact list (`sample_contacts.xlsx`) is included for testing the workflow.
