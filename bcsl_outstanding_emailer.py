#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BCSL Outstanding-Email Generator  (clients + sales team)
========================================================
One script, two jobs, both driven from the same Tally export
"Client_OS_Report.xlsx" (a "Sundry Debtors / Pending Bills" report):

  * job "client" -- e-mails each CLIENT a statement of its pending invoices,
        split into ">30 days overdue" and "<=30 days overdue".
        Template: "Outstanding - <Client>".  Addresses: client address file.

  * job "sales"  -- e-mails each SALES PERSON ONE consolidated list of every
        outstanding invoice ABOVE 35 days across all of their clients.
        Template: "Outstanding Bills more than 35 days".  Addresses: sales file.

----------------------------------------------------------------------
CHOOSING WHAT TO RUN  (run them together or separately, any time)
----------------------------------------------------------------------
    python bcsl_outstanding_emailer.py            # uses RUN_JOBS below (default: both)
    python bcsl_outstanding_emailer.py client     # only the client e-mails
    python bcsl_outstanding_emailer.py sales       # only the sales-person e-mails
    python bcsl_outstanding_emailer.py both        # both

You can also override the delivery mode for a single run, e.g. for scheduling:
    python bcsl_outstanding_emailer.py sales  --transport smtp
    python bcsl_outstanding_emailer.py client --transport outlook

----------------------------------------------------------------------
DELIVERY MODES  (TRANSPORT)
----------------------------------------------------------------------
  "dry_run" : SAFE DEFAULT. Sends nothing. Writes .eml + .html previews and a
              manifest.csv per job under ./outstanding_emails_output/<job>/.
              Review them, then switch to a real transport.
  "smtp"    : send through an SMTP server (fill in the SMTP_* / FROM_* fields).
  "outlook" : send through the locally-installed Outlook desktop app (Windows,
              needs `pip install pywin32`).

Setup:
    pip install openpyxl
    pip install pywin32        # only for the "outlook" transport
"""

import os
import re
import csv
import sys
import html
import json
import time
import smtplib
import logging
import argparse
import datetime
import mimetypes
from collections import OrderedDict
from email.message import EmailMessage
from email.utils import formataddr, make_msgid

from openpyxl import load_workbook

# =====================================================================
# CONFIG  -- edit .env to change these values in production
# =====================================================================

# ---- which job(s) to run by default (overridden by the command-line arg) ----
RUN_JOBS = "both"           # "client" | "sales" | "both"

# ---- shared input / output ----
REPORT_FILE = "Client_OS_Report.xlsx"
ASSETS_DIR  = "assets"                       # logo + social icons
OUTPUT_DIR  = "outstanding_emails_output"    # previews + manifests land here

# ---- delivery ----
TRANSPORT = "dry_run"        # "dry_run" | "smtp" | "outlook"
EMAIL_DELAY_SECONDS = 2.0   # delay in seconds between sending emails

# ---- matching safety (applies to both jobs) ----
SEND_ONLY_WHEN_FLAG_Y = True
ALLOW_FUZZY_MATCH     = False
FUZZY_CUTOFF          = 0.92

# ---- SMTP (only used when TRANSPORT == "smtp") ----
SMTP_HOST     = "smtp.office365.com"
SMTP_PORT     = 587
SMTP_SECURITY = "starttls"   # "starttls" | "ssl" | "none"
SMTP_USERNAME = "accounts@benchmarksolution.com"
SMTP_PASSWORD = "PUT-PASSWORD-OR-APP-PASSWORD-HERE"

# ---- shared company / signature block ----
COMPANY_NAME = "Benchmark Computer Solutions Limited"
COMPANY_ADDR = [
    "501, 5th Floor, Kushwah Chambers, Makhawana Road,",
    "Marol Naka, Andheri East, Mumbai - 400059",
]
COMPANY_CERT = "ISO 9001:2015 CERTIFIED"
SOCIAL_LINKS = [   # (icon-file-in-assets, url)
    ("image002.png", "https://www.facebook.com/benchmarkcomputersolutions/"),
    ("image003.png", "https://www.linkedin.com/company/benchmarkcomputersolutions/"),
    ("image004.png", "https://twitter.com/BCSPLMumbai"),
    ("image005.png", "https://www.instagram.com/benchmarkcomputersolutions/"),
]
LOGO_FILE = "image001.png"

# =====================================================================
#  JOB 1 -- CLIENT  (Outstanding - <Client>)
# =====================================================================
CLIENT_ADDRESS_FILE   = "bcsl-client-outstanding-email-addresses.xlsx"
CLIENT_FROM_NAME      = "Accounts - Benchmark Computer Solutions"
CLIENT_FROM_EMAIL     = "accounts@benchmarksolution.com"
CLIENT_SUBJECT        = "Outstanding - {name}"
CLIENT_THRESHOLD_DAYS = 30
# Section headings. The sample puts older bills first; if your convention is the
# other way round, just swap these two strings.
CLIENT_GREETING       = "Dear Sir/Mam,"
CLIENT_INTRO_OVER_30  = "Please Update Payment Status,"          # > 30 days
CLIENT_INTRO_UPTO_30  = "Please confirm below invoice with your account.,"  # <= 30 days
CLIENT_SIGNER = dict(name="Dipesh Thakur", title="Accounts Executive",
                     direct="+91 (022) 40822105", main="+91 (022) 40822100")
SHOW_OVERDUE_BARS = True

# =====================================================================
#  JOB 2 -- SALES PERSON  (Outstanding Bills more than 35 days)
# =====================================================================
SALES_ADDRESS_FILE   = "bcsl-salesteam-email-addresses.xlsx"
SALES_FROM_NAME      = "Mohit P - Benchmark Computer Solutions"
SALES_FROM_EMAIL     = "mohit.p@benchmarksolution.com"
SALES_SUBJECT        = "Outstanding Bills more than 35 days"
SALES_THRESHOLD_DAYS = 35                                  # bills strictly ABOVE 35 days
SALES_GREETING       = "Dear Sir,"
SALES_INTRO          = ("Please find below outstanding bills above 35 days. "
                        "Please revert with updated status and expected payment dates.")
SALES_SIGNER = dict(name="Mohit Pusalkar", title="Manager-Accounts & Compliance",
                    direct="+91-22-40822122", main="+91 (022) 40822100")

# =====================================================================
#  end of CONFIG
# =====================================================================

CONFIG_SCHEMA = [
    ("RUN_JOBS", RUN_JOBS, "Which job(s) to run by default. Valid options: 'client' (emails outstanding invoice lists to clients), 'sales' (emails consolidated outstanding client lists to sales managers), 'both' (runs both client and sales jobs sequentially)"),
    ("REPORT_FILE", REPORT_FILE, "Name of the Tally outstanding Excel report file. Must reside in the same directory as the executable"),
    ("ASSETS_DIR", ASSETS_DIR, "Directory name where images and logos are stored (default: assets)"),
    ("OUTPUT_DIR", OUTPUT_DIR, "Directory name where program outputs (previews, manifests, log file) are stored (default: outstanding_emails_output)"),
    ("TRANSPORT", TRANSPORT, "Delivery transport mode. Valid options: 'dry_run' (generates draft HTML/EML previews locally under output folder), 'smtp' (sends emails directly using SMTP server configuration), 'outlook' (sends emails via local Windows Outlook desktop application using your signed-in account)"),
    ("EMAIL_DELAY_SECONDS", EMAIL_DELAY_SECONDS, "Time delay (in seconds) to pause between sending each email. Helps avoid spam filters or rate limiting. Options: any number (e.g. 0 to disable, 2.5, 5, etc.)"),
    ("SEND_ONLY_WHEN_FLAG_Y", SEND_ONLY_WHEN_FLAG_Y, "Whether to only send emails to entries marked 'Y' or 'y' in the 'Send Email (Y/N)' column. Valid options: True, False"),
    ("ALLOW_FUZZY_MATCH", ALLOW_FUZZY_MATCH, "Whether to allow fuzzy name matching between Tally reports and address books. Valid options: True, False"),
    ("FUZZY_CUTOFF", FUZZY_CUTOFF, "Cut-off threshold score for fuzzy matching. Value between 0.0 (match anything) and 1.0 (exact match)"),
    ("SMTP_HOST", SMTP_HOST, "SMTP server host address (e.g. smtp.gmail.com or smtp.office365.com)"),
    ("SMTP_PORT", SMTP_PORT, "SMTP server port number (e.g., 587 for STARTTLS, 465 for SSL/TLS)"),
    ("SMTP_SECURITY", SMTP_SECURITY, "SMTP connection security protocol. Valid options: 'starttls', 'ssl', 'none'"),
    ("SMTP_USERNAME", SMTP_USERNAME, "Username credential for the SMTP login (e.g. your email address)"),
    ("SMTP_PASSWORD", SMTP_PASSWORD, "Password or App Password credential for SMTP login. For Gmail or Office 365, generate a 16-character App Password"),
    ("COMPANY_NAME", COMPANY_NAME, "Company's official name used in signatures"),
    ("COMPANY_ADDR", COMPANY_ADDR, "Company's physical address. Must be a JSON array of strings, e.g. [\"Line 1\", \"Line 2\"]"),
    ("COMPANY_CERT", COMPANY_CERT, "Company ISO or certification text line in signatures"),
    ("SOCIAL_LINKS", SOCIAL_LINKS, "Social media links. Format: a JSON array of lists, each containing [\"icon_filename.png\", \"profile_url\"]"),
    ("LOGO_FILE", LOGO_FILE, "Name of the company logo image file in the assets folder"),
    ("CLIENT_ADDRESS_FILE", CLIENT_ADDRESS_FILE, "Name of the Excel address book spreadsheet containing client email addresses"),
    ("CLIENT_FROM_NAME", CLIENT_FROM_NAME, "Display name in the 'From' field for client outstanding emails"),
    ("CLIENT_FROM_EMAIL", CLIENT_FROM_EMAIL, "Email address in the 'From' field for client outstanding emails"),
    ("CLIENT_SUBJECT", CLIENT_SUBJECT, "Subject line template for client outstanding emails. '{name}' placeholder is replaced with client name"),
    ("CLIENT_THRESHOLD_DAYS", CLIENT_THRESHOLD_DAYS, "Age threshold (days) for client overdue splitting. Bills older fall under OVER_30 (overdue), newer fall under UPTO_30 (recent/unconfirmed)"),
    ("CLIENT_GREETING", CLIENT_GREETING, "Greeting text at the start of client emails (e.g. Dear Sir/Mam,)"),
    ("CLIENT_INTRO_OVER_30", CLIENT_INTRO_OVER_30, "Introduction line header above the table for invoices strictly older than CLIENT_THRESHOLD_DAYS"),
    ("CLIENT_INTRO_UPTO_30", CLIENT_INTRO_UPTO_30, "Introduction line header above the table for invoices newer/equal to CLIENT_THRESHOLD_DAYS"),
    ("CLIENT_SIGNER", CLIENT_SIGNER, "Contact details for client email signer. Must be a JSON dictionary containing 'name', 'title', 'direct', and 'main' keys"),
    ("SHOW_OVERDUE_BARS", SHOW_OVERDUE_BARS, "Whether to show color-coded overdue summary bars at the top of client emails. Valid options: True, False"),
    ("SALES_ADDRESS_FILE", SALES_ADDRESS_FILE, "Name of the Excel address book spreadsheet containing sales person email addresses"),
    ("SALES_FROM_NAME", SALES_FROM_NAME, "Display name in the 'From' field for sales consolidated emails"),
    ("SALES_FROM_EMAIL", SALES_FROM_EMAIL, "Email address in the 'From' field for sales consolidated emails"),
    ("SALES_SUBJECT", SALES_SUBJECT, "Subject line text for sales consolidated emails"),
    ("SALES_THRESHOLD_DAYS", SALES_THRESHOLD_DAYS, "Invoice age threshold (days) for sales consolidated reports. Only groups invoices strictly older than this threshold"),
    ("SALES_GREETING", SALES_GREETING, "Greeting text at the start of sales salesperson emails (e.g. Dear Sir,)"),
    ("SALES_INTRO", SALES_INTRO, "Introduction line for salesperson consolidated emails explaining the attached outstanding list"),
    ("SALES_SIGNER", SALES_SIGNER, "Contact details for sales email signer. Must be a JSON dictionary containing 'name', 'title', 'direct', and 'main' keys")
]

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def get_absolute_path(rel_path):
    if not rel_path:
        return ""
    if os.path.isabs(rel_path):
        return rel_path
    return os.path.join(get_base_dir(), rel_path)

def load_config():
    env_path = get_absolute_path(".env")
    if not os.path.exists(env_path):
        try:
            with open(env_path, "w", encoding="utf-8") as f:
                f.write("# BCSL Outstanding-Emailer Configuration\n")
                f.write("# =======================================\n\n")
                for name, default, desc in CONFIG_SCHEMA:
                    f.write(f"# {desc}\n")
                    if isinstance(default, (list, dict)):
                        val_str = json.dumps(default)
                    else:
                        val_str = str(default)
                    f.write(f"{name}={val_str}\n\n")
        except Exception as e:
            print(f"Error creating default .env: {e}", file=sys.stderr)
            return

    try:
        user_config = {}
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    k, v = line.split("=", 1)
                    user_config[k.strip()] = v.strip()
        
        g = globals()
        default_dict = {item[0]: item[1] for item in CONFIG_SCHEMA}
        for k, v in user_config.items():
            if k in default_dict:
                default_val = default_dict[k]
                if isinstance(default_val, bool):
                    g[k] = v.lower() in ("true", "1", "yes")
                elif isinstance(default_val, int):
                    g[k] = int(v)
                elif isinstance(default_val, float):
                    g[k] = float(v)
                elif isinstance(default_val, (list, dict)):
                    parsed = json.loads(v)
                    if k == "SOCIAL_LINKS":
                        g[k] = [tuple(link) for link in parsed]
                    else:
                        g[k] = parsed
                else:
                    g[k] = v
    except Exception as e:
        print(f"Error loading .env file: {e}", file=sys.stderr)

log = logging.getLogger("bcsl")


def setup_logging():
    os.makedirs(get_absolute_path(OUTPUT_DIR), exist_ok=True)
    if log.handlers:
        return
    log.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s  %(levelname)-7s %(message)s", "%H:%M:%S")
    sh = logging.StreamHandler(sys.stdout); sh.setFormatter(fmt); log.addHandler(sh)
    fh = logging.FileHandler(get_absolute_path(os.path.join(OUTPUT_DIR, "run.log")), encoding="utf-8")
    fh.setFormatter(fmt); log.addHandler(fh)


# ---------------------------------------------------------------------
#  small text helpers
# ---------------------------------------------------------------------
def clean_text(s):
    if s is None:
        return ""
    s = str(s).replace("_x000D_", "")
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    return s.strip()


def normalize_name(s):
    s = (s or "").replace("_x000D_", "").lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def fmt_amount(x):
    try:
        return "{:,.2f} Dr".format(float(x))
    except (TypeError, ValueError):
        return ""


def html_cell(text):
    return html.escape(clean_text(text)).replace("\n", "<br>")


def split_emails(raw):
    if raw is None:
        return []
    out = []
    for p in re.split(r"[;,]", str(raw)):
        p = p.strip().strip("<>").strip()
        if p and "@" in p:
            out.append(p)
    return out


def safe_filename(name):
    return re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")[:60] or "item"


# ---------------------------------------------------------------------
#  read the OS report  ->  OrderedDict {party: [invoice, ...]}
#  each invoice also carries inv["party"] so it can be regrouped by sales person
# ---------------------------------------------------------------------
def parse_report(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    as_on = None
    try:
        m = re.search(r"to\s+(.+)$", str(ws["D1"].value or ""))
        if m:
            as_on = datetime.datetime.strptime(m.group(1).strip(), "%d-%b-%y")
    except Exception:
        pass
    if as_on is None:
        as_on = datetime.datetime.now()

    clients, current, started = OrderedDict(), None, False
    for row in ws.iter_rows(values_only=True):
        A, B, C, D, E, F, G, H = (list(row) + [None] * 8)[:8]

        if A == "Date" and C == "Party's Name":
            started = True
            continue
        if not started:
            continue

        if isinstance(A, datetime.datetime):                      # invoice row
            pending = F if F is not None else E
            if pending is None:
                continue
            try:
                pending = float(pending)
            except (TypeError, ValueError):
                continue
            if abs(pending) < 0.005:
                continue
            try:
                opening = float(E) if E is not None else pending
            except (TypeError, ValueError):
                opening = pending

            due = G if isinstance(G, datetime.datetime) else None
            try:
                overdue = int(float(H)) if H not in (None, "") else None
            except (TypeError, ValueError):
                overdue = None
            if overdue is None and due is not None:
                overdue = (as_on - due).days

            if current is None:
                log.warning("Invoice %s before any client header - skipped", B)
                continue
            clients[current].append(dict(
                party=current, date=A, ref=clean_text(B),
                particulars=clean_text(C), sale=clean_text(D),
                opening=opening, pending=pending, due=due,
                overdue=overdue if overdue is not None else 0,
            ))

        elif A is None and C and isinstance(C, str) and C.strip() and (E is None or E == ""):
            current = clean_text(C)
            clients.setdefault(current, [])

    return clients


# ---------------------------------------------------------------------
#  read an address book (same layout for clients and sales team):
#  col0 = name | col1 = Send (Y/N) | col2 = TO | col3 = CC | col4 = BCC
# ---------------------------------------------------------------------
def load_address_book(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    next(rows, None)                                              # header
    book = {}
    for row in rows:
        name = clean_text(row[0]) if len(row) > 0 else ""
        if not name:
            continue
        flag = (str(row[1]).strip().lower()
                if len(row) > 1 and row[1] is not None else "")
        book[normalize_name(name)] = dict(
            raw=name,
            send=flag.startswith("y"),
            to=split_emails(row[2] if len(row) > 2 else None),
            cc=split_emails(row[3] if len(row) > 3 else None),
            bcc=split_emails(row[4] if len(row) > 4 else None),
        )
    return book


def match_entry(name, book):
    """(entry, 'exact'|'fuzzy')  or  (None, closest_name_hint_or_None)."""
    key = normalize_name(name)
    if key in book:
        return book[key], "exact"
    import difflib
    if ALLOW_FUZZY_MATCH:
        hit = difflib.get_close_matches(key, list(book), n=1, cutoff=FUZZY_CUTOFF)
        if hit:
            return book[hit[0]], "fuzzy"
    near = difflib.get_close_matches(key, list(book), n=1, cutoff=0.6)
    return None, (book[near[0]]["raw"] if near else None)


def resolve_recipient(name, book):
    """Return (entry, how, skip_reason). skip_reason is None when good to send."""
    entry, how = match_entry(name, book)
    if entry is None:
        return None, how, "no-address"
    if SEND_ONLY_WHEN_FLAG_Y and not entry["send"]:
        return entry, how, "flag-N"
    if not entry["to"]:
        return entry, how, "no-TO"
    return entry, how, None


# ---------------------------------------------------------------------
#  HTML building blocks (shared)
# ---------------------------------------------------------------------
_BASE_TD = ("border:1px solid #000;padding:2px 7px;"
            "font-family:Calibri,Arial,sans-serif;font-size:11pt;"
            "color:#000;vertical-align:top;")


def _td(content, align="left", width=None, *, bg=None, bold=False):
    style = _BASE_TD + "text-align:%s;" % align
    if width:
        style += "width:%dpx;" % width
    if bg:
        style += "background:%s;" % bg
    inner = "<b>%s</b>" % content if (bold and content) else content
    return '<td style="%s">%s</td>' % (style, inner or "&nbsp;")


def build_signature_html(signer):
    p = "font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000;margin:0;"
    L = ['<div style="margin-top:18px;">']
    L.append('<p style="%s">Thanks &amp; Regards,</p>' % p)
    L.append('<p style="%s">&nbsp;</p>' % p)
    L.append('<p style="%s"><b>%s</b></p>' % (p, html.escape(signer["name"])))
    L.append('<p style="%s">%s</p>' % (p, html.escape(signer["title"])))
    L.append('<p style="%s">Direct: %s</p>' % (p, html.escape(signer["direct"])))
    L.append('<p style="%s">Main: %s</p>' % (p, html.escape(signer["main"])))
    L.append('<p style="%s">&nbsp;</p>' % p)
    if os.path.exists(get_absolute_path(os.path.join(ASSETS_DIR, LOGO_FILE))):
        L.append('<p style="margin:6px 0;"><img src="cid:bcsl_logo" alt="%s" '
                 'style="height:48px;"></p>' % html.escape(COMPANY_NAME))
    L.append('<p style="%s"><b>%s</b></p>' % (p, html.escape(COMPANY_NAME)))
    for a in COMPANY_ADDR:
        L.append('<p style="%s">%s</p>' % (p, html.escape(a)))
    L.append('<p style="%s">%s</p>' % (p, html.escape(COMPANY_CERT)))
    icons = []
    for idx, (icon, url) in enumerate(SOCIAL_LINKS):
        if os.path.exists(get_absolute_path(os.path.join(ASSETS_DIR, icon))):
            icons.append('<a href="%s"><img src="cid:soc%d" '
                         'style="height:22px;border:0;margin-right:6px;"></a>'
                         % (html.escape(url), idx))
    if icons:
        L.append('<p style="margin:8px 0;">' + "".join(icons) + "</p>")
    L.append("</div>")
    return "".join(L)


# ---------------------------------------------------------------------
#  CLIENT e-mail body (two ageing sections, grouped under the client name)
# ---------------------------------------------------------------------
_CLIENT_COLS = [("date", "right", 70), ("ref", "left", 150),
                ("particulars", "left", 235), ("sale", "left", 110),
                ("opening", "right", 100), ("pending", "right", 100)]


def _client_table(client_name, invoices):
    rows = []
    hdr = []
    for i, (key, align, w) in enumerate(_CLIENT_COLS):
        if i == 2:
            hdr.append(_td(html.escape(client_name), align, w, bg="#D9D9D9", bold=True))
        else:
            hdr.append(_td("", align, w, bg="#D9D9D9"))
    rows.append("<tr>" + "".join(hdr) + "</tr>")

    to = tp = 0.0
    for inv in invoices:
        to += inv["opening"]; tp += inv["pending"]
        cells = []
        for key, align, w in _CLIENT_COLS:
            if key == "date":
                v = inv["date"].strftime("%d-%b-%y")
            elif key == "opening":
                v = fmt_amount(inv["opening"])
            elif key == "pending":
                v = fmt_amount(inv["pending"])
            else:
                v = html_cell(inv[key])
            cells.append(_td(v, align, w))
        rows.append("<tr>" + "".join(cells) + "</tr>")

    tcells = []
    for key, align, w in _CLIENT_COLS:
        if key == "opening":
            tcells.append(_td(fmt_amount(to), align, w, bold=True))
        elif key == "pending":
            tcells.append(_td(fmt_amount(tp), align, w, bold=True))
        else:
            tcells.append(_td("", align, w))
    rows.append("<tr>" + "".join(tcells) + "</tr>")
    return ('<table cellpadding="0" cellspacing="0" '
            'style="border-collapse:collapse;margin:6px 0 14px 0;">'
            + "".join(rows) + "</table>")


def _build_overdue_bars_html(invoices):
    # brackets
    b1 = sum(i["pending"] for i in invoices if i["overdue"] < 30)
    b2 = sum(i["pending"] for i in invoices if 30 <= i["overdue"] < 60)
    b3 = sum(i["pending"] for i in invoices if 60 <= i["overdue"] < 90)
    b4 = sum(i["pending"] for i in invoices if 90 <= i["overdue"] < 120)
    b5 = sum(i["pending"] for i in invoices if 120 <= i["overdue"])

    # formatting
    f1 = fmt_amount(b1) or "0.00 Dr"
    f2 = fmt_amount(b2) or "0.00 Dr"
    f3 = fmt_amount(b3) or "0.00 Dr"
    f4 = fmt_amount(b4) or "0.00 Dr"
    f5 = fmt_amount(b5) or "0.00 Dr"

    html_str = """
<table cellpadding="0" cellspacing="4" style="border-collapse:separate;width:100%;max-width:750px;margin:10px 0 20px 0;font-family:Calibri,Arial,sans-serif;">
  <tr>
    <td style="background:#2e7d32;color:#ffffff;padding:8px 6px;text-align:center;border-radius:4px;width:20%;">
      <div style="font-size:8pt;font-weight:bold;margin-bottom:4px;opacity:0.9;">&lt;30Days</div>
      <div style="font-size:10.5pt;font-weight:bold;white-space:nowrap;">{val1}</div>
    </td>
    <td style="background:#689f38;color:#ffffff;padding:8px 6px;text-align:center;border-radius:4px;width:20%;">
      <div style="font-size:8pt;font-weight:bold;margin-bottom:4px;opacity:0.9;">30 to 60 days</div>
      <div style="font-size:10.5pt;font-weight:bold;white-space:nowrap;">{val2}</div>
    </td>
    <td style="background:#fbc02d;color:#000000;padding:8px 6px;text-align:center;border-radius:4px;width:20%;">
      <div style="font-size:8pt;font-weight:bold;margin-bottom:4px;opacity:0.9;">60 to 90 days</div>
      <div style="font-size:10.5pt;font-weight:bold;white-space:nowrap;">{val3}</div>
    </td>
    <td style="background:#e65100;color:#ffffff;padding:8px 6px;text-align:center;border-radius:4px;width:20%;">
      <div style="font-size:8pt;font-weight:bold;margin-bottom:4px;opacity:0.9;">90 to 120 days</div>
      <div style="font-size:10.5pt;font-weight:bold;white-space:nowrap;">{val4}</div>
    </td>
    <td style="background:#c62828;color:#ffffff;padding:8px 6px;text-align:center;border-radius:4px;width:20%;">
      <div style="font-size:8pt;font-weight:bold;margin-bottom:4px;opacity:0.9;">120 to 180 days</div>
      <div style="font-size:10.5pt;font-weight:bold;white-space:nowrap;">{val5}</div>
    </td>
  </tr>
</table>
""".format(val1=f1, val2=f2, val3=f3, val4=f4, val5=f5)
    return html_str


def build_client_html(client_name, over, under):
    p = ("font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000;"
         "margin:0 0 10px 0;")
    B = ['<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000;">']
    B.append('<p style="%s">%s</p>' % (p, html.escape(CLIENT_GREETING)))
    if SHOW_OVERDUE_BARS:
        invoices = over + under
        B.append(_build_overdue_bars_html(invoices))
    for intro, invs in ((CLIENT_INTRO_OVER_30, over), (CLIENT_INTRO_UPTO_30, under)):
        if not invs:
            continue
        B.append('<p style="%s">%s</p>' % (p, html.escape(intro)))
        B.append(_client_table(client_name, invs))
    B.append(build_signature_html(CLIENT_SIGNER))
    B.append("</div>")
    return "".join(B)


def build_client_plain(client_name, over, under):
    out = [CLIENT_GREETING, ""]
    for intro, invs in ((CLIENT_INTRO_OVER_30, over), (CLIENT_INTRO_UPTO_30, under)):
        if not invs:
            continue
        out += [intro, client_name]
        to = tp = 0.0
        for inv in invs:
            to += inv["opening"]; tp += inv["pending"]
            out.append("  {}  {}  {}  {}  {}  {}".format(
                inv["date"].strftime("%d-%b-%y"), inv["ref"],
                clean_text(inv["particulars"]).split("\n")[0], inv["sale"],
                fmt_amount(inv["opening"]), fmt_amount(inv["pending"])))
        out.append("  TOTAL: {}  {}".format(fmt_amount(to), fmt_amount(tp)))
        out.append("")
    out += _plain_sig(CLIENT_SIGNER)
    return "\n".join(out)


# ---------------------------------------------------------------------
#  SALES e-mail body (single >35-day table, columns incl. Party + Remarks)
# ---------------------------------------------------------------------
_SALES_COLS = [("date", "Date", "left", 70), ("ref", "Ref. No.", "left", 130),
               ("party", "Party's Name", "left", 185),
               ("particulars", "Description", "left", 220),
               ("opening", "Opening", "right", 90),
               ("pending", "Pending", "right", 95),
               ("remarks", "Remarks", "left", 90),
               ("overdue", "Overdue", "right", 70)]


def _sales_table(invoices):
    rows = []
    # header row (grey, bold)
    rows.append("<tr>" + "".join(
        _td(html.escape(title), align, w, bg="#D9D9D9", bold=True)
        for key, title, align, w in _SALES_COLS) + "</tr>")

    tp = 0.0
    for inv in invoices:
        tp += inv["pending"]
        cells = []
        for key, title, align, w in _SALES_COLS:
            if key == "date":
                v = inv["date"].strftime("%d-%b-%y")
            elif key == "opening":
                v = fmt_amount(inv["opening"])
            elif key == "pending":
                v = fmt_amount(inv["pending"])
            elif key == "overdue":
                v = str(inv["overdue"])
            elif key == "remarks":
                v = ""
            else:
                v = html_cell(inv[key])
            cells.append(_td(v, align, w))
        rows.append("<tr>" + "".join(cells) + "</tr>")

    # total row (grey) - only the Pending total, matching the sample
    tcells = []
    for key, title, align, w in _SALES_COLS:
        if key == "pending":
            tcells.append(_td(fmt_amount(tp), align, w, bg="#D9D9D9", bold=True))
        else:
            tcells.append(_td("", align, w, bg="#D9D9D9"))
    rows.append("<tr>" + "".join(tcells) + "</tr>")
    return ('<table cellpadding="0" cellspacing="0" '
            'style="border-collapse:collapse;margin:6px 0 14px 0;">'
            + "".join(rows) + "</table>")


def build_sales_html(invoices):
    p = ("font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000;"
         "margin:0 0 10px 0;")
    B = ['<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000;">']
    B.append('<p style="%s">%s</p>' % (p, html.escape(SALES_GREETING)))
    B.append('<p style="%s">%s</p>' % (p, html.escape(SALES_INTRO)))
    B.append(_sales_table(invoices))
    B.append(build_signature_html(SALES_SIGNER))
    B.append("</div>")
    return "".join(B)


def build_sales_plain(invoices):
    out = [SALES_GREETING, "", SALES_INTRO, ""]
    out.append("  Date | Ref. No. | Party's Name | Description | Opening | Pending | Remarks | Overdue")
    tp = 0.0
    for inv in invoices:
        tp += inv["pending"]
        desc = clean_text(inv["particulars"]).split("\n")[0]
        out.append("  {} | {} | {} | {} | {} | {} |  | {}".format(
            inv["date"].strftime("%d-%b-%y"), inv["ref"], inv["party"][:28],
            desc[:40], fmt_amount(inv["opening"]), fmt_amount(inv["pending"]),
            inv["overdue"]))
    out.append("  TOTAL: {}".format(fmt_amount(tp)))
    out.append("")
    out += _plain_sig(SALES_SIGNER)
    return "\n".join(out)


def _plain_sig(signer):
    return ["Thanks & Regards,", "", signer["name"], signer["title"],
            "Direct: " + signer["direct"], "Main: " + signer["main"], "",
            COMPANY_NAME] + COMPANY_ADDR + [COMPANY_CERT]


# ---------------------------------------------------------------------
#  MIME assembly + transports (shared)
# ---------------------------------------------------------------------
def build_message(subject, from_name, from_email, to_list, cc_list, html_body, plain_body):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = formataddr((from_name, from_email))
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Message-ID"] = make_msgid(domain=from_email.split("@")[-1])
    msg.set_content(plain_body)
    msg.add_alternative(html_body, subtype="html")

    html_part = msg.get_payload()[1]
    inline = [(LOGO_FILE, "bcsl_logo")] + \
             [(icon, "soc%d" % i) for i, (icon, _) in enumerate(SOCIAL_LINKS)]
    for fname, cid in inline:
        fpath = get_absolute_path(os.path.join(ASSETS_DIR, fname))
        if not os.path.exists(fpath):
            continue
        ctype, _ = mimetypes.guess_type(fpath)
        maintype, subtype = (ctype or "image/png").split("/", 1)
        with open(fpath, "rb") as f:
            html_part.add_related(f.read(), maintype=maintype, subtype=subtype,
                                  cid="<%s>" % cid, filename=fname)
    return msg


def write_preview(msg, name, subdir):
    folder = get_absolute_path(os.path.join(OUTPUT_DIR, subdir))
    os.makedirs(folder, exist_ok=True)
    base = safe_filename(name)
    eml = os.path.join(folder, base + ".eml")
    with open(eml, "wb") as f:
        f.write(bytes(msg))
    hp = msg.get_body(preferencelist=("html",))
    with open(os.path.join(folder, base + ".html"), "w", encoding="utf-8") as f:
        f.write(hp.get_content() if hp is not None else "")
    return eml


_smtp = {"c": None}


def smtp_connection():
    if _smtp["c"] is not None:
        return _smtp["c"]
    if SMTP_SECURITY == "ssl":
        c = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=60)
    else:
        c = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60); c.ehlo()
        if SMTP_SECURITY == "starttls":
            c.starttls(); c.ehlo()
    if SMTP_USERNAME:
        c.login(SMTP_USERNAME, SMTP_PASSWORD)
    _smtp["c"] = c
    return c


def send_smtp(msg, from_email, to_list, cc_list, bcc_list):
    rcpts = list(dict.fromkeys(to_list + cc_list + bcc_list))
    smtp_connection().send_message(msg, from_addr=from_email, to_addrs=rcpts)


def send_outlook(subject, to_list, cc_list, bcc_list, html_body):
    """Send via the locally installed Outlook desktop app (Windows only).
    The mail is sent from whichever account is signed in to Outlook."""
    import win32com.client  # type: ignore
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.To = "; ".join(to_list)
    mail.CC = "; ".join(cc_list)
    mail.BCC = "; ".join(bcc_list)
    inline = [(LOGO_FILE, "bcsl_logo")] + \
             [(icon, "soc%d" % i) for i, (icon, _) in enumerate(SOCIAL_LINKS)]
    for fname, cid in inline:
        fpath = os.path.abspath(get_absolute_path(os.path.join(ASSETS_DIR, fname)))
        if not os.path.exists(fpath):
            continue
        att = mail.Attachments.Add(fpath, 1)
        att.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)
    mail.HTMLBody = html_body
    mail.Send()


def deliver(subdir, subject, from_name, from_email, entry, html_body, plain_body, name):
    """Dispatch one e-mail through the active transport. Returns (decision, file)."""
    if TRANSPORT == "dry_run":
        msg = build_message(subject, from_name, from_email,
                            entry["to"], entry["cc"], html_body, plain_body)
        return "preview", write_preview(msg, name, subdir)
    if TRANSPORT == "smtp":
        msg = build_message(subject, from_name, from_email,
                            entry["to"], entry["cc"], html_body, plain_body)
        send_smtp(msg, from_email, entry["to"], entry["cc"], entry["bcc"])
        return "sent", ""
    if TRANSPORT == "outlook":
        send_outlook(subject, entry["to"], entry["cc"], entry["bcc"], html_body)
        return "sent", ""
    log.error("Unknown TRANSPORT %r", TRANSPORT); sys.exit(1)


# ---------------------------------------------------------------------
#  JOB RUNNERS
# ---------------------------------------------------------------------
def run_client_job(report):
    log.info("=" * 60)
    log.info("CLIENT job  |  threshold %d days  |  transport %s",
             CLIENT_THRESHOLD_DAYS, TRANSPORT)
    if not os.path.exists(get_absolute_path(CLIENT_ADDRESS_FILE)):
        log.error("Client address file not found: %s", get_absolute_path(CLIENT_ADDRESS_FILE)); return (0, 0, 0)
    book = load_address_book(get_absolute_path(CLIENT_ADDRESS_FILE))
    log.info("Clients with pending bills: %d  |  address book: %d",
             sum(1 for v in report.values() if v), len(book))

    mfile = open(get_absolute_path(os.path.join(OUTPUT_DIR, "manifest_client.csv")), "w",
                 newline="", encoding="utf-8-sig")
    w = csv.writer(mfile)
    w.writerow(["client", "decision", "match", "to", "cc", "bcc",
                "inv_over_30", "inv_upto_30", "total_pending", "file"])

    sent = skipped = errors = 0
    for client_name, invoices in report.items():
        if not invoices:
            continue
        total = sum(i["pending"] for i in invoices)
        over  = [i for i in invoices if i["overdue"] > CLIENT_THRESHOLD_DAYS]
        under = [i for i in invoices if i["overdue"] <= CLIENT_THRESHOLD_DAYS]
        entry, how, skip = resolve_recipient(client_name, book)

        if skip:
            hint = " (closest: '%s')" % how if (skip == "no-address" and how) else ""
            lvl = log.info if skip == "flag-N" else log.warning
            lvl("SKIP  %-45s  %s%s", client_name[:45], skip, hint)
            w.writerow([client_name, "skip:" + skip, "" if skip == "no-address" else how,
                        ";".join(entry["to"]) if entry else "",
                        ";".join(entry["cc"]) if entry else "",
                        ";".join(entry["bcc"]) if entry else "",
                        len(over), len(under), "%.2f" % total, ""])
            skipped += 1
            continue

        if how == "fuzzy":
            log.warning("FUZZY: report '%s' -> book '%s'", client_name, entry["raw"])
        try:
            decision, f = deliver("client", CLIENT_SUBJECT.format(name=client_name),
                                  CLIENT_FROM_NAME, CLIENT_FROM_EMAIL, entry,
                                  build_client_html(client_name, over, under),
                                  build_client_plain(client_name, over, under),
                                  client_name)
            sent += 1
            verb = "DRAFT" if decision == "preview" else "SENT "
            log.info("%s %-45s  client overdue count=%d upto=%d  client pending count=%s", verb, client_name[:45],
                     len(over), len(under), "{:,.2f}".format(total))
            w.writerow([client_name, decision, how, ";".join(entry["to"]),
                         ";".join(entry["cc"]), ";".join(entry["bcc"]),
                         len(over), len(under), "%.2f" % total, f])
            if decision == "sent" and EMAIL_DELAY_SECONDS > 0:
                log.info("Waiting %.1f seconds before next email...", EMAIL_DELAY_SECONDS)
                time.sleep(EMAIL_DELAY_SECONDS)
        except Exception as e:
            errors += 1
            log.exception("ERROR client %s: %s", client_name, e)
            w.writerow([client_name, "error:%s" % e, how, ";".join(entry["to"]),
                        ";".join(entry["cc"]), ";".join(entry["bcc"]),
                        len(over), len(under), "%.2f" % total, ""])
    mfile.close()
    return sent, skipped, errors


def run_sales_job(report):
    log.info("=" * 60)
    log.info("SALES job  |  bills above %d days  |  transport %s",
             SALES_THRESHOLD_DAYS, TRANSPORT)
    if not os.path.exists(get_absolute_path(SALES_ADDRESS_FILE)):
        log.error("Sales address file not found: %s", get_absolute_path(SALES_ADDRESS_FILE)); return (0, 0, 0)
    book = load_address_book(get_absolute_path(SALES_ADDRESS_FILE))

    # regroup every >35-day invoice by sales person (preserve first-seen order)
    buckets = OrderedDict()
    for invoices in report.values():
        for inv in invoices:
            if inv["overdue"] > SALES_THRESHOLD_DAYS:
                buckets.setdefault(inv["sale"], []).append(inv)
    log.info("Sales people with >%d-day bills: %d  |  address book: %d",
             SALES_THRESHOLD_DAYS,
             sum(1 for k in buckets if k.strip()), len(book))

    mfile = open(get_absolute_path(os.path.join(OUTPUT_DIR, "manifest_sales.csv")), "w",
                 newline="", encoding="utf-8-sig")
    w = csv.writer(mfile)
    w.writerow(["sale_person", "decision", "match", "to", "cc", "bcc",
                "bills_over_35", "total_pending", "file"])

    sent = skipped = errors = 0
    for sp, invs in buckets.items():
        invs = sorted(invs, key=lambda i: i["date"])
        total = sum(i["pending"] for i in invs)

        if not sp.strip():                                  # unassigned bills
            log.warning("SKIP  (No Sale Person)  %d aged bills with no sales person "
                        "(total %s) - need manual handling", len(invs),
                        "{:,.2f}".format(total))
            w.writerow(["(No Sale Person)", "skip:no-saleperson", "", "", "", "",
                        len(invs), "%.2f" % total, ""])
            skipped += 1
            continue

        entry, how, skip = resolve_recipient(sp, book)
        if skip:
            hint = " (closest: '%s')" % how if (skip == "no-address" and how) else ""
            lvl = log.info if skip == "flag-N" else log.warning
            lvl("SKIP  %-25s  %s%s", sp, skip, hint)
            w.writerow([sp, "skip:" + skip, "" if skip == "no-address" else how,
                        ";".join(entry["to"]) if entry else "",
                        ";".join(entry["cc"]) if entry else "",
                        ";".join(entry["bcc"]) if entry else "",
                        len(invs), "%.2f" % total, ""])
            skipped += 1
            continue

        if how == "fuzzy":
            log.warning("FUZZY: report '%s' -> book '%s'", sp, entry["raw"])
        try:
            decision, f = deliver("sales", SALES_SUBJECT, SALES_FROM_NAME,
                                  SALES_FROM_EMAIL, entry, build_sales_html(invs),
                                  build_sales_plain(invs), sp)
            sent += 1
            verb = "DRAFT" if decision == "preview" else "SENT "
            log.info("%s %-25s  bills=%d  pend=%s", verb, sp, len(invs),
                     "{:,.2f}".format(total))
            w.writerow([sp, decision, how, ";".join(entry["to"]),
                         ";".join(entry["cc"]), ";".join(entry["bcc"]),
                         len(invs), "%.2f" % total, f])
            if decision == "sent" and EMAIL_DELAY_SECONDS > 0:
                log.info("Waiting %.1f seconds before next email...", EMAIL_DELAY_SECONDS)
                time.sleep(EMAIL_DELAY_SECONDS)
        except Exception as e:
            errors += 1
            log.exception("ERROR sales %s: %s", sp, e)
            w.writerow([sp, "error:%s" % e, how, ";".join(entry["to"]),
                        ";".join(entry["cc"]), ";".join(entry["bcc"]),
                        len(invs), "%.2f" % total, ""])
    mfile.close()
    return sent, skipped, errors


# ---------------------------------------------------------------------
#  main
# ---------------------------------------------------------------------
def main():
    global TRANSPORT
    load_config()
    setup_logging()

    ap = argparse.ArgumentParser(
        description="Send BCSL outstanding e-mails to clients and/or the sales team.")
    ap.add_argument("jobs", nargs="?", choices=["client", "sales", "both"],
                    default=RUN_JOBS, help="which job(s) to run (default: %(default)s)")
    ap.add_argument("--transport", choices=["dry_run", "smtp", "outlook"],
                    default=None, help="override the delivery mode for this run")
    args = ap.parse_args()
    if args.transport:
        TRANSPORT = args.transport

    if not os.path.exists(get_absolute_path(REPORT_FILE)):
        log.error("Report file not found: %s", get_absolute_path(REPORT_FILE)); sys.exit(1)

    log.info("Run jobs = %s  |  transport = %s", args.jobs, TRANSPORT)
    report = parse_report(get_absolute_path(REPORT_FILE))

    totals = {}
    if args.jobs in ("client", "both"):
        totals["client"] = run_client_job(report)
    if args.jobs in ("sales", "both"):
        totals["sales"] = run_sales_job(report)

    if _smtp["c"] is not None:
        try:
            _smtp["c"].quit()
        except Exception:
            pass

    verb = "previewed" if TRANSPORT == "dry_run" else "sent"
    log.info("=" * 60)
    for job, (s, k, e) in totals.items():
        log.info("%-7s : %d %s, %d skipped, %d errors", job, s, verb, k, e)
    if TRANSPORT == "dry_run":
        log.info("Previews under '%s/<job>/'. Review them, then re-run with "
                 "--transport smtp (or outlook).", OUTPUT_DIR)


if __name__ == "__main__":
    main()
