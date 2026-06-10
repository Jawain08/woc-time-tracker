import streamlit as st
import pandas as pd
import io
import datetime
import zoneinfo
import os
import requests
import smtplib
import time
from email.mime.text import MIMEText
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# =============================================================================
# SECTION 1: SYSTEM CONFIGURATION & CONSTANTS
# =============================================================================
st.set_page_config(page_title="WOC - Time Tracking System", layout="wide", page_icon="📝")

# Endpoints for writing data
REGISTRATION_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSdY6ydD4YLYQEicFkk21DIRefUTT5ht8v4lbdZVr6hSbGOBAA/formResponse"
LOG_ENTRY_FORM_URL    = "https://docs.google.com/forms/d/1G8flLQrWJWGl5CwOEUe48zuAPre5mhJrbanx33uSkZk/formResponse"

# Endpoints for reading data
SHEET_ID       = "1zop4YKXKA1H8Iv89YwkGpP4c4YlGGFgz5jDYLT3psik"
TIMESHEETS_GID = "742432797"
ACCOUNTS_GID   = "1781560298"

# =============================================================================
# SECTION 2: CUSTOM CSS — color-coded rows, metric bar, mobile nav
# =============================================================================

st.markdown("""
<style>
/* ── Color-coded activity badge pills ── */
.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 12px;
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 0.5px;
}
.badge-NOFA   { background:#DBEAFE; color:#1D4ED8; }
.badge-CARP   { background:#DCFCE7; color:#15803D; }
.badge-WOC    { background:#FEF3C7; color:#B45309; }
.badge-JJ     { background:#F3E8FF; color:#7E22CE; }
.badge-TRICAP { background:#FFE4E6; color:#BE123C; }
.badge-OTHER  { background:#F1F5F9; color:#475569; }

/* ── Persistent hours bar ── */
.hours-bar {
    background: linear-gradient(135deg, #7B2CBF 0%, #9D4EDD 100%);
    border-radius: 12px;
    padding: 14px 24px;
    margin-bottom: 18px;
    display: flex;
    align-items: center;
    gap: 32px;
    flex-wrap: wrap;
}
.hours-bar .metric { text-align: center; }
.hours-bar .metric .val {
    font-size: 26px;
    font-weight: 800;
    color: #ffffff;
    line-height: 1;
}
.hours-bar .metric .lbl {
    font-size: 11px;
    color: #E0AAFF;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-top: 2px;
}
.hours-bar .divider {
    width: 1px; height: 40px;
    background: rgba(255,255,255,0.25);
}
.hours-bar .period-label {
    font-size: 13px;
    color: #E0AAFF;
    margin-left: auto;
}

/* ── Mobile-friendly nav ── */
@media (max-width: 640px) {
    .nav-btn { font-size: 12px !important; padding: 6px 4px !important; }
}

/* ── Admin panel highlight ── */
.admin-badge {
    display: inline-block;
    background: #7B2CBF;
    color: white;
    font-size: 11px;
    font-weight: 700;
    padding: 2px 10px;
    border-radius: 20px;
    letter-spacing: 1px;
    margin-left: 8px;
    vertical-align: middle;
}

/* ── Duplicate warning ── */
.dup-warning {
    background: #FEF9C3;
    border-left: 4px solid #EAB308;
    padding: 10px 16px;
    border-radius: 6px;
    font-size: 14px;
    margin-bottom: 8px;
}
</style>
""", unsafe_allow_html=True)


# =============================================================================
# SECTION 3: UI HEADERS & BADGE HELPERS
# =============================================================================

def render_woc_header():
    logo_filename = "woclogo.png"
    if os.path.exists(logo_filename):
        st.image(logo_filename, width=280)
    st.markdown(
        """
        <div style="background-color: #7B2CBF; padding: 15px 20px; border-radius: 10px;
                    margin-top: 10px; margin-bottom: 25px;">
            <h1 style="color: white; margin: 0; font-family: 'Calibri', sans-serif;
                       font-size: 30px; font-weight: bold; letter-spacing: 0.5px;">
                Women of Colors, Inc.</h1>
            <p style="color: #E0AAFF; margin: 4px 0 0 0; font-size: 16px;
                      font-family: 'Calibri', sans-serif;">
                Saginaw Community Prevention &amp; Training Program Hub</p>
        </div>
        """,
        unsafe_allow_html=True
    )

CODE_COLORS = {
    "NOFA":   "NOFA",
    "CARP":   "CARP",
    "WOC":    "WOC",
    "JJ":     "JJ",
    "TRICAP": "TRICAP",
}

def code_badge(code):
    cls = CODE_COLORS.get(str(code).strip().upper(), "OTHER")
    return f'<span class="badge badge-{cls}">{code}</span>'

def hours_to_hhmm(decimal_hours):
    """Convert 1.75 → '1:45' safely with rounding to prevent 1:60 bugs"""
    try:
        total_minutes = int(round(float(decimal_hours) * 60))
        h = total_minutes // 60
        m = total_minutes % 60
        return f"{h}:{m:02d}"
    except Exception:
        return str(decimal_hours)


# =============================================================================
# SECTION 4: DATA FETCHING & MAPPING HELPERS
# =============================================================================

def map_columns(df, rules):
    cols_map = {}
    for col in df.columns:
        c_lower = str(col).lower().strip()
        if ".1" in c_lower or ".2" in c_lower:
            continue
        for standard_name, keywords in rules.items():
            if standard_name not in cols_map:
                if any(kw in c_lower for kw in keywords):
                    cols_map[standard_name] = col
                    break
    if len(cols_map) >= max(1, len(rules) - 2):
        result = pd.DataFrame()
        for standard_name, actual_col in cols_map.items():
            result[standard_name] = df[actual_col]
        return result, True
    return df.copy(), False

def _csv_url(gid):
    return f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}"

def _fetch_csv_with_retry(url, max_attempts=3):
    """Fetch a CSV URL with exponential backoff on failure."""
    for attempt in range(max_attempts):
        try:
            return pd.read_csv(url)
        except Exception as e:
            if attempt == max_attempts - 1:
                raise e
            time.sleep(2 ** attempt)

@st.cache_data(ttl=60)
def fetch_timesheets():
    try:
        raw = _fetch_csv_with_retry(_csv_url(TIMESHEETS_GID))
        rules = {
            "Timestamp":       ["timestamp"],
            "Date":            ["date"],
            "Instructor Name": ["instructor", "name", "staff"],
            "Time In":         ["time in"],
            "Time Out":        ["time out"],
            "Activity":        ["activity"],
            "Code":            ["code"],
            "Category":        ["category"],
            "Description":     ["description"],
            "Minutes":         ["minutes"],
            "Hours":           ["hours"],
        }
        df, mapped = map_columns(raw, rules)

        if not mapped:
            if len(raw.columns) >= 11:
                df = raw.copy()
                df.columns = (
                    ["Timestamp", "Date", "Instructor Name", "Time In", "Time Out",
                     "Activity", "Code", "Category", "Description", "Minutes", "Hours"]
                    + list(raw.columns[11:])
                )
                st.warning("⚠️ Timesheet columns could not be auto-detected — using positional "
                           "fallback. Check Google Sheet headers if data looks wrong.")
            else:
                st.error("🛑 Timesheet sheet structure is unrecognised. Contact your admin.")
                return pd.DataFrame(columns=["Timestamp","Date","Instructor Name","Time In",
                                             "Time Out","Activity","Code","Category",
                                             "Description","Minutes","Hours"])

        if "Code" in df.columns:
            df["Code"] = (df["Code"].astype(str).str.strip()
                            .replace({"MPHI": "CARP", "mphi": "CARP"}))
        return df

    except Exception as e:
        st.error(f"🛑 Timesheet Database Connection Error (after retries): {e}")
        return pd.DataFrame(columns=["Timestamp","Date","Instructor Name","Time In",
                                     "Time Out","Activity","Code","Category",
                                     "Description","Minutes","Hours"])

@st.cache_data(ttl=60)
def fetch_accounts():
    try:
        raw = _fetch_csv_with_retry(_csv_url(ACCOUNTS_GID))
        rules = {
            "Timestamp":       ["timestamp"],
            "Instructor Name": ["instructor", "name"],
            "Email Address":   ["email"],
            "PIN":             ["pin"],
        }
        df, mapped = map_columns(raw, rules)

        if not mapped:
            if len(raw.columns) >= 4:
                df = raw.copy()
                df.columns = (["Timestamp", "Instructor Name", "Email Address", "PIN"]
                               + list(raw.columns[4:]))
                st.warning("⚠️ Accounts columns could not be auto-detected — using positional "
                           "fallback. Check Sheet headers if logins fail.")
            else:
                st.error("🛑 Accounts sheet structure is unrecognised. Contact your admin.")
                return pd.DataFrame(columns=["Timestamp","Instructor Name","Email Address","PIN"])
        return df

    except Exception as e:
        st.error(f"🛑 Account Profile Connection Error (after retries): {e}")
        return pd.DataFrame(columns=["Timestamp","Instructor Name","Email Address","PIN"])


# =============================================================================
# SECTION 5: EMAIL HELPER
# =============================================================================

def send_pin_email(recipient_email, recipient_name, user_pin):
    if "smtp" in st.secrets:
        try:
            msg = MIMEText(
                f"Hello {recipient_name},\n\n"
                f"Your PIN for the WOC Time Tracking Hub is: {user_pin}\n\n"
                f"Log in here: https://share.streamlit.io/\n\n"
                f"Regards,\nWomen of Colors Payroll Admin"
            )
            msg['Subject'] = "WOC Time Tracker - PIN Recovery"
            msg['From']    = st.secrets["smtp"]["username"]
            msg['To']      = recipient_email
            with smtplib.SMTP_SSL(st.secrets["smtp"]["server"], int(st.secrets["smtp"]["port"])) as server:
                server.login(st.secrets["smtp"]["username"], st.secrets["smtp"]["password"])
                server.sendmail(st.secrets["smtp"]["username"], [recipient_email], msg.as_string())
            return True, "Success"
        except Exception as e:
            return False, str(e)
    return False, "Fallback"


# =============================================================================
# SECTION 6: DUPLICATE DETECTION HELPER
# =============================================================================

def check_duplicate(entry_date, time_in, time_out, name, df):
    """Returns True if an identical entry already exists in the cached data."""
    if df.empty or "ParsedDate" not in df.columns:
        return False
    match = df[
        (df["Instructor Name"].astype(str).str.strip().str.lower() == name.strip().lower()) &
        (df["ParsedDate"] == entry_date) &
        (df["Time In"].astype(str).str.strip()  == time_in.strip()) &
        (df["Time Out"].astype(str).str.strip() == time_out.strip())
    ]
    return not match.empty


# =============================================================================
# SECTION 7: EXCEL GENERATION HELPERS (Hoisted for Performance)
# =============================================================================

def _make_styles():
    """Shared Excel Style Factory"""
    thin = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC'),
    )
    shaded = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    return {
        "title":   Font(name="Calibri", size=14, bold=True),
        "bold":    Font(name="Calibri", size=11, bold=True),
        "regular": Font(name="Calibri", size=11),
        "small":   Font(name="Calibri", size=9, italic=True),
        "cursive": Font(name="Brush Script MT", size=16, italic=True, color="002060"),
        "thin":    thin,
        "shaded":  shaded,
    }

@st.cache_data(ttl=60)
def build_timesheet_bytes(instr, p_start, p_end, period_data_json, today_str):
    """Cached Excel generation — only rebuilds when data or period changes."""
    period_data = pd.read_json(io.StringIO(period_data_json), orient="records")
    wb = Workbook()
    ws = wb.active
    ws.title = "Time Sheet"
    ws.sheet_view.showGridLines  = True
    ws.print_options.gridLines   = True
    ws.page_setup.orientation    = 'landscape'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 1

    S = _make_styles()

    ws["A1"] = "BiWeekly Employee Time Sheet";  ws["A1"].font = S["title"]
    ws["A2"] = "Women of Colors";               ws["A2"].font = S["bold"]
    ws["A3"] = "Employee Details:";             ws["A3"].font = S["bold"]
    ws["D3"] = f"Name :  {instr}";              ws["D3"].font = S["regular"]
    ws["F3"] = "Email: payroll@yeoandyeo.com";  ws["F3"].font = S["regular"]
    ws["A4"] = "Manager Details:";              ws["A4"].font = S["bold"]
    ws["D4"] = "Name: Vicki Hill";              ws["D4"].font = S["regular"]
    ws["F4"] = "Fax: 989-793-0186";             ws["F4"].font = S["regular"]
    ws["A5"] = f"Period Start Date: {p_start.strftime('%m/%d/%Y')}"; ws["A5"].font = S["bold"]
    ws["E5"] = f"Period End Date:  {p_end.strftime('%m/%d/%Y')}";    ws["E5"].font = S["bold"]

    for col_idx, text in enumerate(["","","","","","Total Hours","NOFA","WOC","JJ","TRICAP","CARP"], 1):
        if text:
            c = ws.cell(row=7, column=col_idx, value=text)
            c.font = S["bold"]
            c.alignment = Alignment(horizontal="center", wrap_text=True)

    for col_idx, text in enumerate(["","Day","Date","Time In","Time Out","Hours Worked","NOFA","WOC","JJ","TRICAP","CARP"], 1):
        if text:
            c = ws.cell(row=9, column=col_idx, value=text)
            c.font = S["bold"]
            c.alignment = Alignment(horizontal="center")

    code_col_map   = {"NOFA": 7, "WOC": 8, "JJ": 9, "TRICAP": 10, "CARP": 11}
    days_in_period = max(1, (p_end - p_start).days + 1)
    date_list      = [p_start + datetime.timedelta(days=x) for x in range(days_in_period)]

    if "Date" in period_data.columns:
        period_data["ParsedDate"] = pd.to_datetime(period_data["Date"], errors='coerce').dt.date

    r = 10
    for idx, d in enumerate(date_list):
        if idx == 7:
            r += 1
        ws.cell(row=r, column=2, value=d.strftime("%A")).font       = S["regular"]
        ws.cell(row=r, column=3, value=d.strftime("%Y-%m-%d")).font = S["regular"]

        day_logs = period_data[period_data['ParsedDate'] == d] if "ParsedDate" in period_data.columns else pd.DataFrame()

        if not day_logs.empty:
            ws.cell(row=r, column=4, value=" / ".join(day_logs['Time In'].astype(str).tolist())).font  = S["regular"]
            ws.cell(row=r, column=5, value=" / ".join(day_logs['Time Out'].astype(str).tolist())).font = S["regular"]
            ws.cell(row=r, column=6, value=day_logs['Hours'].astype(float).sum()).font                 = S["regular"]
            
            for c_idx in range(7, 12):
                c = ws.cell(row=r, column=c_idx, value=0)
                c.font = S["regular"]; c.fill = S["shaded"]; c.border = S["thin"]
            for _, rl in day_logs.iterrows():
                code = str(rl.get('Code', ''))
                hw   = float(rl.get('Hours', 0.0))
                if code in code_col_map:
                    ci  = code_col_map[code]
                    cv  = ws.cell(row=r, column=ci).value or 0.0
                    ac  = ws.cell(row=r, column=ci, value=cv + hw)
                    ac.fill = PatternFill(fill_type=None)
        else:
            for c_idx in range(4, 12):
                c = ws.cell(row=r, column=c_idx, value=0)
                c.font = S["regular"]; c.fill = S["shaded"]; c.border = S["thin"]
        r += 1

    total_hrs = period_data['Hours'].astype(float).sum() if not period_data.empty else 0.0
    r += 1
    ws.cell(row=r, column=2, value="Total Target Hours:").font  = S["bold"]
    ws.cell(row=r, column=3, value=75.0).font                   = S["bold"]
    r += 1
    ws.cell(row=r, column=2, value="Actual Hours Worked:").font = S["bold"]
    ws.cell(row=r, column=3, value=total_hrs).font              = S["bold"]
    ws.cell(row=r, column=6, value=total_hrs).font              = S["bold"]

    r += 2
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=11)
    cert = ws.cell(row=r, column=2, value="CLIENT: I CERTIFY THAT THE HOURS WORKED ON THIS TIME SLIP ARE CORRECT.")
    cert.font = S["bold"]
    cert.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[r].height = 24

    r += 2
    ws.cell(row=r, column=2, value="Employee Signature:").font = S["bold"]
    ws.cell(row=r, column=3, value=instr).font                 = S["cursive"]
    ws.cell(row=r, column=5, value="Date:").font               = S["bold"]
    ws.cell(row=r, column=6, value=today_str).font             = S["regular"]
    r += 2
    ws.cell(row=r, column=2, value="Manager Signature:").font  = S["bold"]
    ws.cell(row=r, column=5, value="Date:").font               = S["bold"]

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if cell.row in [1,2,3,4,5] or "CLIENT: I CERTIFY" in val_str:
                continue
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

@st.cache_data(ttl=60)
def build_additional_bytes(instr, p_start, p_end, period_data_json):
    """Cached additional hours report generation."""
    period_data = pd.read_json(io.StringIO(period_data_json), orient="records")
    wb = Workbook()
    ws = wb.active
    ws.title = "Report Form"
    ws.sheet_view.showGridLines  = True
    ws.print_options.gridLines   = True
    ws.page_setup.orientation    = 'landscape'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 1

    S = _make_styles()

    ws["A1"] = f"Additional Hours Report FY20 - Due {p_start.strftime('%m/%d/%y')} - {p_end.strftime('%m/%d/%Y')}"
    ws["A1"].font = S["bold"]
    ws["A2"] = "Agency Name";              ws["A2"].font = S["bold"]
    ws["C2"] = "ADDITIONAL HOURS REPORT";  ws["C2"].font = S["bold"]

    for col_idx, h in enumerate(["Date","Staff Name","Category","Description","Time in minutes"], 1):
        c = ws.cell(row=3, column=col_idx, value=h)
        c.font = S["bold"]
        c.alignment = Alignment(horizontal="left")

    curr_row   = 4
    total_mins = 0
    for _, log in period_data.iterrows():
        mins = int(log.get('Minutes', 0))
        ws.cell(row=curr_row, column=1, value=str(log.get('Date',''))).font        = S["regular"]
        ws.cell(row=curr_row, column=2, value=instr).font                          = S["regular"]
        ws.cell(row=curr_row, column=3, value=str(log.get('Category',''))).font    = S["regular"]
        ws.cell(row=curr_row, column=4, value=str(log.get('Description',''))).font = S["regular"]
        ws.cell(row=curr_row, column=5, value=mins).font                           = S["regular"]
        total_mins += mins
        curr_row   += 1

    ws.cell(row=curr_row, column=4, value="Total").font    = S["bold"]
    ws.cell(row=curr_row, column=5, value=total_mins).font = S["bold"]

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# =============================================================================
# SECTION 8: LOGIN / REGISTRATION PORTAL
# =============================================================================

existing_data    = fetch_timesheets()
account_registry = fetch_accounts()

if "user" in st.query_params and "logged_in" not in st.session_state:
    st.session_state["logged_in"]       = True
    st.session_state["instructor_name"] = st.query_params["user"]
    st.session_state["is_admin"]        = False

if not st.session_state.get("logged_in"):
    render_woc_header()
    st.markdown("<h3 style='color: #7B2CBF; margin-bottom: 10px;'>🔐 Instructor Access Hub</h3>",
                unsafe_allow_html=True)
    portal_tab = st.radio(
        "Choose Action:",
        ["Sign In", "Create Custom Account / PIN", "Forgot PIN / Reset Option"],
        horizontal=True, label_visibility="collapsed"
    )

    col_portal, _ = st.columns([1.5, 2])
    with col_portal:

        # ── SIGN IN ──────────────────────────────────────────────────────────
        if portal_tab == "Sign In":
            with st.form("signin_panel"):
                login_name   = st.text_input("Instructor Name:", placeholder="First & Last Name")
                login_pin    = st.text_input("Enter Personal PIN:", type="password",
                                             placeholder="Type your PIN")
                submit_login = st.form_submit_button("🔓 Log In")

                if submit_login:
                    cleaned_name  = login_name.strip()
                    matched_users = account_registry[
                        account_registry["Instructor Name"]
                            .astype(str).str.strip().str.lower() == cleaned_name.lower()
                    ]
                    if not matched_users.empty:
                        correct_pin = str(matched_users.iloc[-1]["PIN"]).strip()
                        if login_pin.strip() == correct_pin:
                            st.session_state["logged_in"]       = True
                            st.session_state["instructor_name"] = cleaned_name
                            st.session_state["is_admin"]        = False
                            st.query_params["user"] = cleaned_name
                            st.rerun()
                        else:
                            st.error("Authentication Error: Invalid PIN for this profile.")
                    else:
                        master_pw = st.secrets.get("admin", {}).get("master_password", "")
                        if master_pw and login_pin.strip() == master_pw:
                            st.session_state["logged_in"]       = True
                            st.session_state["instructor_name"] = cleaned_name
                            st.session_state["is_admin"]        = True
                            st.query_params["user"] = cleaned_name
                            st.rerun()
                        else:
                            st.error(f"Profile Not Found: '{cleaned_name}' is not registered yet.")

        # ── REGISTRATION ─────────────────────────────────────────────────────
        elif portal_tab == "Create Custom Account / PIN":
            st.info("💡 Setting up your profile connects your name to a custom passcode.")
            with st.form("registration_panel"):
                reg_name  = st.text_input("Full Instructor Name:", placeholder="First and Last Name")
                reg_email = st.text_input("Email Address:",        placeholder="username@domain.com")
                reg_pin   = st.text_input("Create 4-to-6 Digit PIN:", type="password",
                                          placeholder="Choose your passcode")
                submit_reg = st.form_submit_button("📝 Register Account")

                if submit_reg:
                    if not reg_name.strip() or not reg_email.strip() or not reg_pin.strip():
                        st.error("Validation Error: All fields are required.")
                    else:
                        acc_data = {
                            "entry.576544689":  reg_name.strip(),
                            "entry.836662014":  reg_email.strip(),
                            "entry.2099667226": reg_pin.strip()
                        }
                        with st.spinner("Creating your secure profile..."):
                            try:
                                res = requests.post(REGISTRATION_FORM_URL, data=acc_data, timeout=10)
                                if res.ok:
                                    st.success("Account created! Switch to 'Sign In' to enter.")
                                else:
                                    st.error(f"Database Error (Code {res.status_code}).")
                            except Exception as e:
                                st.error(f"Network Connection Failed: {e}")

        # ── PIN RECOVERY ──────────────────────────────────────────────────────
        elif portal_tab == "Forgot PIN / Reset Option":
            with st.form("recovery_panel"):
                recover_email   = st.text_input("Registered Email:", placeholder="username@domain.com")
                submit_recovery = st.form_submit_button("🔍 Retrieve Access Passcode")

                if submit_recovery:
                    matched_emails = account_registry[
                        account_registry["Email Address"].astype(str).str.strip().str.lower()
                        == recover_email.strip().lower()
                    ]
                    if not matched_emails.empty:
                        user_account = matched_emails.iloc[-1]
                        found_name   = user_account["Instructor Name"]
                        found_pin    = user_account["PIN"]
                        with st.spinner("Dispatching secure recovery instructions..."):
                            status, _ = send_pin_email(recover_email.strip(), found_name, found_pin)
                        if status:
                            st.success(f"📬 Recovery instructions sent to {recover_email.strip()}.")
                        else:
                            st.warning("⚙️ Email offline. Showing PIN below:")
                            st.info(f"Account: **{found_name}** | PIN: `{found_pin}`")
                    else:
                        st.error("That email is not in our registry.")
    st.stop()


# =============================================================================
# SECTION 9: TIMEZONE-AWARE TODAY  (America/Detroit = Saginaw, MI)
# =============================================================================

TODAY = datetime.datetime.now(zoneinfo.ZoneInfo("America/Detroit")).date()

# =============================================================================
# SECTION 10: AUTHENTICATED SESSION HEADER
# =============================================================================

render_woc_header()
instructor_input = st.session_state["instructor_name"]
is_admin         = st.session_state.get("is_admin", False)

# Trigger Success Toast after a page refresh
if "toast_message" in st.session_state:
    st.toast(st.session_state["toast_message"], icon="🎉")
    del st.session_state["toast_message"]

col_user1, col_user2 = st.columns([3, 1])
with col_user1:
    admin_badge = '<span class="admin-badge">ADMIN</span>' if is_admin else ""
    st.markdown(f"#### Welcome back, **{instructor_input}**! 👋 {admin_badge}", unsafe_allow_html=True)
with col_user2:
    if st.button("🚪 Log Out / Clear Session", use_container_width=True):
        st.session_state.clear()
        st.query_params.clear()
        st.rerun()


# =============================================================================
# SECTION 11: PAY PERIOD NAVIGATION
# =============================================================================

if "period_offset" not in st.session_state:
    st.session_state.period_offset = 0

ANCHOR_DATE       = datetime.date(2026, 5, 24)
days_since_anchor = (TODAY - ANCHOR_DATE).days
completed_periods = days_since_anchor // 14

active_period_index = completed_periods + st.session_state.period_offset
auto_period_start   = ANCHOR_DATE + datetime.timedelta(days=active_period_index * 14)
auto_period_end     = auto_period_start + datetime.timedelta(days=13)

st.subheader("🗓️ Pay Period Review Settings")

col_nav1, col_nav2, col_nav3 = st.columns([1, 2, 1])
with col_nav1:
    if st.button("⬅️ Previous", use_container_width=True, key="nav_prev"):
        st.session_state.period_offset -= 1
        st.rerun()
with col_nav2:
    col_date, col_reset = st.columns([2, 1])
    with col_date:
        st.markdown(
            f"<h4 style='text-align:center;margin-top:5px;color:#7B2CBF;'>"
            f"{auto_period_start.strftime('%b %d')} — {auto_period_end.strftime('%b %d, %Y')}"
            f"</h4>", unsafe_allow_html=True
        )
    with col_reset:
        if st.session_state.period_offset != 0:
            if st.button("🔄 Current", use_container_width=True, help="Jump back to the active pay period"):
                st.session_state.period_offset = 0
                st.rerun()
with col_nav3:
    if st.button("Next ➡️", use_container_width=True, key="nav_next"):
        st.session_state.period_offset += 1
        st.rerun()

col_p1, col_p2 = st.columns(2)
with col_p1:
    pay_period_start = st.date_input("Start Date (editable override)", value=auto_period_start)
with col_p2:
    pay_period_end   = st.date_input("End Date (editable override)",   value=auto_period_end)

if pay_period_start != auto_period_start or pay_period_end != auto_period_end:
    st.info("ℹ️ Manual date override active. Use nav buttons to return to automatic mode.")

st.markdown("---")


# =============================================================================
# SECTION 12: FILTER DATA
# =============================================================================

total_database_records = 0
current_period_df      = pd.DataFrame()
running_hours          = 0.0
running_minutes        = 0
all_instructors_df     = pd.DataFrame()

if not existing_data.empty and "Instructor Name" in existing_data.columns:
    all_instructors_df = existing_data.copy()
    if "Date" in all_instructors_df.columns:
        all_instructors_df["ParsedDate"] = pd.to_datetime(all_instructors_df["Date"], errors='coerce').dt.date
        all_instructors_df = all_instructors_df.dropna(subset=["ParsedDate"])

    if instructor_input.strip():
        user_filtered_df = all_instructors_df[
            all_instructors_df["Instructor Name"].astype(str).str.strip().str.lower() == instructor_input.strip().lower()
        ].copy()

        if not user_filtered_df.empty:
            current_period_df = user_filtered_df[
                (user_filtered_df["ParsedDate"] >= pay_period_start) &
                (user_filtered_df["ParsedDate"] <= pay_period_end)
            ].sort_values(by="ParsedDate", ascending=True)

            total_database_records = len(user_filtered_df)
            running_hours   = current_period_df['Hours'].astype(float).sum() if 'Hours' in current_period_df.columns else 0.0
            running_minutes = current_period_df['Minutes'].astype(int).sum() if 'Minutes' in current_period_df.columns else 0


# =============================================================================
# SECTION 13: PERSISTENT HOURS BAR
# =============================================================================

hhmm = hours_to_hhmm(running_hours)
period_label = f"{pay_period_start.strftime('%b %d')} – {pay_period_end.strftime('%b %d, %Y')}"

st.markdown(f"""
<div class="hours-bar">
    <div class="metric">
        <div class="val">{running_hours:.2f}</div>
        <div class="lbl">Hours This Period</div>
    </div>
    <div class="divider"></div>
    <div class="metric">
        <div class="val">{hhmm}</div>
        <div class="lbl">HH:MM Format</div>
    </div>
    <div class="divider"></div>
    <div class="metric">
        <div class="val">{running_minutes}</div>
        <div class="lbl">Total Minutes</div>
    </div>
    <div class="divider"></div>
    <div class="metric">
        <div class="val">{len(current_period_df)}</div>
        <div class="lbl">Entries Logged</div>
    </div>
    <div class="period-label">📅 {period_label}</div>
</div>
""", unsafe_allow_html=True)


# =============================================================================
# SECTION 14: ACTIVITY DICTIONARY & TIME SLOTS
# =============================================================================

activity_to_code_mapping = {
    "PFL instructor Training (Juvenile)": {"code": "JJ",     "category": "Other", "description": "PFL Instructor Training - Juvenile"},
    "PFL instructor Training (Tri-Cap)":  {"code": "TRICAP", "category": "Other", "description": "PFL Instructor Training - Tri-Cap"},
    "PFL (Training/Data Entry)":          {"code": "NOFA",   "category": "Other", "description": "PFL Training / Data Entry"},
    "Botvin Life Skills Training":        {"code": "NOFA",   "category": "Other", "description": "Botvin Life Skills Training"},
    "Prevention Team Meeting":            {"code": "NOFA",   "category": "Other", "description": "Prevention Team Meeting"},
    "WOC Facility Maintenance":           {"code": "WOC",    "category": "Other", "description": "WOC Facility Maintenance"},
    "WOC IT Support":                     {"code": "WOC",    "category": "Other", "description": "WOC IT Support"},
    "Sick Day":                           {"code": "NOFA",   "category": "Other", "description": "Sick Day"},
    "CARP":                               {"code": "CARP",   "category": "Other", "description": "CARP"},
    "Pathway To Purpose":                 {"code": "JJ",     "category": "Other", "description": "Pathway To Purpose"},
    "Office Admin NOFA":                  {"code": "NOFA",   "category": "Other", "description": "Office Admin NOFA"},
    "Office Admin CARP":                  {"code": "CARP",   "category": "Other", "description": "Office Admin CARP"},
}
all_activities = list(activity_to_code_mapping.keys())

def generate_time_slots():
    slots = []
    for period in ["AM", "PM"]:
        for hour in range(1, 13):
            for minute in ["00", "15", "30", "45"]:
                slots.append(f"{hour:02d}:{minute} {period}")
    return slots

time_dropdown_options = generate_time_slots()


# =============================================================================
# SECTION 15: DAILY LOG ENTRY FORM 
# =============================================================================

st.subheader("⏳ Log Daily Activity")

with st.form("daily_time_entry_form", clear_on_submit=True):
    entry_col1, entry_col2, entry_col3, entry_col4 = st.columns(4)
    with entry_col1:
        default_date = TODAY if auto_period_start <= TODAY <= auto_period_end else auto_period_start
        entry_date   = st.date_input(
            "Date Worked", value=default_date,
            min_value=pay_period_start - datetime.timedelta(days=365),
            max_value=pay_period_end   + datetime.timedelta(days=365)
        )
    with entry_col2:
        time_in_str  = st.selectbox("Time In",  options=time_dropdown_options, index=67)
    with entry_col3:
        time_out_str = st.selectbox("Time Out", options=time_dropdown_options, index=74)
    with entry_col4:
        activity_selected = st.selectbox("Activity Classification", all_activities)

    add_btn = st.form_submit_button("➕ Save Entry to Log")

if add_btn:
    start_time_dt = datetime.datetime.strptime(f"{entry_date} {time_in_str}",  "%Y-%m-%d %I:%M %p")
    end_time_dt   = datetime.datetime.strptime(f"{entry_date} {time_out_str}", "%Y-%m-%d %I:%M %p")

    if end_time_dt <= start_time_dt:
        st.error("Validation Error: 'Time Out' must occur after 'Time In'.")
    else:
        is_dup = check_duplicate(entry_date, time_in_str, time_out_str, instructor_input, all_instructors_df)
        if is_dup:
            st.markdown(
                '<div class="dup-warning">⚠️ <strong>Duplicate Detected:</strong> An entry with '
                'the same date, time in, and time out already exists in your log. '
                'Submission blocked to prevent double-counting.</div>',
                unsafe_allow_html=True
            )
        else:
            duration_minutes = int((end_time_dt - start_time_dt).total_seconds() / 60)
            duration_hours   = round(duration_minutes / 60, 2)
            mapping_result   = activity_to_code_mapping[activity_selected]

            form_data = {
                "entry.1205527392": entry_date.strftime("%Y-%m-%d"),
                "entry.1822017875": instructor_input.strip(),
                "entry.1148008178": time_in_str,
                "entry.1036423098": time_out_str,
                "entry.1565734482": activity_selected,
                "entry.1863736208": mapping_result['code'],
                "entry.835834590":  mapping_result['category'],
                "entry.693720626":  mapping_result['description'],
                "entry.2039394575": duration_minutes,
                "entry.1380701779": duration_hours,
            }
            
            with st.spinner("Encrypting and saving your entry to the cloud..."):
                try:
                    response = requests.post(LOG_ENTRY_FORM_URL, data=form_data, timeout=10)
                    if response.ok or response.status_code == 200:
                        # Setup the toast message in session state and trigger a refresh
                        st.session_state["toast_message"] = f"Entry saved! {activity_selected} — {hours_to_hhmm(duration_hours)} logged."
                        fetch_timesheets.clear()
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(f"Submission Error (Code {response.status_code}): Verify Google Form Settings.")
                except Exception as e:
                    st.error(f"Network Connection Error: {e}")


# =============================================================================
# SECTION 16: HISTORY REVIEW & COLOR-CODED TABLE
# =============================================================================

st.markdown("---")
st.subheader("📊 Review Period History")

if total_database_records > 0:
    if not current_period_df.empty:
        st.success(f"🔍 Found {len(current_period_df)} entries for this pay period.")

        col_history_table, col_history_stats = st.columns([3, 1])

        with col_history_table:
            rows_html = ""
            for _, row in current_period_df.iterrows():
                code     = str(row.get('Code', '')).strip()
                hrs_raw  = row.get('Hours', 0)
                hrs_disp = f"{float(hrs_raw):.2f} ({hours_to_hhmm(hrs_raw)})"
                badge    = code_badge(code)
                rows_html += f"""
                <tr>
                    <td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;">{row.get('Date','')}</td>
                    <td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;">{row.get('Time In','')}</td>
                    <td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;">{row.get('Time Out','')}</td>
                    <td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;max-width:200px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{row.get('Activity','')}</td>
                    <td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;text-align:center;">{badge}</td>
                    <td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;text-align:right;font-weight:600;">{hrs_disp}</td>
                </tr>"""

            table_html = f"""
            <div style="overflow-x:auto;">
            <table style="width:100%;border-collapse:collapse;font-size:14px;font-family:'Calibri',sans-serif;">
                <thead>
                    <tr style="background:#7B2CBF;color:white;">
                        <th style="padding:8px 10px;text-align:left;">Date</th>
                        <th style="padding:8px 10px;text-align:left;">Time In</th>
                        <th style="padding:8px 10px;text-align:left;">Time Out</th>
                        <th style="padding:8px 10px;text-align:left;">Activity</th>
                        <th style="padding:8px 10px;text-align:center;">Code</th>
                        <th style="padding:8px 10px;text-align:right;">Hours</th>
                    </tr>
                </thead>
                <tbody>{rows_html}</tbody>
            </table>
            </div>"""
            st.markdown(table_html, unsafe_allow_html=True)

        with col_history_stats:
            st.metric("Total Hours",   f"{running_hours:.2f} hrs")
            st.metric("HH:MM Format",  hours_to_hhmm(running_hours))
            st.metric("Total Minutes", f"{running_minutes} mins")
            st.metric("Entries",       len(current_period_df))

            if "Code" in current_period_df.columns and "Hours" in current_period_df.columns:
                st.markdown("**Hours by Code:**")
                code_summary = (current_period_df.groupby("Code")["Hours"]
                                .apply(lambda x: x.astype(float).sum())
                                .reset_index())
                for _, cs_row in code_summary.iterrows():
                    c   = str(cs_row["Code"])
                    h   = float(cs_row["Hours"])
                    bdg = code_badge(c)
                    st.markdown(f'{bdg} &nbsp; <strong>{h:.2f} hrs</strong> ({hours_to_hhmm(h)})', unsafe_allow_html=True)

        # =============================================================================
        # SECTION 17: EXCEL DOWNLOADS
        # =============================================================================
        st.markdown("### 📥 Download Excel Files")
        col_dl1, col_dl2 = st.columns(2)
        safe_name = instructor_input.replace(" ", "_")

        # Serialise period data for cache key (orient="records" handles dataframe rows cleanly)
        period_json = current_period_df.to_json(orient="records", date_format="iso")
        today_str   = TODAY.strftime("%m/%d/%Y")

        with col_dl1:
            ts_bytes = build_timesheet_bytes(instructor_input, pay_period_start, pay_period_end, period_json, today_str)
            st.download_button(
                label="📥 Download Timesheet (.xlsx)",
                data=ts_bytes,
                file_name=f"{safe_name}_Official_Timesheet_{pay_period_start}_to_{pay_period_end}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col_dl2:
            add_bytes = build_additional_bytes(instructor_input, pay_period_start, pay_period_end, period_json)
            st.download_button(
                label="📥 Download Additional Hours Report (.xlsx)",
                data=add_bytes,
                file_name=f"{safe_name}_Additional_Hours_{pay_period_start}_to_{pay_period_end}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.warning(
            f"ℹ️ **{total_database_records} total entries** found for your profile, but "
            f"**0 entries** fall within {pay_period_start.strftime('%m/%d/%Y')} – "
            f"{pay_period_end.strftime('%m/%d/%Y')}."
        )
        st.info("💡 Adjust the Start/End Date inputs above to reveal your logged entries.")


# =============================================================================
# SECTION 18: ADMIN PANEL  (only visible when logged in with master password)
# =============================================================================

if is_admin and not all_instructors_df.empty:
    st.markdown("---")
    st.markdown("## 🛡️ Admin Dashboard")
    st.caption("Visible only to admin accounts.")

    period_all = all_instructors_df[
        (all_instructors_df["ParsedDate"] >= pay_period_start) &
        (all_instructors_df["ParsedDate"] <= pay_period_end)
    ].copy()

    if not period_all.empty and "Instructor Name" in period_all.columns:
        st.markdown("### 👥 Team Hours — Current Period")
        registered_names = set(
            account_registry["Instructor Name"].astype(str).str.strip().tolist()
        ) if not account_registry.empty else set()

        summary = (
            period_all.groupby("Instructor Name")
            .agg(
                Hours=("Hours", lambda x: x.astype(float).sum()),
                Entries=("Hours", "count")
            )
            .reset_index()
            .sort_values("Hours", ascending=False)
        )
        summary["HH:MM"]      = summary["Hours"].apply(hours_to_hhmm)
        summary["Hours"]      = summary["Hours"].apply(lambda x: f"{x:.2f}")
        summary["⚠️ No Log"]  = summary["Instructor Name"].apply(lambda n: "✅" if n in registered_names else "—")

        logged_names    = set(period_all["Instructor Name"].astype(str).str.strip().tolist())
        missing_names   = registered_names - logged_names
        if missing_names:
            st.warning(
                f"⚠️ **{len(missing_names)} instructor(s) have not logged any time this period:** "
                + ", ".join(sorted(missing_names))
            )

        st.dataframe(summary, use_container_width=True, hide_index=True)

        st.markdown("### 📊 Team Hours by Code — Current Period")
        if "Code" in period_all.columns:
            code_totals = (
                period_all.groupby("Code")["Hours"]
                .apply(lambda x: x.astype(float).sum())
                .reset_index()
                .sort_values("Hours", ascending=False)
            )
            cols_admin = st.columns(len(code_totals))
            for i, (_, cr) in enumerate(code_totals.iterrows()):
                c    = str(cr["Code"])
                h    = float(cr["Hours"])
                bdg  = code_badge(c)
                with cols_admin[i]:
                    st.markdown(
                        f'<div style="text-align:center;padding:12px;background:#f8f5ff;'
                        f'border-radius:10px;">{bdg}<br>'
                        f'<span style="font-size:22px;font-weight:800;">{h:.2f}</span><br>'
                        f'<span style="font-size:12px;color:#64748b;">hrs ({hours_to_hhmm(h)})</span>'
                        f'</div>',
                        unsafe_allow_html=True
                    )

        with st.expander("📋 View Full Team Log for This Period"):
            cols_to_show = [c for c in ['Date','Instructor Name','Activity','Code','Hours','Minutes'] if c in period_all.columns]
            st.dataframe(period_all[cols_to_show].sort_values(["Instructor Name","Date"]), use_container_width=True, hide_index=True)
            
    else:
        st.info("ℹ️ No hours have been logged by any team members for this pay period yet.")
