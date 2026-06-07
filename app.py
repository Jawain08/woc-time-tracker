import streamlit as st
import pandas as pd
import io
import datetime
import os
import requests
import smtplib
import time
from email.mime.text import MIMEText
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# --- SYSTEM CONFIGURATION ---
st.set_page_config(page_title="WOC - Time Tracking System", layout="wide", page_icon="📝")

# =============================================================================
# SECTION 1: HEADER
# =============================================================================

def render_woc_header():
    logo_filename = "woclogo.png"
    if os.path.exists(logo_filename):
        st.image(logo_filename, width=280)
    st.markdown(
        """
        <div style="background-color: #7B2CBF; padding: 15px 20px; border-radius: 10px; margin-top: 10px; margin-bottom: 25px;">
            <h1 style="color: white; margin: 0; font-family: 'Calibri', sans-serif; font-size: 30px; font-weight: bold; letter-spacing: 0.5px;">Women of Colors, Inc.</h1>
            <p style="color: #E0AAFF; margin: 4px 0 0 0; font-size: 16px; font-family: 'Calibri', sans-serif;">Saginaw Community Prevention & Training Program Hub</p>
        </div>
        """,
        unsafe_allow_html=True
    )

# =============================================================================
# SECTION 2: COLUMN MAPPING HELPER
# =============================================================================

def map_columns(df, rules):
    """
    Maps raw dataframe columns to standardised names using a rules dict.
    rules = { "Standard Name": ["keyword1", "keyword2", ...], ... }
    Returns a new dataframe with only the standardised columns.
    """
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


# =============================================================================
# SECTION 3: CACHED DATA FETCHERS
# =============================================================================

SHEET_ID = "1zop4YKXKA1H8Iv89YwkGpP4c4YlGGFgz5jDYLT3psik"
TIMESHEETS_GID = "742432797"
ACCOUNTS_GID   = "1781560298"

def _csv_url(gid):
    return f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}"

@st.cache_data(ttl=60)
def fetch_timesheets():
    """Downloads and normalises the timesheet sheet. Cached for 60 s."""
    try:
        raw = pd.read_csv(_csv_url(TIMESHEETS_GID))
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
            # Positional fallback — warn the admin so silent mismap is visible
            if len(raw.columns) >= 11:
                df = raw.copy()
                df.columns = (
                    ["Timestamp", "Date", "Instructor Name", "Time In", "Time Out",
                     "Activity", "Code", "Category", "Description", "Minutes", "Hours"]
                    + list(raw.columns[11:])
                )
                st.warning(
                    "⚠️ Timesheet columns could not be auto-detected — falling back to "
                    "positional mapping. If data looks wrong, check the Google Sheet headers."
                )
            else:
                st.error("🛑 Timesheet sheet structure is unrecognised. Contact your admin.")
                return pd.DataFrame(columns=["Timestamp","Date","Instructor Name","Time In",
                                             "Time Out","Activity","Code","Category",
                                             "Description","Minutes","Hours"])

        # -----------------------------------------------------------------
        # MPHI → CARP normalisation data interceptor
        # -----------------------------------------------------------------
        if "Code" in df.columns:
            df["Code"] = (
                df["Code"].astype(str).str.strip()
                    .replace({"MPHI": "CARP", "mphi": "CARP"})
            )

        return df

    except Exception as e:
        st.error(f"🛑 Timesheet Database Connection Error: {e}")
        return pd.DataFrame(columns=["Timestamp","Date","Instructor Name","Time In",
                                     "Time Out","Activity","Code","Category",
                                     "Description","Minutes","Hours"])


@st.cache_data(ttl=60)
def fetch_accounts():
    """Downloads and normalises the accounts sheet. Cached for 60 s."""
    try:
        raw = pd.read_csv(_csv_url(ACCOUNTS_GID))
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
                df.columns = (
                    ["Timestamp", "Instructor Name", "Email Address", "PIN"]
                    + list(raw.columns[4:])
                )
                st.warning(
                    "⚠️ Accounts columns could not be auto-detected — falling back to "
                    "positional mapping. Check the Google Sheet headers if logins fail."
                )
            else:
                st.error("🛑 Accounts sheet structure is unrecognised. Contact your admin.")
                return pd.DataFrame(columns=["Timestamp","Instructor Name","Email Address","PIN"])

        return df

    except Exception as e:
        st.error(f"🛑 Account Profile Connection Error: {e}")
        return pd.DataFrame(columns=["Timestamp","Instructor Name","Email Address","PIN"])


# =============================================================================
# SECTION 4: EMAIL HELPER
# =============================================================================

def send_pin_email(recipient_email, recipient_name, user_pin):
    if "smtp" in st.secrets:
        try:
            msg = MIMEText(
                f"Hello {recipient_name},\n\n"
                f"Your requested PIN retrieval for the WOC Time Tracking Hub is: {user_pin}\n\n"
                f"Log in here: https://share.streamlit.io/\n\n"
                f"Regards,\nWomen of Colors Payroll Admin"
            )
            msg['Subject'] = "WOC Time Tracker - PIN Recovery"
            msg['From']    = st.secrets["smtp"]["username"]
            msg['To']      = recipient_email

            with smtplib.SMTP_SSL(st.secrets["smtp"]["server"],
                                  int(st.secrets["smtp"]["port"])) as server:
                server.login(st.secrets["smtp"]["username"],
                             st.secrets["smtp"]["password"])
                server.sendmail(st.secrets["smtp"]["username"],
                                [recipient_email], msg.as_string())
            return True, "Success"
        except Exception as e:
            return False, str(e)
    return False, "Fallback"


# =============================================================================
# SECTION 5: LOGIN / REGISTRATION PORTAL
# =============================================================================

existing_data    = fetch_timesheets()
account_registry = fetch_accounts()

if "user" in st.query_params and "logged_in" not in st.session_state:
    st.session_state["logged_in"] = True
    st.session_state["instructor_name"] = st.query_params["user"]

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
                login_name = st.text_input("Instructor Name:", placeholder="First & Last Name")
                login_pin  = st.text_input("Enter Personal PIN:", type="password",
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
                            st.session_state["logged_in"] = True
                            st.session_state["instructor_name"] = cleaned_name
                            st.query_params["user"] = cleaned_name
                            st.rerun()
                        else:
                            st.error("Authentication Error: Invalid PIN entered for this profile.")
                    else:
                        # Master password — read from secrets
                        master_pw = st.secrets.get("admin", {}).get("master_password", "")
                        if master_pw and login_pin.strip() == master_pw:
                            st.session_state["logged_in"] = True
                            st.session_state["instructor_name"] = cleaned_name
                            st.query_params["user"] = cleaned_name
                            st.rerun()
                        else:
                            st.error(
                                f"Profile Not Found: '{cleaned_name}' is not registered "
                                f"in our database yet."
                            )

        # ── REGISTRATION ─────────────────────────────────────────────────────
        elif portal_tab == "Create Custom Account / PIN":
            st.info("💡 Setting up your profile connects your name to a custom passcode "
                    "so your timesheets remain secure.")
            with st.form("registration_panel"):
                reg_name  = st.text_input("Full Instructor Name:", placeholder="First and Last Name")
                reg_email = st.text_input("Email Address:", placeholder="username@domain.com")
                reg_pin   = st.text_input("Create 4-to-6 Digit PIN:", type="password",
                                          placeholder="Choose your passcode")
                submit_reg = st.form_submit_button("📝 Register Account")

                if submit_reg:
                    if not reg_name.strip() or not reg_email.strip() or not reg_pin.strip():
                        st.error("Validation Error: All registry fields are strictly required.")
                    else:
                        ACC_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSdY6ydD4YLYQEicFkk21DIRefUTT5ht8v4lbdZVr6hSbGOBAA/formResponse"
                        acc_data = {
                            "entry.576544689":  reg_name.strip(),
                            "entry.836662014":  reg_email.strip(),
                            "entry.2099667226": reg_pin.strip()
                        }
                        try:
                            res = requests.post(ACC_FORM_URL, data=acc_data)
                            if res.ok or res.status_code == 200:
                                st.success("Account securely instantiated! Switch to 'Sign In' to enter.")
                            else:
                                st.error(f"Database Error (Code {res.status_code}): "
                                         f"Verify Form Settings layout parameters.")
                        except Exception as e:
                            st.error(f"Network Connection Failed: {e}")

        # ── PIN RECOVERY ──────────────────────────────────────────────────────
        elif portal_tab == "Forgot PIN / Reset Option":
            with st.form("recovery_panel"):
                recover_email = st.text_input("Enter Your Registered Email Address:",
                                              placeholder="username@domain.com")
                submit_recovery = st.form_submit_button("🔍 Retrieve Access Passcode")

                if submit_recovery:
                    matched_emails = account_registry[
                        account_registry["Email Address"]
                            .astype(str).str.strip().str.lower()
                            == recover_email.strip().lower()
                    ]
                    if not matched_emails.empty:
                        user_account = matched_emails.iloc[-1]
                        found_name   = user_account["Instructor Name"]
                        found_pin    = user_account["PIN"]

                        status, message = send_pin_email(recover_email.strip(),
                                                         found_name, found_pin)
                        if status:
                            st.success(
                                f"📬 Recovery instructions dispatched to {recover_email.strip()}."
                            )
                        else:
                            st.warning("⚙️ Automated email is offline. Showing PIN below:")
                            st.info(f"Account: **{found_name}** | PIN: `{found_pin}`")
                    else:
                        st.error("Verification Mismatch: That email is not in our registry.")
    st.stop()


# =============================================================================
# SECTION 6: AUTHENTICATED SESSION
# =============================================================================

render_woc_header()
instructor_input = st.session_state["instructor_name"]

col_user1, col_user2 = st.columns([3, 1])
with col_user1:
    st.markdown(f"#### Welcome back, **{instructor_input}**! 👋")
with col_user2:
    if st.button("🚪 Log Out / Clear Session", use_container_width=True):
        st.session_state.clear()
        st.query_params.clear()
        st.rerun()

# =============================================================================
# SECTION 7: PAY PERIOD NAVIGATION
# =============================================================================

if "period_offset" not in st.session_state:
    st.session_state.period_offset = 0

ANCHOR_DATE        = datetime.date(2026, 5, 24)
TODAY              = datetime.date.today()
days_since_anchor  = (TODAY - ANCHOR_DATE).days
completed_periods  = days_since_anchor // 14

active_period_index = completed_periods + st.session_state.period_offset
auto_period_start   = ANCHOR_DATE + datetime.timedelta(days=active_period_index * 14)
auto_period_end     = auto_period_start + datetime.timedelta(days=13)

st.subheader("🗓️ Pay Period Review Settings")

col_nav1, col_nav2, col_nav3 = st.columns([1, 2, 1])
with col_nav1:
    if st.button("⬅️ Previous Period", use_container_width=True):
        st.session_state.period_offset -= 1
        st.rerun()
with col_nav2:
    col_date, col_reset = st.columns([2, 1])
    with col_date:
        st.markdown(
            f"<h4 style='text-align: center; margin-top: 5px; color: #7B2CBF;'>"
            f"{auto_period_start.strftime('%b %d')} — {auto_period_end.strftime('%b %d, %Y')}"
            f"</h4>",
            unsafe_allow_html=True
        )
    with col_reset:
        if st.session_state.period_offset != 0:
            if st.button("🔄 Current", use_container_width=True,
                         help="Jump back to the active pay period"):
                st.session_state.period_offset = 0
                st.rerun()
with col_nav3:
    if st.button("Next Period ➡️", use_container_width=True):
        st.session_state.period_offset += 1
        st.rerun()

col_profile1, col_profile2 = st.columns(2)
with col_profile1:
    pay_period_start = st.date_input("Start Date (editable override)", value=auto_period_start)
with col_profile2:
    pay_period_end   = st.date_input("End Date (editable override)",   value=auto_period_end)

if pay_period_start != auto_period_start or pay_period_end != auto_period_end:
    st.info("ℹ️ You have manually overridden the pay period dates. "
            "Click ⬅️ / ➡️ or 🔄 Current to return to automatic mode.")

st.markdown("---")


# =============================================================================
# SECTION 8: FILTER DATA FOR CURRENT INSTRUCTOR & PERIOD
# =============================================================================

total_database_records = 0
current_period_df      = pd.DataFrame()
running_hours          = 0.0
running_minutes        = 0

if instructor_input.strip() and not existing_data.empty and "Instructor Name" in existing_data.columns:
    user_filtered_df = existing_data[
        existing_data["Instructor Name"]
            .astype(str).str.strip().str.lower() == instructor_input.strip().lower()
    ].copy()

    if not user_filtered_df.empty and "Date" in user_filtered_df.columns:
        user_filtered_df["ParsedDate"] = pd.to_datetime(
            user_filtered_df["Date"], errors='coerce'
        ).dt.date
        user_filtered_df = user_filtered_df.dropna(subset=["ParsedDate"])

        current_period_df = user_filtered_df[
            (user_filtered_df["ParsedDate"] >= pay_period_start) &
            (user_filtered_df["ParsedDate"] <= pay_period_end)
        ].sort_values(by="ParsedDate", ascending=True)

        total_database_records = len(user_filtered_df)
        running_hours   = current_period_df['Hours'].astype(float).sum() \
                          if 'Hours'   in current_period_df.columns else 0.0
        running_minutes = current_period_df['Minutes'].astype(int).sum() \
                          if 'Minutes' in current_period_df.columns else 0


# =============================================================================
# SECTION 9: ACTIVITY DICTIONARY & TIME SLOTS
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
# SECTION 10: DAILY LOG ENTRY FORM
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
        duration_minutes = int((end_time_dt - start_time_dt).total_seconds() / 60)
        duration_hours   = round(duration_minutes / 60, 2)
        mapping_result   = activity_to_code_mapping[activity_selected]

        FORM_URL  = "https://docs.google.com/forms/d/1G8flLQrWJWGl5CwOEUe48zuAPre5mhJrbanx33uSkZk/formResponse"
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
        try:
            response = requests.post(FORM_URL, data=form_data)
            if response.ok or response.status_code == 200:
                st.success("Entry securely saved to central database sheet!")
                fetch_timesheets.clear()
                time.sleep(1)
                st.rerun()
            else:
                st.error(f"Submission Error (Code {response.status_code}): "
                         f"Verify Google Form Settings.")
        except Exception as e:
            st.error(f"Network Connection Error: {e}")


# =============================================================================
# SECTION 11: HISTORY REVIEW & EXCEL EXPORT
# =============================================================================

st.markdown("---")
st.subheader("📊 Review Period History")

if total_database_records > 0:
    if not current_period_df.empty:
        st.success(f"🔍 Found {len(current_period_df)} entries for this pay period.")
        col_history_table, col_history_stats = st.columns([3, 1])
        with col_history_table:
            display_df = current_period_df[['Date','Time In','Time Out','Activity','Code','Hours']].copy()
            st.dataframe(display_df, use_container_width=True, hide_index=True)
        with col_history_stats:
            st.metric(label="Total Hours Tracked",   value=f"{running_hours:.2f} hrs")
            st.metric(label="Total Minutes Tracked", value=f"{running_minutes} mins")

        st.markdown("### 📥 Download Excel Files")
        col_dl1, col_dl2 = st.columns(2)
        safe_name = instructor_input.replace(" ", "_")

        # ── EXPORT HELPER: shared styles ─────────────────────────────────────

        def _make_styles():
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

        # ── EXPORT 1: OFFICIAL TIMESHEET ─────────────────────────────────────

        def build_timesheet_workbook():
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
            ws["D3"] = f"Name :  {instructor_input}";  ws["D3"].font = S["regular"]
            ws["F3"] = "Email: payroll@yeoandyeo.com";  ws["F3"].font = S["regular"]
            ws["A4"] = "Manager Details:";              ws["A4"].font = S["bold"]
            ws["D4"] = "Name: Vicki Hill";              ws["D4"].font = S["regular"]
            ws["F4"] = "Fax: 989-793-0186";             ws["F4"].font = S["regular"]
            ws["A5"] = f"Period Start Date: {pay_period_start.strftime('%m/%d/%Y')}"; ws["A5"].font = S["bold"]
            ws["E5"] = f"Period End Date:  {pay_period_end.strftime('%m/%d/%Y')}";    ws["E5"].font = S["bold"]

            # Row 7 — group headers
            for col_idx, text in enumerate(
                ["", "", "", "", "", "Total Hours", "NOFA", "WOC", "JJ", "TRICAP", "CARP"], 1
            ):
                if text:
                    c = ws.cell(row=7, column=col_idx, value=text)
                    c.font      = S["bold"]
                    c.alignment = Alignment(horizontal="center", wrap_text=True)

            # Row 9 — column headers
            for col_idx, text in enumerate(
                ["", "Day", "Date", "Time In", "Time Out", "Hours Worked",
                 "NOFA", "WOC", "JJ", "TRICAP", "CARP"], 1
            ):
                if text:
                    c = ws.cell(row=9, column=col_idx, value=text)
                    c.font      = S["bold"]
                    c.alignment = Alignment(horizontal="center")

            code_col_map = {"NOFA": 7, "WOC": 8, "JJ": 9, "TRICAP": 10, "CARP": 11}
            days_in_period = max(1, (pay_period_end - pay_period_start).days + 1)
            date_list      = [pay_period_start + datetime.timedelta(days=x)
                              for x in range(days_in_period)]

            row_index = 10
            for idx, d in enumerate(date_list):
                if idx == 7:
                    row_index += 1  # blank row between weeks

                ws.cell(row=row_index, column=2, value=d.strftime("%A")).font  = S["regular"]
                ws.cell(row=row_index, column=3, value=d.strftime("%Y-%m-%d")).font = S["regular"]

                day_logs = current_period_df[current_period_df['ParsedDate'] == d]

                if not day_logs.empty:
                    ws.cell(row=row_index, column=4,
                            value=" / ".join(day_logs['Time In'].astype(str).tolist())).font  = S["regular"]
                    ws.cell(row=row_index, column=5,
                            value=" / ".join(day_logs['Time Out'].astype(str).tolist())).font = S["regular"]
                    ws.cell(row=row_index, column=6,
                            value=day_logs['Hours'].astype(float).sum()).font                 = S["regular"]

                    for c_idx in range(7, 12):
                        c = ws.cell(row=row_index, column=c_idx, value=0)
                        c.font   = S["regular"]
                        c.fill   = S["shaded"]
                        c.border = S["thin"]

                    for _, row_log in day_logs.iterrows():
                        code        = str(row_log.get('Code', ''))
                        hours_worked = float(row_log.get('Hours', 0.0))
                        if code in code_col_map:
                            c_idx       = code_col_map[code]
                            current_val = ws.cell(row=row_index, column=c_idx).value or 0.0
                            active_cell = ws.cell(row=row_index, column=c_idx,
                                                  value=current_val + hours_worked)
                            active_cell.fill = PatternFill(fill_type=None)
                else:
                    for c_idx in range(4, 12):
                        c = ws.cell(row=row_index, column=c_idx, value=0)
                        c.font   = S["regular"]
                        c.fill   = S["shaded"]
                        c.border = S["thin"]

                row_index += 1

            row_index += 1
            ws.cell(row=row_index, column=2, value="Total Target Hours:").font  = S["bold"]
            ws.cell(row=row_index, column=3, value=75.0).font                   = S["bold"]
            row_index += 1
            ws.cell(row=row_index, column=2, value="Actual Hours Worked:").font = S["bold"]
            ws.cell(row=row_index, column=3, value=running_hours).font          = S["bold"]
            ws.cell(row=row_index, column=6, value=running_hours).font          = S["bold"]

            row_index += 2
            ws.merge_cells(start_row=row_index, start_column=2,
                           end_row=row_index,   end_column=11)
            cert = ws.cell(row=row_index, column=2,
                           value="CLIENT: I CERTIFY THAT THE HOURS WORKED ON THIS TIME SLIP ARE CORRECT.")
            cert.font      = S["bold"]
            cert.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[row_index].height = 24

            row_index += 2
            ws.cell(row=row_index, column=2, value="Employee Signature:").font   = S["bold"]
            ws.cell(row=row_index, column=3, value=instructor_input).font        = S["cursive"]
            ws.cell(row=row_index, column=5, value="Date:").font                 = S["bold"]
            ws.cell(row=row_index, column=6,
                    value=datetime.date.today().strftime("%m/%d/%Y")).font        = S["regular"]
            row_index += 2
            ws.cell(row=row_index, column=2, value="Manager Signature:").font    = S["bold"]
            ws.cell(row=row_index, column=5, value="Date:").font                 = S["bold"]

            for col in ws.columns:
                max_len    = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    val_str = str(cell.value or '')
                    if cell.row in [1, 2, 3, 4, 5] or "CLIENT: I CERTIFY" in val_str:
                        continue
                    max_len = max(max_len, len(val_str))
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

            return wb

        # ── EXPORT 2: ADDITIONAL HOURS REPORT ────────────────────────────────

        def build_additional_hours_workbook():
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

            ws["A1"] = (f"Additional Hours Report FY20 - Due "
                        f"{pay_period_start.strftime('%m/%d/%y')} - "
                        f"{pay_period_end.strftime('%m/%d/%Y')}")
            ws["A1"].font = S["bold"]
            ws["A2"] = "Agency Name";              ws["A2"].font = S["bold"]
            ws["C2"] = "ADDITIONAL HOURS REPORT";  ws["C2"].font = S["bold"]

            for col_idx, h in enumerate(
                ["Date", "Staff Name", "Category", "Description", "Time in minutes"], 1
            ):
                c = ws.cell(row=3, column=col_idx, value=h)
                c.font      = S["bold"]
                c.alignment = Alignment(horizontal="left")

            curr_row = 4
            for _, log in current_period_df.iterrows():
                ws.cell(row=curr_row, column=1, value=str(log.get('Date',        ''))).font = S["regular"]
                ws.cell(row=curr_row, column=2, value=instructor_input).font                = S["regular"]
                ws.cell(row=curr_row, column=3, value=str(log.get('Category',    ''))).font = S["regular"]
                ws.cell(row=curr_row, column=4, value=str(log.get('Description', ''))).font = S["regular"]
                ws.cell(row=curr_row, column=5, value=int(log.get('Minutes', 0))).font      = S["regular"]
                curr_row += 1

            ws.cell(row=curr_row, column=4, value="Total").font         = S["bold"]
            ws.cell(row=curr_row, column=5, value=running_minutes).font = S["bold"]

            for col in ws.columns:
                max_len    = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

            return wb

        # ── RENDER DOWNLOAD BUTTONS ───────────────────────────────────────────

        with col_dl1:
            buf1 = io.BytesIO()
            build_timesheet_workbook().save(buf1)
            st.download_button(
                label="📥 Download Template-Matched Timesheet (.xlsx)",
                data=buf1.getvalue(),
                file_name=f"{safe_name}_Official_Timesheet_{pay_period_start}_to_{pay_period_end}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col_dl2:
            buf2 = io.BytesIO()
            build_additional_hours_workbook().save(buf2)
            st.download_button(
                label="📥 Download Additional Hours Report (.xlsx)",
                data=buf2.getvalue(),
                file_name=f"{safe_name}_Additional_Hours_{pay_period_start}_to_{pay_period_end}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.warning(
            f"ℹ️ Pay Period Filter Notice: We detected **{total_database_records} entries total** "
            f"linked to your profile, but **0 entries** fall within the selected period "
            f"({pay_period_start.strftime('%m/%d/%Y')} to {pay_period_end.strftime('%m/%d/%Y')})."
        )
        st.info(
            "💡 **Solution:** Adjust the **Start Date** or **End Date** inputs above to cover "
            "the dates of the entries you logged. The history table and export buttons will appear."
        )
