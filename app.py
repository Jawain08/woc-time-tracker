import streamlit as st
import pandas as pd
import io
import datetime
import os
import requests

# --- SYSTEM CONFIGURATION ---
st.set_page_config(page_title="WOC - Time Tracking System", layout="wide", page_icon="📝")

# --- CUSTOM THEMED APPLICATION HEADER & LOGO ---
if os.path.exists("woc_logo.png"):
    st.markdown(
        """
        <div style="background-color: #7B2CBF; padding: 15px 25px; border-radius: 10px; margin-bottom: 25px; display: flex; align-items: center; gap: 25px;">
            <div style="background-color: white; padding: 8px; border-radius: 8px; display: flex; align-items: center; justify-content: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                <img src="app/static/woc_logo.png" width="100" style="display: block; max-height: 80px; object-fit: contain;">
            </div>
            <div>
                <h1 style="color: white; margin: 0; font-family: 'Calibri', sans-serif; font-size: 32px; font-weight: bold; letter-spacing: 0.5px;">Women of Colors, Inc.</h1>
                <p style="color: #E0AAFF; margin: 4px 0 0 0; font-size: 16px; font-family: 'Calibri', sans-serif;">Saginaw Community Prevention & Training Program Hub</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
else:
    st.markdown(
        """
        <div style="background-color: #7B2CBF; padding: 20px; border-radius: 10px; margin-bottom: 25px;">
            <h1 style="color: white; margin: 0; font-family: 'Calibri', sans-serif;">Women of Colors, Inc.</h1>
            <p style="color: #E0AAFF; margin: 5px 0 0 0; font-size: 16px;">Saginaw Community Prevention & Training Program Hub</p>
        </div>
        """,
        unsafe_allow_html=True
    )

# --- LIVE REFRESH DATA ENGINE ---
# PASTE YOUR GOOGLE SHEET LINK HERE: Replace with your actual "Publish to Web" CSV link
PUBLIC_CSV_URL = "PASTE_YOUR_PUBLISHED_CSV_LINK_HERE"

try:
    # Safely pull down raw data entries directly from your spreadsheet web stream
    existing_data = pd.read_csv(PUBLIC_CSV_URL)
    # Ensure all columns are consistently mapped and named
    existing_data.columns = ["Timestamp", "Date", "Instructor Name", "Time In", "Time Out", "Activity", "Code", "Category", "Description", "Minutes", "Hours"]
except Exception:
    # Fallback to empty framework layout if the web asset isn't fully published yet
    existing_data = pd.DataFrame(columns=["Timestamp", "Date", "Instructor Name", "Time In", "Time Out", "Activity", "Code", "Category", "Description", "Minutes", "Hours"])

activity_to_code_mapping = {
    "Prime For Life instructor Training (Juvenile)": {"code": "JJ", "category": "Other", "description": "Prime For Life Instructor Training - Juvenile"},
    "Prime For Life instructor Training (Tri-Cap)":   {"code": "TRICAP", "category": "Other", "description": "Prime For Life Instructor Training - Tri-Cap"},
    "Prime For Life instructor Training (Notes Update)": {"code": "TRICAP", "category": "Other", "description": "Prime For Life Instructor Training - Notes"},
    "Botvin Life Skills Training":                    {"code": "BOTVIN", "category": "Other", "description": "Botvin Life Skills Training"},
    "Prevention Team Meeting":                        {"code": "NOFA",   "category": "Other", "description": "Prevention Team Meeting"}
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

# --- STEP 1: AUTOMATED PAY PERIOD ENGINE ---
ANCHOR_DATE = datetime.date(2026, 5, 23)
TODAY = datetime.date.today()
days_since_anchor = (TODAY - ANCHOR_DATE).days
completed_periods = days_since_anchor // 14

auto_period_start = ANCHOR_DATE + datetime.timedelta(days=completed_periods * 14)
auto_period_end = auto_period_start + datetime.timedelta(days=13)

# --- STEP 2: PROFILE CONFIGURATION ---
st.subheader("👤 Employee & Pay Period Details")
col_profile1, col_profile2, col_profile3 = st.columns(3)

with col_profile1:
    instructor_input = st.text_input("Instructor Name:", value="", placeholder="e.g. Jawain Swint")
with col_profile2:
    pay_period_start = st.date_input("Start Date", value=auto_period_start)
with col_profile3:
    pay_period_end = st.date_input("End Date", value=auto_period_end)

st.markdown("---")

# Filter database rows for the currently typed instructor instantly
if instructor_input.strip() and not existing_data.empty:
    user_filtered_df = existing_data[existing_data["Instructor Name"].str.lower() == instructor_input.strip().lower()].copy()
    user_filtered_df["ParsedDate"] = pd.to_datetime(user_filtered_df["Date"]).dt.date
    current_period_df = user_filtered_df[(user_filtered_df["ParsedDate"] >= pay_period_start) & (user_filtered_df["ParsedDate"] <= pay_period_end)]
    
    running_hours = current_period_df['Hours'].astype(float).sum()
    running_minutes = current_period_df['Minutes'].astype(int).sum()
else:
    current_period_df = pd.DataFrame()
    running_hours = 0.0
    running_minutes = 0

# --- STEP 3: DAILY DATA LOG ENTRY FORM ---
st.subheader("⏳ Log Daily Activity")
with st.form("daily_time_entry_form", clear_on_submit=True):
    entry_col1, entry_col2, entry_col3, entry_col4 = st.columns(4)
    
    with entry_col1:
        entry_date = st.date_input("Date Worked", value=pay_period_start, min_value=pay_period_start, max_value=pay_period_end)
    with entry_col2:
        time_in_str = st.selectbox("Time In", options=time_dropdown_options, index=67)
    with entry_col3:
        time_out_str = st.selectbox("Time Out", options=time_dropdown_options, index=74)
    with entry_col4:
        activity_selected = st.selectbox("Activity Classification", all_activities)
        
    add_btn = st.form_submit_button("➕ Save Entry to Log")

if add_btn:
    if not instructor_input.strip():
        st.error("Validation Error: You must enter your 'Instructor Name' before saving entries.")
    else:
        start_time_dt = datetime.datetime.strptime(f"{entry_date} {time_in_str}", "%Y-%m-%d %I:%M %p")
        end_time_dt = datetime.datetime.strptime(f"{entry_date} {time_out_str}", "%Y-%m-%d %I:%M %p")
        
        if end_time_dt <= start_time_dt:
            st.error("Validation Error: 'Time Out' must occur after 'Time In'.")
        else:
            duration_delta = end_time_dt - start_time_dt
            duration_minutes = int(duration_delta.total_seconds() / 60)
            duration_hours = round(duration_minutes / 60, 2)
            
            mapping_result = activity_to_code_mapping.get(activity_selected)
            
            # Direct target submission URL link using your Form ID
            FORM_URL = "https://docs.google.com/forms/d/1G8flLQrWJWGl5CwOEUe48zuAPre5mhJrbanx33uSkZk/formResponse"
            
            # Verified Entry IDs mapped exactly from your inspector window
            form_data = {
                "entry.1205527392": entry_date.strftime("%Y-%m-%d"), # Date
                "entry.1822017875": instructor_input.strip(),        # Instructor Name
                "entry.1148008178": time_in_str,                      # Time In
                "entry.1036423098": time_out_str,                     # Time Out
                "entry.1565734482": activity_selected,                # Activity
                "entry.1863736208": mapping_result['code'],           # Code
                "entry.835834590": mapping_result['category'],       # Category
                "entry.693720626": mapping_result['description'],    # Description
                "entry.2039394575": duration_minutes,                 # Minutes
                "entry.1380701779": duration_hours                    # Hours
            }
            
            # Modern verification engine block
            try:
                response = requests.post(FORM_URL, data=form_data)
                if response.status_code == 200 or response.ok:
                    st.success("Entry securely saved to central database sheet!")
                    st.rerun()
                else:
                    st.error(f"Submission Error (Code {response.status_code}): Verify Google Form Settings -> Responses -> 'Limit to 1 response' is turned OFF.")
            except Exception as e:
                st.error(f"Network Connection Error: {e}")

# --- STEP 4: REVIEW HISTORY & EXPORT PANELS ---
if not current_period_df.empty:
    st.markdown("---")
    st.subheader("📊 Review Period History")
    
    col_history_table, col_history_stats = st.columns([3, 1])
    
    with col_history_table:
        display_df = current_period_df[['Date', 'Time In', 'Time Out', 'Activity', 'Code', 'Hours']].copy()
        st.dataframe(display_df, use_container_width=True, hide_index=True)
        
    with col_history_stats:
        st.metric(label="Total Hours Tracked", value=f"{running_hours:.2f} hrs")
        st.metric(label="Total Minutes Tracked", value=f"{running_minutes} mins")

    st.markdown("### 📥 Down Excel Files")
    col_dl1, col_dl2 = st.columns(2)
    safe_name = instructor_input.replace(" ", "_")

    # --- EXPORT 1: TIMESHEET GENERATOR ---
    with col_dl1:
        wb = Workbook()
        ws = wb.active
        ws.title = "Time Sheet"
        
        font_title = Font(name="Calibri", size=14, bold=True)
        font_bold = Font(name="Calibri", size=11, bold=True)
        font_regular = Font(name="Calibri", size=11)
        font_small = Font(name="Calibri", size=9, italic=True)
        
        ws["A1"] = "BiWeekly Employee Time Sheet"
        ws["A1"].font = font_title
        ws["A2"] = "Women of Colors"
        ws["A2"].font = font_bold
        
        ws["A3"] = "Employee Details:"
        ws["A3"].font = font_bold
        ws["D3"] = f"Name :  {instructor_input}"
        ws["D3"].font = font_regular
        ws["F3"] = "Email: payroll@yeoandyeo.com"
        ws["F3"].font = font_regular
        
        ws["A4"] = "Manager Details:"
        ws["A4"].font = font_bold
        ws["D4"] = "Name: Vicki Hill"
        ws["D4"].font = font_regular
        ws["F4"] = "Fax: 989-793-0186"
        ws["F4"].font = font_regular
        
        ws["A5"] = f"Period Start Date: {pay_period_start.strftime('%m/%d/%Y')}"
        ws["A5"].font = font_bold
        ws["E5"] = f"Period End Date:  {pay_period_end.strftime('%m/%d/%Y')}"
        ws["E5"].font = font_bold

        headers_r7 = ["", "", "Total Work Week Hours", "Total Hours Worked", "Regular Hours", "Overtime Hours", "NOFA", "WOC", "JJ", "TRICAP", "BOTVIN"]
        for col_idx, text in enumerate(headers_r7, 1):
            cell = ws.cell(row=7, column=col_idx, value=text)
            cell.font = font_bold
            cell.alignment = Alignment(horizontal="center", wrap_text=True)

        headers_r9 = ["", "", "Date(s)", "Time In", "Time out", "Time In", "Time Out", "Hours Worked", "NOFA", "WOC", "JJ", "TRICAP", "BOTVIN"]
        for col_idx, text in enumerate(headers_r9, 1):
            cell = ws.cell(row=9, column=col_idx, value=text)
            cell.font = font_bold
            cell.alignment = Alignment(horizontal="center")

        date_list = [pay_period_start + datetime.timedelta(days=x) for x in range(14)]
        row_index = 10
        week_1_hours = 0.0
        week_2_hours = 0.0

        for idx, d in enumerate(date_list):
            if idx == 7:
                row_index += 1
            
            day_name = d.strftime("%A")
            date_str = d.strftime("%Y-%m-%d")
            
            ws.cell(row=row_index, column=2, value=day_name).font = font_regular
            ws.cell(row=row_index, column=3, value=date_str).font = font_regular
            
            day_logs = current_period_df[current_period_df['ParsedDate'] == d]
            
            if not day_logs.empty:
                log_entry = day_logs.iloc[0]
                ws.cell(row=row_index, column=4, value=log_entry['Time In']).font = font_regular
                ws.cell(row=row_index, column=5, value=log_entry['Time Out']).font = font_regular
                
                hours_worked = float(log_entry['Hours'])
                ws.cell(row=row_index, column=8, value=hours_worked).font = font_regular
                
                code = log_entry['Code']
                code_col_map = {"NOFA": 9, "WOC": 10, "JJ": 11, "TRICAP": 12, "BOTVIN": 13}
                if code in code_col_map:
                    ws.cell(row=row_index, column=code_col_map[code], value=hours_worked).font = font_regular
                
                if idx < 7:
                    week_1_hours += hours_worked
                else:
                    week_2_hours += hours_worked
            else:
                ws.cell(row=row_index, column=8, value=0).font = font_regular
                ws.cell(row=row_index, column=9, value=0).font = font_regular
                ws.cell(row=row_index, column=11, value=0).font = font_regular
                ws.cell(row=row_index, column=12, value=0).font = font_regular
            
            row_index += 1

        ws["C8"] = 37.5
        ws["D8"] = week_1_hours
        ws["C8"].font = font_bold
        ws["D8"].font = font_bold
        
        ws["C18"] = 37.5
        ws["D18"] = week_2_hours
        ws["C18"].font = font_bold
        ws["D18"].font = font_bold

        row_index += 1
        cert_cell = ws.cell(row=row_index, column=2, value="CLIENT: I CERTIFY THAT THE HOURS WORKED ON THIS TIME SLIP ARE CORRECT.")
        cert_cell.font = font_bold
        cert_cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[row_index].height = 28
        
        row_index += 1
        ws.cell(row=row_index, column=2, value=instructor_input).font = font_regular
        ws.cell(row=row_index, column=5, value=datetime.date.today().strftime("%Y-%m-%d")).font = font_regular
        
        row_index += 1
        ws.cell(row=row_index, column=2, value="Employee Signature").font = font_small
        ws.cell(row=row_index, column=5, value="Date").font = font_small

        row_index += 2
        ws.cell(row=row_index, column=2, value="Manager Signature").font = font_small
        ws.cell(row=row_index, column=5, value="Date").font = font_small

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if cell.row > 24:
                    continue
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        buffer_grid = io.BytesIO()
        wb.save(buffer_grid)
        
        st.download_button(
            label="📥 Download Template-Matched Timesheet (.xlsx)",
            data=buffer_grid.getvalue(),
            file_name=f"{safe_name}_Official_Timesheet_{pay_period_start}_to_{pay_period_end}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- EXPORT 2: ADDITIONAL HOURS REPORT ---
    with col_dl2:
        wb_add = Workbook()
        ws_add = wb_add.active
        ws_add.title = "Report Form"
        
        font_add_bold = Font(name="Calibri", size=11, bold=True)
        font_add_reg = Font(name="Calibri", size=11)
        
        p_start_str = pay_period_start.strftime("%m/%d/%y")
        p_end_str = pay_period_end.strftime("%m/%d/%Y")
        ws_add["A1"] = f"Additional Hours Report FY20 - Due {p_start_str} - {p_end_str}"
        ws_add["A1"].font = font_add_bold
        
        ws_add["A2"] = "Agency Name"
        ws_add["C2"] = "ADDITIONAL HOURS REPORT"
        ws_add["A2"].font = font_add_bold
        
        add_headers = ["Date", "Staff Name", "Category", "Description", "Time in minutes"]
        for col_idx, h_text in enumerate(add_headers, 1):
            cell = ws_add.cell(row=3, column=col_idx, value=h_text)
            cell.font = font_add_bold
            cell.alignment = Alignment(horizontal="left")
            
        curr_row = 4
        for idx, log in current_period_df.iterrows():
            ws_add.cell(row=curr_row, column=1, value=str(log['Date'])).font = font_add_reg
            ws_add.cell(row=curr_row, column=2, value=instructor_input).font = font_add_reg
            ws_add.cell(row=curr_row, column=3, value=log['Category']).font = font_add_reg
            ws_add.cell(row=curr_row, column=4, value=log['Description']).font = font_add_reg
            ws_add.cell(row=curr_row, column=5, value=int(log['Minutes'])).font = font_add_reg
            curr_row += 1
            
        ws_add.cell(row=curr_row, column=4, value="Total").font = font_add_bold
        ws_add.cell(row=curr_row, column=5, value=running_minutes).font = font_add_bold
        
        for col in ws_add.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws_add.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
        buffer_additional = io.BytesIO()
        wb_add.save(buffer_additional)
            
        st.download_button(
            label="📥 Download Additional Hours Report (.xlsx)",
            data=buffer_additional.getvalue(),
            file_name=f"{safe_name}_Additional_Hours_{pay_period_start}_to_{pay_period_end}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("No time entries logged yet for this instructor in this pay period. Type your name above and add your hours to open the Down Excel Files console panels.")
