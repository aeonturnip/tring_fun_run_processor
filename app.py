import streamlit as st
import pandas as pd
import io
from streamlit_gsheets import GSheetsConnection
from openpyxl.styles import Font, Border, Side, PatternFill
import datetime

# 1. Page Configuration & Branding
st.set_page_config(page_title="Tring Fun Run Tools", layout="wide")
st.title("🏃‍♂️ Tring Fun Run: Entry & Results Dashboard")

# 2. Establish Cloud Connection
conn = st.connection("gsheets", type=GSheetsConnection)

def load_gsheet_memory(worksheet_name):
    try:
        return conn.read(worksheet=worksheet_name).fillna('')
    except:
        return pd.DataFrame(columns=["Raw Name", "Cleaned Name"])

# 3. Define Main Interface Tabs
tab_entry, tab_timer = st.tabs(["📋 Race Entry Processor", "⏱️ Timer Reconciliation"])

# --- TAB 1: RACE ENTRY PROCESSOR ---
with tab_entry:
    st.header("Entry Formatter & Audit")
    st.info("Upload the raw registration CSV to clean names and generate the zebra-striped race pack.")
    
    st.sidebar.header("Registration Upload")
    uploaded_entries = st.sidebar.file_uploader("Upload Raw Tring Race CSV", type=['csv'], key="entry_upload")

    YEAR_ORDER = {
        'Pre-school': 0, 'Reception': 1, 'Year 1': 2, 'Year 2': 3, 'Year 3': 4, 
        'Year 4': 5, 'Year 5': 6, 'Year 6': 7, 'Year 7': 8, 'Year 8': 9, 'Year 9': 10
    }

    if uploaded_entries is not None:
        df = pd.read_csv(uploaded_entries).fillna('')
        total_input_rows = len(df)
        
        # Name Pre-processing
        name_split = df['Full name'].str.split(' ', n=1, expand=True)
        df.insert(2, 'Forename', name_split[0].str.strip().str.title())
        df.insert(3, 'Surname', name_split[1].fillna('').str.strip().str.title())
        df.insert(0, 'Race Number', '')

        if 'School name' in df.columns:
            df['School name'] = df['School name'].astype(str).str.strip().str.title()
        if 'Team name' in df.columns:
            df['Team name'] = df['Team name'].astype(str).str.strip().str.title()

        # Persistent Memory
        school_mem_df = load_gsheet_memory("Schools")
        team_mem_df = load_gsheet_memory("Teams")
        school_memory = dict(zip(school_mem_df['Raw Name'], school_mem_df['Cleaned Name']))
        team_memory = dict(zip(team_mem_df['Raw Name'], team_mem_df['Cleaned Name']))

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("School Cleaning")
            unique_schools = [s for s in df['School name'].unique() if str(s).strip() != '' and str(s).lower() != 'nan']
            school_mapped = [school_memory.get(s, s) for s in unique_schools]
            school_editor_df = pd.DataFrame({"Raw Name": unique_schools, "Cleaned Name": school_mapped}).sort_values(by='Raw Name')
            edited_schools = st.data_editor(school_editor_df, hide_index=True, use_container_width=True, key="sch_ed")
        with col2:
            st.subheader("Team Cleaning")
            unique_teams = [t for t in df.get('Team name', pd.Series([])).unique() if str(t).strip() != '' and str(t).lower() != 'nan']
            team_mapped = [team_memory.get(t, t) for t in unique_teams]
            team_df = pd.DataFrame({"Raw Name": unique_teams, "Cleaned Name": team_mapped}).sort_values(by='Raw Name')
            edited_teams = st.data_editor(team_df, hide_index=True, use_container_width=True, key="tm_ed")

        if st.button("Process Race Entries & Update Cloud"):
            conn.update(worksheet="Schools", data=edited_schools)
            conn.update(worksheet="Teams", data=edited_teams)
            
            school_dict = dict(zip(edited_schools["Raw Name"], edited_schools["Cleaned Name"]))
            df['Cleaned School Name'] = df['School name'].replace(school_dict)
            team_dict = dict(zip(edited_teams["Raw Name"], edited_teams["Cleaned Name"]))
            df['Cleaned Team Name'] = df['Team name'].replace(team_dict) if 'Team name' in df.columns else ''
            
            output = io.BytesIO()
            audit_records = []
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Styles
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                header_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
                alt_fill = PatternFill(start_color="EAF1FB", end_color="EAF1FB", fill_type="solid")
                alert_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                title_font, header_font = Font(bold=True, size=16), Font(bold=True, color="FFFFFF")

                def apply_style(ws, col_count, is_alert=False):
                    for cell in ws[2]:
                        cell.fill, cell.font, cell.border = header_fill, header_font, thin_border
                    for i, row in enumerate(ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=col_count), start=1):
                        for cell in row:
                            cell.border = thin_border
                            if is_alert: cell.fill = alert_fill
                            elif i % 2 == 0: cell.fill = alt_fill
                    for col in ws.columns:
                        max_l = max((len(str(c.value)) for c in col), default=0)
                        ws.column_dimensions[col[0].column_letter].width = max_l + 4

                # Logical Separation
                is_kids_tkt = (df['Ticket'] == 'Pre-school to Year 9')
                missing_yr = (df['School year'].astype(str).str.strip() == '') | (df['School year'].astype(str).str.lower() == 'nan')
                
                discrepancy_df = df[is_kids_tkt & missing_yr].copy()
                adult_mask = (df['Ticket'] == 'Senior / Adult Race') | (df['School year'].isin(['Year 10', 'Year 11', 'Year 12', 'Year 13']))
                adult_df = df[adult_mask].copy().sort_values(by='Surname')
                kids_mask = is_kids_tkt & ~missing_yr & ~df.index.isin(adult_df.index)
                kids_df = df[kids_mask].copy()
                assigned = list(discrepancy_df.index) + list(adult_df.index) + list(kids_df.index)
                donation_df = df[~df.index.isin(assigned)].copy()

                if not discrepancy_df.empty:
                    discrepancy_df.to_excel(writer, sheet_name='Action Required', index=False, startrow=1)
                    apply_style(writer.sheets['Action Required'], len(discrepancy_df.columns), True)
                    audit_records.append({'Sort': -2, 'Category': '🚨 Action Required (Missing Year)', 'Count': len(discrepancy_df)})

                if not adult_df.empty:
                    cols = ['Race Number', 'Surname', 'Forename', 'Full name', 'Gender', 'Cleaned Team Name', 'School name', 'School year']
                    adult_df[cols].to_excel(writer, sheet_name='Senior Adult Race', index=False, startrow=1)
                    ws_a = writer.sheets['Senior Adult Race']
                    ws_a['A1'], ws_a['A1'].font = "Senior Adult Race", title_font
                    apply_style(ws_a, len(cols))
                    audit_records.append({'Sort': -1, 'Category': 'Senior Adult Race', 'Count': len(adult_df)})

                years = sorted([y for y in kids_df['School year'].unique() if y != ''], key=lambda x: YEAR_ORDER.get(x, 99))
                for y in years:
                    y_df = kids_df[kids_df['School year'] == y].sort_values(by='Surname')
                    y_cols = ['Race Number', 'Surname', 'Forename', 'Full name', 'Gender', 'Cleaned School Name']
                    y_df[y_cols].to_excel(writer, sheet_name=str(y)[:31], index=False, startrow=1)
                    ws_y = writer.sheets[str(y)[:31]]
                    ws_y['A1'], ws_y['A1'].font = f"Race Entries: {y}", title_font
                    apply_style(ws_y, 6)
                    audit_records.append({'Sort': YEAR_ORDER.get(y, 99), 'Category': f'Kids: {y}', 'Count': len(y_df)})

                audit_records.append({'Sort': 100, 'Category': 'Donations / Other', 'Count': len(donation_df)})
                audit_df = pd.DataFrame(audit_records).sort_values(by='Sort')
                total_out = audit_df['Count'].sum()
                audit_df[['Category', 'Count']].to_excel(writer, sheet_name='Audit Summary', index=False, startrow=3)
                ws_sum = writer.sheets['Audit Summary']
                ws_sum['A1'], ws_sum['A1'].font = "Data Reconciliation Report", title_font
                ws_sum['A2'] = f"Input: {total_input_rows} | Output: {total_out}"
                ws_sum['A2'].font = Font(color="008000" if total_input_rows == total_out else "FF0000", bold=True)
                apply_style(ws_sum, 2)

            st.success("Tring Race Pack Generated!")
            st.table(audit_df[['Category', 'Count']])
            st.download_button("📥 Download Race Pack", output.getvalue(), "Tring_Race_Pack.xlsx")

# --- TAB 2: TIMER RECONCILIATION ---
with tab_timer:
    st.header("Timer Results Reconciliation")
    st.info("Upload multiple timer CSVs to find the median consensus time and flag variances.")
    
    timer_files = st.file_uploader("Upload Timer CSVs", type=['csv'], accept_multiple_files=True, key="timer_upload")
    
    if timer_files:
        all_timers = []
        for t_file in timer_files:
            t_df = pd.read_csv(t_file, header=None)
            t_df = t_df[t_df[0].apply(lambda x: str(x).isdigit())]
            t_df[0] = t_df[0].astype(int) + 1
            t_df = t_df[[0, 2]].rename(columns={0: 'Position', 2: t_file.name}).set_index('Position')
            all_timers.append(t_df)
        
        master_timer = pd.concat(all_timers, axis=1)
        
        def to_sec(t):
            if pd.isna(t): return None
            h, m, s = map(int, str(t).split(':'))
            return h * 3600 + m * 60 + s

        sec_df = master_timer.map(to_sec)
        master_timer['Consensus'] = sec_df.median(axis=1).apply(lambda x: str(datetime.timedelta(seconds=int(x))).zfill(8) if pd.notna(x) else "")
        master_timer['Variance (Sec)'] = sec_df.max(axis=1) - sec_df.min(axis=1)
        
        st.subheader("Master Timing List")
        st.dataframe(master_timer.style.apply(lambda r: ['background-color: #ffcccc' if r['Variance (Sec)'] > 1 else '' for _ in r], axis=1), use_container_width=True)
        
        csv_out = io.StringIO()
        master_timer.to_csv(csv_out)
        st.download_button("📥 Download Consensus Times", csv_out.getvalue(), "Tring_Master_Times.csv")