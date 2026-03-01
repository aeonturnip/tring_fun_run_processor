import streamlit as st
import pandas as pd
import io
from streamlit_gsheets import GSheetsConnection
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
import datetime

# 1. Page Configuration
st.set_page_config(page_title="Tring Fun Run Tools", layout="wide")
st.title("🏃‍♂️ Tring Fun Run: Registration & Results Dashboard")

# 2. Establish Cloud Connection
conn = st.connection("gsheets", type=GSheetsConnection)

def load_gsheet(worksheet_name):
    try:
        return conn.read(worksheet=worksheet_name).fillna('')
    except:
        return pd.DataFrame()

# 3. Define Main Interface Tabs
tab_entry, tab_timer, tab_results = st.tabs([
    "📋 Registration & Bibs", 
    "⏱️ Timer Reconciliation", 
    "🏁 Final Results Marriage"
])

# --- TAB 1: REGISTRATION & BIB ALLOCATION ---
with tab_entry:
    col_late, col_pre = st.columns([1, 2])
    
    with col_late:
        st.header("1. On-The-Day Entry")
        st.info("Record new registrations and assign bibs.")
        with st.form("otd_form", clear_on_submit=True):
            f_name = st.text_input("Forename")
            s_name = st.text_input("Surname")
            gender = st.selectbox("Gender", ["Male", "Female", "Other"])
            year_opts = ["Adult", "Pre-school", "Reception"] + [f"Year {i}" for i in range(1, 14)]
            yr = st.selectbox("Year Group", year_opts)
            team = st.text_input("Team (Optional)")
            school = st.text_input("School (Optional)")
            bib = st.text_input("Assigned Bib Number")
            
            if st.form_submit_button("Submit & Save Late Entry"):
                if f_name and s_name and bib:
                    otd_data = pd.DataFrame([{
                        "Forename": f_name.title(), "Surname": s_name.title(),
                        "Gender": gender, "School year": yr, "Team name": team,
                        "School name": school, "Race Number": bib, "Ticket": "Senior / Adult Race" if yr == "Adult" or yr in ["Year 10", "Year 11", "Year 12", "Year 13"] else "Pre-school to Year 9"
                    }])
                    existing = load_gsheet("LateEntries")
                    updated = pd.concat([existing, otd_data], ignore_index=True)
                    conn.update(worksheet="LateEntries", data=updated)
                    st.success(f"Bib {bib} assigned to {f_name} {s_name}")
                else:
                    st.error("Name and Bib Number are required.")

    with col_pre:
        st.header("2. Pre-Registered Bib Assignment")
        csv_file = st.sidebar.file_uploader("Upload Registration CSV", type=['csv'], key="pre_reg_upload")
        
        if csv_file:
            df = pd.read_csv(csv_file).fillna('')
            name_split = df['Full name'].str.split(' ', n=1, expand=True)
            df['Forename'] = name_split[0].str.strip().str.title()
            df['Surname'] = name_split[1].fillna('').str.strip().str.title()
            
            # Identify Senior Race participants
            adult_mask = (df['Ticket'] == 'Senior / Adult Race') | (df['School year'].isin(['Year 10', 'Year 11', 'Year 12', 'Year 13']))
            pre_reg_adults = df[adult_mask][['Forename', 'Surname', 'Gender', 'Team name', 'School year', 'Ticket']].copy()
            
            existing_bibs = load_gsheet("BibAllocations")
            if not existing_bibs.empty:
                # Merge current assignments with the registration list
                pre_reg_adults = pre_reg_adults.merge(existing_bibs[['Forename', 'Surname', 'Race Number']], on=['Forename', 'Surname'], how='left')
            else:
                pre_reg_adults['Race Number'] = ""
            
            st.write("Assign bib numbers to pre-registered runners:")
            edited_bibs = st.data_editor(
                pre_reg_adults.sort_values('Surname'),
                hide_index=True, use_container_width=True,
                column_config={"Race Number": st.column_config.TextColumn("Bib Number")}
            )
            
            if st.button("Save Assignments to Cloud"):
                conn.update(worksheet="BibAllocations", data=edited_bibs)
                st.success("Cloud storage updated with latest bib assignments.")
                st.session_state['processed_reg'] = edited_bibs

# --- TAB 2: TIMER RECONCILIATION ---
with tab_timer:
    st.header("Timer Results Reconciliation")
    timer_files = st.file_uploader("Upload parkrun Timer CSVs", type=['csv'], accept_multiple_files=True)
    
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
            parts = str(t).split(':')
            return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])

        sec_df = master_timer.map(to_sec)
        master_timer['Consensus Time'] = sec_df.median(axis=1).apply(lambda x: str(datetime.timedelta(seconds=int(x))).zfill(8) if pd.notna(x) else "")
        master_timer['Variance (Sec)'] = sec_df.max(axis=1) - sec_df.min(axis=1)
        
        st.subheader("Master Timing List")
        st.dataframe(master_timer.style.apply(lambda r: ['background-color: #ffcccc' if r['Variance (Sec)'] > 1 else '' for _ in r], axis=1), use_container_width=True)
        st.session_state['master_timer'] = master_timer

# --- TAB 3: FINAL RESULTS MARRIAGE ---
with tab_results:
    st.header("Official Results Generation")
    scrut_files = st.file_uploader("Upload Scrutineer CSVs (Position, Race Number)", type=['csv'], accept_multiple_files=True)
    
    if scrut_files:
        all_scruts = []
        for s_file in scrut_files:
            s_df = pd.read_csv(s_file).fillna('')
            s_df = s_df.rename(columns={s_df.columns[0]: 'Position', s_df.columns[1]: s_file.name}).set_index('Position')
            all_scruts.append(s_df)
        
        master_scrut = pd.concat(all_scruts, axis=1)
        master_scrut['Consensus Bib'] = master_scrut.apply(lambda row: row.iloc[0] if row.nunique() == 1 else "CONFLICT", axis=1)
        
        st.subheader("1. Bib Reconciliation")
        st.dataframe(master_scrut.style.apply(lambda r: ['background-color: #ffcccc' if r['Consensus Bib'] == "CONFLICT" else '' for _ in r], axis=1), use_container_width=True)

        if st.button("Generate Formatted Results"):
            if 'master_timer' in st.session_state:
                # Build Master Runner List from Pre-reg and Late Entries
                late_entries = load_gsheet("LateEntries")
                pre_reg_assigned = load_gsheet("BibAllocations")
                master_runners = pd.concat([late_entries, pre_reg_assigned], ignore_index=True)
                master_runners['Race Number'] = master_runners['Race Number'].astype(str).str.strip()

                # Merge Timer + Scrutineer
                final_base = pd.merge(st.session_state['master_timer'][['Consensus Time']], master_scrut[['Consensus Bib']], left_index=True, right_index=True, how='left')
                final_base['Consensus Bib'] = final_base['Consensus Bib'].astype(str).str.strip()
                
                # Merge with Runner Profiles
                results_complete = final_base.merge(master_runners, left_on='Consensus Bib', right_on='Race Number', how='left')
                
                # --- EXCEL OUTPUT ---
                output_res = io.BytesIO()
                with pd.ExcelWriter(output_res, engine='openpyxl') as res_writer:
                    # Filter for Adult/Senior race finishers
                    res_adult = results_complete[results_complete['Ticket'] == 'Senior / Adult Race'].copy()
                    res_adult.insert(0, 'Results Position', range(1, len(res_adult) + 1))
                    
                    final_cols = ['Results Position', 'Race Number', 'Forename', 'Surname', 'Gender', 'Team name', 'School year', 'Consensus Time']
                    res_adult[final_cols].to_excel(res_writer, sheet_name='Results - Adult Senior Race', index=False, startrow=1)
                    
                    ws = res_writer.sheets['Results - Adult Senior Race']
                    ws['A1'] = "Official Results: Tring Fun Run - Adult Senior Race"
                    ws['A1'].font = Font(bold=True, size=14)
                    
                    # Formatting
                    header_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
                    alt_fill = PatternFill(start_color="EAF1FB", end_color="EAF1FB", fill_type="solid")
                    header_font = Font(bold=True, color="FFFFFF")
                    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                    
                    for cell in ws[2]:
                        cell.fill, cell.font, cell.border = header_fill, header_font, border
                    for i, row in enumerate(ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=len(final_cols)), start=1):
                        for cell in row:
                            cell.border = border
                            if i % 2 == 0: cell.fill = alt_fill
                    for col in ws.columns:
                        ws.column_dimensions[col[0].column_letter].width = 20

                st.success("Official Results generated successfully.")
                st.download_button("📥 Download Official Senior Results", output_res.getvalue(), "Tring_Senior_Results.xlsx")
            else:
                st.error("Please reconcile timers in Tab 2 before generating results.")