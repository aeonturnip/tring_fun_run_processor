import streamlit as st
import pandas as pd
import io
from streamlit_gsheets import GSheetsConnection
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
import datetime

# 1. Page Configuration
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
tab_entry, tab_timer, tab_results = st.tabs(["📋 Race Entry Processor", "⏱️ Timer Reconciliation", "🏁 Final Results Marriage"])

# --- TAB 1: RACE ENTRY PROCESSOR ---
with tab_entry:
    st.header("Entry Formatter & Audit")
    uploaded_entries = st.sidebar.file_uploader("Upload Raw Tring Race CSV", type=['csv'], key="entry_upload")
    YEAR_ORDER = {'Pre-school': 0, 'Reception': 1, 'Year 1': 2, 'Year 2': 3, 'Year 3': 4, 'Year 4': 5, 'Year 6': 6, 'Year 7': 7, 'Year 8': 8, 'Year 9': 9, 'Year 10': 10}

    if uploaded_entries is not None:
        df = pd.read_csv(uploaded_entries).fillna('')
        total_input_rows = len(df)
        name_split = df['Full name'].str.split(' ', n=1, expand=True)
        df.insert(2, 'Forename', name_split[0].str.strip().str.title())
        df.insert(3, 'Surname', name_split[1].fillna('').str.strip().str.title())
        df.insert(0, 'Race Number', '')
        if 'School name' in df.columns:
            df['School name'] = df['School name'].astype(str).str.strip().str.title()
        if 'Team name' in df.columns:
            df['Team name'] = df['Team name'].astype(str).str.strip().str.title()

        school_mem_df = load_gsheet_memory("Schools")
        team_mem_df = load_gsheet_memory("Teams")
        school_memory = dict(zip(school_mem_df['Raw Name'], school_mem_df['Cleaned Name']))
        team_memory = dict(zip(team_mem_df['Raw Name'], team_mem_df['Cleaned Name']))

        col1, col2 = st.columns(2)
        with col1:
            unique_schools = [s for s in df['School name'].unique() if str(s).strip() != '' and str(s).lower() != 'nan']
            school_mapped = [school_memory.get(s, s) for s in unique_schools]
            school_editor_df = pd.DataFrame({"Raw Name": unique_schools, "Cleaned Name": school_mapped}).sort_values(by='Raw Name')
            edited_schools = st.data_editor(school_editor_df, hide_index=True, use_container_width=True, key="sch_ed")
        with col2:
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
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Styles (standard openpyxl)
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                header_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
                alt_fill = PatternFill(start_color="EAF1FB", end_color="EAF1FB", fill_type="solid")
                title_font, header_font = Font(bold=True, size=16), Font(bold=True, color="FFFFFF")

                def apply_style(ws, col_count):
                    for cell in ws[2]:
                        cell.fill, cell.font, cell.border = header_fill, header_font, thin_border
                    for i, row in enumerate(ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=col_count), start=1):
                        for cell in row:
                            cell.border = thin_border
                            if i % 2 == 0: cell.fill = alt_fill
                    for col in ws.columns:
                        max_l = max((len(str(c.value)) for c in col), default=0)
                        ws.column_dimensions[col[0].column_letter].width = max_l + 4

                is_kids_tkt = (df['Ticket'] == 'Pre-school to Year 9')
                adult_mask = (df['Ticket'] == 'Senior / Adult Race') | (df['School year'].isin(['Year 10', 'Year 11', 'Year 12', 'Year 13']))
                adult_df = df[adult_mask].copy().sort_values(by='Surname')
                kids_mask = is_kids_tkt & ~df.index.isin(adult_df.index)
                kids_df = df[kids_mask].copy()

                if not adult_df.empty:
                    cols = ['Race Number', 'Surname', 'Forename', 'Full name', 'Gender', 'Cleaned Team Name', 'School name', 'School year']
                    adult_df[cols].to_excel(writer, sheet_name='Senior Adult Race', index=False, startrow=1)
                    apply_style(writer.sheets['Senior Adult Race'], len(cols))

                years = sorted([y for y in kids_df['School year'].unique() if y != ''], key=lambda x: YEAR_ORDER.get(x, 99))
                for y in years:
                    y_df = kids_df[kids_df['School year'] == y].sort_values(by='Surname')
                    y_cols = ['Race Number', 'Surname', 'Forename', 'Full name', 'Gender', 'Cleaned School Name']
                    y_df[y_cols].to_excel(writer, sheet_name=str(y)[:31], index=False, startrow=1)
                    apply_style(writer.sheets[str(y)[:31]], 6)

            st.success("Tring Race Pack Generated!")
            st.download_button("📥 Download Race Pack", output.getvalue(), "Tring_Race_Pack.xlsx")
            st.session_state['reg_df'] = df

# --- TAB 2: TIMER RECONCILIATION ---
with tab_timer:
    st.header("Timer Results Reconciliation")
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
        sec_df = master_timer.map(lambda t: (int(t.split(':')[0])*3600 + int(t.split(':')[1])*60 + int(t.split(':')[2])) if pd.notna(t) else None)
        master_timer['Consensus Time'] = sec_df.median(axis=1).apply(lambda x: str(datetime.timedelta(seconds=int(x))).zfill(8) if pd.notna(x) else "")
        master_timer['Variance (Sec)'] = sec_df.max(axis=1) - sec_df.min(axis=1)
        st.dataframe(master_timer.style.apply(lambda r: ['background-color: #ffcccc' if r['Variance (Sec)'] > 1 else '' for _ in r], axis=1), use_container_width=True)
        st.session_state['master_timer'] = master_timer

# --- TAB 3: FINAL RESULTS MARRIAGE ---
with tab_results:
    st.header("Results Marriage & Scrutineer Check")
    scrut_files = st.file_uploader("Upload Scrutineer CSVs (Position, Race Number)", type=['csv'], accept_multiple_files=True, key="scrut_upload")
    
    if scrut_files:
        all_scruts = []
        for s_file in scrut_files:
            s_df = pd.read_csv(s_file).fillna('')
            s_df = s_df.rename(columns={s_df.columns[0]: 'Position', s_df.columns[1]: s_file.name}).set_index('Position')
            all_scruts.append(s_df)
        
        master_scrut = pd.concat(all_scruts, axis=1)
        master_scrut['Consensus Bib'] = master_scrut.apply(lambda row: row.iloc[0] if row.nunique() == 1 else "CONFLICT", axis=1)
        st.dataframe(master_scrut.style.apply(lambda r: ['background-color: #ffcccc' if r['Consensus Bib'] == "CONFLICT" else '' for _ in r], axis=1), use_container_width=True)

        if st.button("Generate Official Results"):
            if 'master_timer' in st.session_state and 'reg_df' in st.session_state:
                final = pd.merge(st.session_state['master_timer'][['Consensus Time']], master_scrut[['Consensus Bib']], left_index=True, right_index=True, how='left')
                reg = st.session_state['reg_df']
                reg['Race Number'] = reg['Race Number'].astype(str)
                final['Consensus Bib'] = final['Consensus Bib'].astype(str)
                results_complete = final.merge(reg, left_on='Consensus Bib', right_on='Race Number', how='left')
                
                # --- EXCEL OUTPUT FORMATTING ---
                output_res = io.BytesIO()
                with pd.ExcelWriter(output_res, engine='openpyxl') as res_writer:
                    # Filter for only Senior Adult Race (Extrapolated from 2025 format)
                    res_adult = results_complete[results_complete['Ticket'] == 'Senior / Adult Race'].copy()
                    res_adult['Position'] = range(1, len(res_adult) + 1)
                    
                    # Columns matching the Results - Adult Senior Race tab
                    final_cols = ['Position', 'Consensus Bib', 'Full name', 'Gender', 'Cleaned Team Name', 'School year', 'Consensus Time']
                    res_adult[final_cols].to_excel(res_writer, sheet_name='Results - Adult Senior Race', index=False, startrow=1)
                    
                    ws = res_writer.sheets['Results - Adult Senior Race']
                    ws['A1'] = "Official Results: Tring Fun Run - Senior Adult Race"
                    ws['A1'].font = Font(bold=True, size=14)
                    
                    # Apply Zebra Styling
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

                st.success("Official Results Formatted Successfully!")
                st.download_button("📥 Download Official Formatted Results", output_res.getvalue(), "Tring_Senior_Adult_Results.xlsx")