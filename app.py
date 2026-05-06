import streamlit as st
import pandas as pd
import io
import datetime
from streamlit_gsheets import GSheetsConnection
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# 1. Page Configuration
st.set_page_config(page_title="Tring Fun Run Tools", layout="wide")
st.title("🏃‍♂️ Tring Fun Run: 2026 Registration & Results")

# 2. Establish Cloud Connection
conn = st.connection("gsheets", type=GSheetsConnection)

def load_gsheet(worksheet_name, ttl_val=None):
    """
    Safely loads a Google Sheet tab. If the tab is missing or empty, 
    it returns a DataFrame with the correct headers for that specific tab.
    """
    try:
        df = conn.read(worksheet=worksheet_name, ttl=ttl_val).fillna('')
        # If the sheet exists but has no columns/data, force the correct headers
        if df.empty:
            raise ValueError("Empty Sheet")
        return df
    except:
        # Define correct headers for each specific worksheet
        if worksheet_name in ["Schools", "Teams"]:
            return pd.DataFrame(columns=["Raw Name", "Cleaned Name"])
        elif worksheet_name == "SchoolRolls":
            return pd.DataFrame(columns=["School Name", "Infants Roll", "Juniors Roll"])
        elif worksheet_name == "BibAllocations":
             return pd.DataFrame(columns=["Forename", "Surname", "Gender", "Team name", "School year", "Race Number", "Ticket"])
        else:
            return pd.DataFrame(columns=["Forename", "Surname", "Gender", "School name", "Team name", "School year", "Race Number", "Ticket"])

# 3. Interface Tabs
tab_entry, tab_timer, tab_results, tab_stats = st.tabs([
    "📋 Registration & Bibs", 
    "⏱️ Timer Reconciliation", 
    "🏁 Final Results Marriage",
    "📊 Participation Stats"
])

# --- TAB 1: REGISTRATION & BIB ALLOCATION ---
with tab_entry:
    col_late, col_pre = st.columns([1, 2])
    
    with col_late:
        st.header("1. On-The-Day Entry")
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
                    tkt = "Senior / Adult Race" if yr == "Adult" or yr in ["Year 10", "Year 11", "Year 12", "Year 13"] else "Pre-school to Year 9"
                    new_entry = pd.DataFrame([{
                        "Forename": f_name.title(), "Surname": s_name.title(),
                        "Gender": gender, "School year": yr, "Team name": team.title(),
                        "School name": school.title(), "Race Number": str(bib).strip(), "Ticket": tkt
                    }])
                    existing = load_gsheet("LateEntries", ttl_val=0)
                    updated = pd.concat([existing, new_entry], ignore_index=True)
                    conn.update(worksheet="LateEntries", data=updated)
                    st.success(f"Bib {bib} saved to cloud.")

    with col_pre:
        st.header("2. Pre-Reg & Race Pack Generation")
        csv_file = st.sidebar.file_uploader("Upload Registration CSV", type=['csv'], key="pre_reg_upload")
        
        if csv_file:
            df = pd.read_csv(csv_file).fillna('')
            if 'Full name' in df.columns:
                name_split = df['Full name'].str.split(' ', n=1, expand=True)
                df['Forename'] = name_split[0].str.strip().str.title()
                df['Surname'] = name_split[1].fillna('').str.strip().str.title()

            for col in ['Race Number', 'Gender', 'Team name', 'School name', 'School year', 'Ticket']:
                if col not in df.columns: df[col] = ""

            # Memory Logic
            full_school_mem = load_gsheet("Schools", ttl_val=0)
            full_team_mem = load_gsheet("Teams", ttl_val=0)

            new_schools = [s for s in df['School name'].unique() if str(s).strip() != '']
            new_teams = [t for t in df['Team name'].unique() if str(t).strip() != '']

            schools_to_add = [s for s in new_schools if s not in full_school_mem['Raw Name'].values]
            teams_to_add = [t for t in new_teams if t not in full_team_mem['Raw Name'].values]

            if schools_to_add:
                full_school_mem = pd.concat([full_school_mem, pd.DataFrame({"Raw Name": schools_to_add, "Cleaned Name": schools_to_add})], ignore_index=True)
            if teams_to_add:
                full_team_mem = pd.concat([full_team_mem, pd.DataFrame({"Raw Name": teams_to_add, "Cleaned Name": teams_to_add})], ignore_index=True)

            c1, c2 = st.columns(2)
            with c1:
                st.subheader("School Mapping")
                edited_schools = st.data_editor(full_school_mem.sort_values("Raw Name"), key="sch_ed", hide_index=True)
            with c2:
                st.subheader("Team Mapping")
                edited_teams = st.data_editor(full_team_mem.sort_values("Raw Name"), key="tm_ed", hide_index=True)

            st.divider()
            st.subheader("3. Bib Assignment (Senior Adult Race)")
            df['Ticket'] = df['Ticket'].str.strip()
            adult_mask = (df['Ticket'] == 'Senior / Adult Race') | (df['School year'].isin(['Year 10', 'Year 11', 'Year 12', 'Year 13']))
            pre_reg_adults = df[adult_mask][['Forename', 'Surname', 'Gender', 'Team name', 'School year', 'Ticket']].copy()
            
            existing_bibs = load_gsheet("BibAllocations", ttl_val=0)
            if not existing_bibs.empty:
                # Merge and preserve Ticket/Gender/Team info
                pre_reg_adults = pre_reg_adults.merge(existing_bibs[['Forename', 'Surname', 'Race Number']], on=['Forename', 'Surname'], how='left')
            else:
                pre_reg_adults['Race Number'] = ""
            
            edited_bibs = st.data_editor(pre_reg_adults.sort_values('Surname'), key="bib_ed", hide_index=True)

            if st.button("Process Race Entries & Generate Pack"):
                conn.update(worksheet="Schools", data=edited_schools)
                conn.update(worksheet="Teams", data=edited_teams)
                # Store full runner details so Results tab can find Ticket type
                conn.update(worksheet="BibAllocations", data=edited_bibs)
                
                school_dict = dict(zip(edited_schools["Raw Name"], edited_schools["Cleaned Name"]))
                team_dict = dict(zip(edited_teams["Raw Name"], edited_teams["Cleaned Name"]))
                
                df_export = df.copy()
                df_export['School name'] = df_export['School name'].replace(school_dict)
                df_export['Team name'] = df_export['Team name'].replace(team_dict)

                output = io.BytesIO()
                YEAR_ORDER = {'Pre-school': 0, 'Reception': 1, 'Year 1': 2, 'Year 2': 3, 'Year 3': 4, 'Year 4': 5, 'Year 5': 6, 'Year 6': 7, 'Year 7': 8, 'Year 8': 9, 'Year 9': 10, 'Year 10': 11}
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    header_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
                    alt_fill = PatternFill(start_color="EAF1FB", end_color="EAF1FB", fill_type="solid")
                    header_font = Font(bold=True, color="FFFFFF")
                    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

                    def apply_style(ws, col_count, sheet_display_name):
                        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=col_count)
                        title_cell = ws.cell(row=1, column=1)
                        title_cell.value = f"Tring Fun Run 2026: {sheet_display_name}"
                        title_cell.font = Font(bold=True, size=16)
                        title_cell.alignment = Alignment(horizontal='center', vertical='center')
                        ws.row_dimensions[1].height = 45

                        ws.row_dimensions[2].height = 25
                        for col_idx in range(1, col_count + 1):
                            cell = ws.cell(row=2, column=col_idx)
                            cell.fill, cell.font, cell.border = header_fill, header_font, border
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            header_text = str(cell.value)
                            if "School" in header_text: ws.column_dimensions[get_column_letter(col_idx)].width = 45
                            elif "Team" in header_text: ws.column_dimensions[get_column_letter(col_idx)].width = 35
                            elif "Forename" in header_text or "Surname" in header_text: ws.column_dimensions[get_column_letter(col_idx)].width = 25
                            else: ws.column_dimensions[get_column_letter(col_idx)].width = 18
                        
                        for i, row in enumerate(ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=col_count), start=1):
                            for cell in row:
                                cell.border = border
                                if i % 2 == 0: cell.fill = alt_fill
                                cell.alignment = Alignment(vertical='center')

                    res_adult_final = df_export[adult_mask].copy().sort_values('Surname')
                    s_cols = ['Race Number', 'Surname', 'Forename', 'Gender', 'Team name', 'School name', 'School year']
                    res_adult_final[s_cols].to_excel(writer, sheet_name='Senior Adult Race', index=False, startrow=1)
                    apply_style(writer.sheets['Senior Adult Race'], len(s_cols), "Senior Adult Race")

                    kids_mask = (df_export['Ticket'].str.strip() == 'Pre-school to Year 9') & ~df_export.index.isin(res_adult_final.index)
                    kids_df = df_export[kids_mask].copy()
                    years = sorted([y for y in kids_df['School year'].unique() if str(y).strip() != ''], key=lambda x: YEAR_ORDER.get(x, 99))
                    for y in years:
                        y_df = kids_df[kids_df['School year'] == y].sort_values('Surname')
                        k_cols = ['Race Number', 'Surname', 'Forename', 'Gender', 'School name']
                        y_df[k_cols].to_excel(writer, sheet_name=str(y)[:31], index=False, startrow=1)
                        apply_style(writer.sheets[str(y)[:31]], len(k_cols), y)

                st.success("Tring Race Pack Created!")
                st.download_button("📥 Download Race Pack", output.getvalue(), "Tring_Race_Pack_2026.xlsx")
                st.session_state['processed_reg'] = df_export

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
        
        def to_sec(t):
            if pd.isna(t): return None
            parts = str(t).split(':'); return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
        
        sec_df = master_timer.map(to_sec)
        master_timer['Consensus Time'] = sec_df.median(axis=1).apply(lambda x: str(datetime.timedelta(seconds=int(x))).zfill(8) if pd.notna(x) else "")
        master_timer['Variance (Sec)'] = sec_df.max(axis=1) - sec_df.min(axis=1)
        st.dataframe(master_timer.style.apply(lambda r: ['background-color: #ffcccc' if r['Variance (Sec)'] > 1 else '' for _ in r], axis=1), use_container_width=True)
        st.session_state['master_timer'] = master_timer

# --- TAB 3: FINAL RESULTS MARRIAGE ---
with tab_results:
    st.header("Official Results Generation")
    scrut_files = st.file_uploader("Upload Scrutineer CSVs", type=['csv'], accept_multiple_files=True, key="scrut_upload")
    if scrut_files:
        all_scruts = []
        for s_file in scrut_files:
            s_df = pd.read_csv(s_file).fillna('')
            s_df = s_df.rename(columns={s_df.columns[0]: 'Position', s_df.columns[1]: s_file.name}).set_index('Position')
            all_scruts.append(s_df)
        master_scrut = pd.concat(all_scruts, axis=1)
        master_scrut['Consensus Bib'] = master_scrut.apply(lambda row: str(row.iloc[0]).strip() if row.nunique() == 1 else "CONFLICT", axis=1)
        st.dataframe(master_scrut.style.apply(lambda r: ['background-color: #ffcccc' if r['Consensus Bib'] == "CONFLICT" else '' for _ in r], axis=1), use_container_width=True)

        if 'master_timer' in st.session_state:
            # Refresh data from cloud
            late_entries = load_gsheet("LateEntries", ttl_val=0)
            pre_reg_assigned = load_gsheet("BibAllocations", ttl_val=0)
            master_runners = pd.concat([late_entries, pre_reg_assigned], ignore_index=True)
            master_runners['Race Number'] = master_runners['Race Number'].astype(str).str.strip()

            # Marriage
            final_base = pd.merge(st.session_state['master_timer'][['Consensus Time']], master_scrut[['Consensus Bib']], left_index=True, right_index=True, how='left')
            results_complete = final_base.merge(master_runners, left_on='Consensus Bib', right_on='Race Number', how='left')
            
            # Senior Filter
            res_adult = results_complete[results_complete['Ticket'] == 'Senior / Adult Race'].copy()
            res_adult.insert(0, 'Position', range(1, len(res_adult) + 1))
            
            st.subheader("Live Results Preview (Seniors Only)")
            if not res_adult.empty:
                st.dataframe(res_adult[['Position', 'Race Number', 'Forename', 'Surname', 'Consensus Time']], use_container_width=True, hide_index=True)
                
                if st.button("Generate Senior Results Excel"):
                    output_res = io.BytesIO()
                    with pd.ExcelWriter(output_res, engine='openpyxl') as res_writer:
                        final_cols = ['Position', 'Race Number', 'Forename', 'Surname', 'Gender', 'Team name', 'School year', 'Consensus Time']
                        res_adult[final_cols].to_excel(res_writer, sheet_name='Results - Adult Senior Race', index=False, startrow=1)
                        ws = res_writer.sheets['Results - Adult Senior Race']
                        ws['A1'] = "Tring Fun Run 2026: Official Senior Results"; ws['A1'].font = Font(bold=True, size=14)
                    st.download_button("📥 Download Official Results", output_res.getvalue(), "Tring_Senior_Results_2026.xlsx")
            else:
                st.warning("No matched Senior runners found. Ensure bibs in scrutineer files are recorded in Tab 1 under the Senior / Adult ticket.")

# --- TAB 4: PARTICIPATION STATS ---
with tab_stats:
    st.header("📊 Live Participation Leaderboard")
    late_df = load_gsheet("LateEntries", ttl_val=0); pre_reg_df = st.session_state.get('processed_reg', pd.DataFrame())
    if not late_df.empty or not pre_reg_df.empty:
        df_all = pd.concat([late_df, pre_reg_df], ignore_index=True)
        df_all['School year'] = df_all['School year'].astype(str).str.strip().str.title()
        infant_yrs = ['Reception', 'Year 1', 'Year 2']; junior_yrs = ['Year 3', 'Year 4', 'Year 5', 'Year 6']
        df_all['Tier'] = 'Other'; df_all.loc[df_all['School year'].isin(infant_yrs), 'Tier'] = 'Infants'; df_all.loc[df_all['School year'].isin(junior_yrs), 'Tier'] = 'Juniors'
        
        col_sch, col_tm = st.columns(2)
        with col_sch:
            st.subheader("🏫 School Participation"); rolls_df = load_gsheet("SchoolRolls", ttl_val=0)
            if not rolls_df.empty:
                tier_counts = df_all[df_all['Tier'].isin(['Infants', 'Juniors'])].groupby(['School name', 'Tier']).size().unstack(fill_value=0).reset_index()
                rolls_df['Infants Roll'] = pd.to_numeric(rolls_df['Infants Roll'], errors='coerce').fillna(0); rolls_df['Juniors Roll'] = pd.to_numeric(rolls_df['Juniors Roll'], errors='coerce').fillna(0)
                sch_stats = rolls_df.merge(tier_counts, left_on='School Name', right_on='School name', how='left').fillna(0)
                sch_stats['Infant %'] = (sch_stats['Infants'] / sch_stats['Infants Roll'] * 100).round(1).replace([float('inf')], 0).fillna(0)
                sch_stats['Junior %'] = (sch_stats['Juniors'] / sch_stats['Juniors Roll'] * 100).round(1).replace([float('inf')], 0).fillna(0)
                st.dataframe(sch_stats[['School Name', 'Infants', 'Infants Roll', 'Infant %', 'Juniors', 'Juniors Roll', 'Junior %']].sort_values('Junior %', ascending=False), use_container_width=True, hide_index=True)
        with col_tm:
            st.subheader("🏃‍♂️ Team Entry Totals"); team_counts = df_all[df_all['Team name'] != '']['Team name'].value_counts().reset_index()
            team_counts.columns = ['Team Name', 'Entrants']
            st.dataframe(team_counts.sort_values('Entrants', ascending=False), hide_index=True, use_container_width=True)
    else:
        st.info("No participants found yet.")
