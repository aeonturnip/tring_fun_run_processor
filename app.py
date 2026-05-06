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

# --- CACHED LOADING LOGIC (Quota Protection) ---
@st.cache_data(ttl=600)
def load_gsheet_cached(worksheet_name):
    try:
        df = conn.read(worksheet=worksheet_name, ttl=0).fillna('')
        if df.empty: raise ValueError
        return df
    except:
        if worksheet_name in ["Schools", "Teams"]:
            return pd.DataFrame(columns=["Raw Name", "Cleaned Name"])
        elif worksheet_name == "SchoolRolls":
            return pd.DataFrame(columns=["School Name", "Infants Roll", "Juniors Roll"])
        return pd.DataFrame(columns=["Forename", "Surname", "Gender", "School name", "Team name", "School year", "Race Number", "Ticket"])

# Sidebar Sync
if st.sidebar.button("🔄 Sync with Google Sheets"):
    st.cache_data.clear()
    st.toast("Fetching fresh data from cloud...")

# Pre-load data once
late_entries_master = load_gsheet_cached("LateEntries")
bib_allocs_master = load_gsheet_cached("BibAllocations")
school_mem_master = load_gsheet_cached("Schools")
team_mem_master = load_gsheet_cached("Teams")
rolls_master = load_gsheet_cached("SchoolRolls")

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
            yr = st.selectbox("Year Group", ["Adult", "Pre-school", "Reception"] + [f"Year {i}" for i in range(1, 14)])
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
                    updated = pd.concat([late_entries_master, new_entry], ignore_index=True)
                    conn.update(worksheet="LateEntries", data=updated)
                    st.cache_data.clear()
                    st.success(f"Bib {bib} saved.")

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

            new_schools = [s for s in df['School name'].unique() if str(s).strip() != '']
            new_teams = [t for t in df['Team name'].unique() if str(t).strip() != '']
            
            sch_to_add = [s for s in new_schools if s not in school_mem_master['Raw Name'].values]
            tm_to_add = [t for t in new_teams if t not in team_mem_master['Raw Name'].values]

            if sch_to_add:
                school_mem_master = pd.concat([school_mem_master, pd.DataFrame({"Raw Name": sch_to_add, "Cleaned Name": sch_to_add})], ignore_index=True)
            if tm_to_add:
                team_mem_master = pd.concat([team_mem_master, pd.DataFrame({"Raw Name": tm_to_add, "Cleaned Name": tm_to_add})], ignore_index=True)

            c1, c2 = st.columns(2)
            with c1: ed_sch = st.data_editor(school_mem_master.sort_values("Raw Name"), key="sch_ed", hide_index=True)
            with c2: ed_tm = st.data_editor(team_mem_master.sort_values("Raw Name"), key="tm_ed", hide_index=True)

            st.divider()
            st.subheader("3. Bib Assignment (Senior Adult Race)")
            df['Ticket'] = df['Ticket'].str.strip()
            adult_mask = (df['Ticket'] == 'Senior / Adult Race') | (df['School year'].isin(['Year 10', 'Year 11', 'Year 12', 'Year 13']))
            pre_reg_adults = df[adult_mask][['Forename', 'Surname', 'Gender', 'Team name', 'School year', 'Ticket']].copy()
            
            if not bib_allocs_master.empty:
                pre_reg_adults = pre_reg_adults.merge(bib_allocs_master[['Forename', 'Surname', 'Race Number']], on=['Forename', 'Surname'], how='left')
            else:
                pre_reg_adults['Race Number'] = ""
            
            ed_bibs = st.data_editor(pre_reg_adults.sort_values('Surname'), key="bib_ed", hide_index=True)

            if st.button("Process & Generate Pack"):
                conn.update(worksheet="Schools", data=ed_sch)
                conn.update(worksheet="Teams", data=ed_tm)
                conn.update(worksheet="BibAllocations", data=ed_bibs)
                
                sch_dict = dict(zip(ed_sch["Raw Name"], ed_sch["Cleaned Name"]))
                tm_dict = dict(zip(ed_tm["Raw Name"], ed_tm["Cleaned Name"]))
                df_export = df.copy()
                df_export['School name'] = df_export['School name'].replace(sch_dict)
                df_export['Team name'] = df_export['Team name'].replace(tm_dict)

                output = io.BytesIO()
                YEAR_ORDER = {'Pre-school': 0, 'Reception': 1, 'Year 1': 2, 'Year 2': 3, 'Year 3': 4, 'Year 4': 5, 'Year 5': 6, 'Year 6': 7, 'Year 7': 8, 'Year 8': 9, 'Year 9': 10, 'Year 10': 11}
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    header_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
                    alt_fill = PatternFill(start_color="EAF1FB", end_color="EAF1FB", fill_type="solid")
                    header_font = Font(bold=True, color="FFFFFF")
                    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

                    def apply_style(ws, col_count, sheet_display_name):
                        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=col_count)
                        title_cell = ws.cell(row=1, column=1); title_cell.value = f"Tring Fun Run 2026: {sheet_display_name}"
                        title_cell.font = Font(bold=True, size=16); title_cell.alignment = Alignment(horizontal='center', vertical='center')
                        ws.row_dimensions[1].height = 45
                        ws.row_dimensions[2].height = 25
                        for col_idx in range(1, col_count + 1):
                            cell = ws.cell(row=2, column=col_idx); cell.fill, cell.font, cell.border = header_fill, header_font, border
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

                    # Sheets Logic
                    res_adult_final = df_export[adult_mask].copy().sort_values('Surname')
                    s_cols = ['Race Number', 'Surname', 'Forename', 'Gender', 'Team name', 'School name', 'School year']
                    
                    sheets_created = 0
                    if not res_adult_final.empty:
                        res_adult_final[s_cols].to_excel(writer, sheet_name='Senior Adult Race', index=False, startrow=1)
                        apply_style(writer.sheets['Senior Adult Race'], len(s_cols), "Senior Adult Race")
                        sheets_created += 1

                    kids_mask = (df_export['Ticket'].str.strip() == 'Pre-school to Year 9') & ~df_export.index.isin(res_adult_final.index)
                    kids_df = df_export[kids_mask].copy()
                    years = sorted([y for y in kids_df['School year'].unique() if str(y).strip() != ''], key=lambda x: YEAR_ORDER.get(x, 99))
                    for y in years:
                        y_df = kids_df[kids_df['School year'] == y].sort_values('Surname')
                        k_cols = ['Race Number', 'Surname', 'Forename', 'Gender', 'School name']
                        sheet_name = str(y)[:31]
                        y_df[k_cols].to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
                        apply_style(writer.sheets[sheet_name], len(k_cols), y)
                        sheets_created += 1
                    
                    if sheets_created == 0:
                        pd.DataFrame([["No runners found with the selected filters"]]).to_excel(writer, sheet_name="Empty Template")

                st.success("Race Pack Ready!")
                st.download_button("📥 Download", output.getvalue(), "Tring_Race_Pack.xlsx")
                st.session_state['processed_reg'] = df_export
                st.cache_data.clear()

# --- TAB 2: TIMER ---
with tab_timer:
    st.header("Timer Reconciliation")
    t_files = st.file_uploader("Upload Timer CSVs", type=['csv'], accept_multiple_files=True, key="timer_u")
    if t_files:
        all_timers = []
        for f in t_files:
            t_df = pd.read_csv(f, header=None)
            t_df = t_df[t_df[0].apply(lambda x: str(x).isdigit())]
            t_df[0] = t_df[0].astype(int) + 1
            t_df = t_df[[0, 2]].rename(columns={0: 'Position', 2: f.name}).set_index('Position')
            all_timers.append(t_df)
        master_t = pd.concat(all_timers, axis=1)
        def to_s(t):
            if pd.isna(t): return None
            p = str(t).split(':'); return int(p[0])*3600 + int(p[1])*60 + int(p[2])
        sec_df = master_t.map(to_s)
        master_t['Consensus Time'] = sec_df.median(axis=1).apply(lambda x: str(datetime.timedelta(seconds=int(x))).zfill(8) if pd.notna(x) else "")
        master_t['Variance (Sec)'] = sec_df.max(axis=1) - sec_df.min(axis=1)
        st.dataframe(master_t.style.apply(lambda r: ['background-color: #ffcccc' if r['Variance (Sec)'] > 1 else '' for _ in r], axis=1), use_container_width=True)
        st.session_state['master_timer'] = master_t

# --- TAB 3: RESULTS ---
with tab_results:
    st.header("Results Marriage")
    s_files = st.file_uploader("Upload Scrutineer CSVs", type=['csv'], accept_multiple_files=True, key="scrut_u")
    if s_files:
        all_scruts = []
        for f in s_files:
            s_df = pd.read_csv(f).fillna('')
            s_df = s_df.rename(columns={s_df.columns[0]: 'Position', s_df.columns[1]: f.name}).set_index('Position')
            all_scruts.append(s_df)
        master_s = pd.concat(all_scruts, axis=1)
        master_s['Consensus Bib'] = master_s.apply(lambda row: str(row.iloc[0]).strip() if row.nunique() == 1 else "CONFLICT", axis=1)
        st.dataframe(master_s.style.apply(lambda r: ['background-color: #ffcccc' if r['Consensus Bib'] == "CONFLICT" else '' for _ in r], axis=1), use_container_width=True)

        if 'master_timer' in st.session_state:
            runners = pd.concat([late_entries_master, bib_allocs_master], ignore_index=True)
            runners['Race Number'] = runners['Race Number'].astype(str).str.strip()
            final_b = pd.merge(st.session_state['master_timer'][['Consensus Time']], master_s[['Consensus Bib']], left_index=True, right_index=True, how='left')
            res_comp = final_b.merge(runners, left_on='Consensus Bib', right_on='Race Number', how='left')
            res_adult = res_comp[res_comp['Ticket'] == 'Senior / Adult Race'].copy()
            res_adult.insert(0, 'Position', range(1, len(res_adult) + 1))
            
            st.subheader("Results Preview (Seniors)")
            if not res_adult.empty:
                st.dataframe(res_adult[['Position', 'Race Number', 'Forename', 'Surname', 'Consensus Time']], use_container_width=True, hide_index=True)
                if st.button("Download Official Results"):
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine='openpyxl') as wr:
                        res_adult[['Position', 'Race Number', 'Forename', 'Surname', 'Gender', 'Team name', 'School year', 'Consensus Time']].to_excel(wr, sheet_name='Senior Results', index=False, startrow=1)
                        wr.sheets['Senior Results']['A1'] = "Tring Fun Run 2026: Official Results"
                    st.download_button("📥 Download Excel", out.getvalue(), "Tring_Results.xlsx")
            else:
                st.warning("No Senior runners matched.")

# --- TAB 4: STATS ---
with tab_stats:
    st.header("📊 Stats")
    pr = st.session_state.get('processed_reg', pd.DataFrame())
    df_all = pd.concat([late_entries_master, pr], ignore_index=True)
    if not df_all.empty:
        df_all['School year'] = df_all['School year'].astype(str).str.strip().str.title()
        infant = ['Reception', 'Year 1', 'Year 2']; junior = ['Year 3', 'Year 4', 'Year 5', 'Year 6']
        df_all['Tier'] = 'Other'; df_all.loc[df_all['School year'].isin(infant), 'Tier'] = 'Infants'; df_all.loc[df_all['School year'].isin(junior), 'Tier'] = 'Juniors'
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Schools")
            if not rolls_master.empty:
                cnts = df_all[df_all['Tier'].isin(['Infants', 'Juniors'])].groupby(['School name', 'Tier']).size().unstack(fill_value=0).reset_index()
                stat = rolls_master.merge(cnts, left_on='School Name', right_on='School name', how='left').fillna(0)
                stat['Inf %'] = (stat['Infants'] / pd.to_numeric(stat['Infants Roll'], errors='coerce') * 100).round(1).fillna(0)
                stat['Jun %'] = (stat['Juniors'] / pd.to_numeric(stat['Juniors Roll'], errors='coerce') * 100).round(1).fillna(0)
                st.dataframe(stat[['School Name', 'Infants', 'Inf %', 'Juniors', 'Jun %']].sort_values('Jun %', ascending=False), use_container_width=True, hide_index=True)
        with c2:
            st.subheader("Teams")
            tm_cnts = df_all[df_all['Team name'] != '']['Team name'].value_counts().reset_index()
            st.dataframe(tm_cnts, use_container_width=True, hide_index=True)
