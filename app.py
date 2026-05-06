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

# --- CACHED LOADING LOGIC ---
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
                    late_df = conn.read(worksheet="LateEntries", ttl=0).fillna('')
                    updated = pd.concat([late_df, new_entry], ignore_index=True)
                    conn.update(worksheet="LateEntries", data=updated)
                    st.cache_data.clear()
                    st.success(f"Bib {bib} saved.")

    with col_pre:
        st.header("2. Pre-Reg & Race Pack Generation")
        csv_file = st.sidebar.file_uploader("Upload Registration CSV", type=['csv'], key="pre_reg_upload")
        
        if csv_file:
            df_raw = pd.read_csv(csv_file).fillna('')
            if 'Full name' in df_raw.columns:
                name_split = df_raw['Full name'].str.split(' ', n=1, expand=True)
                df_raw['Forename'] = name_split[0].str.strip().str.title()
                df_raw['Surname'] = name_split[1].fillna('').str.strip().str.title()

            for col in ['Race Number', 'Gender', 'Team name', 'School name', 'School year', 'Ticket']:
                if col not in df_raw.columns: df_raw[col] = ""

            new_schools = [s for s in df_raw['School name'].unique() if str(s).strip() != '']
            new_teams = [t for t in df_raw['Team name'].unique() if str(t).strip() != '']
            
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
            df_raw['Ticket'] = df_raw['Ticket'].str.strip()
            adult_mask = (df_raw['Ticket'] == 'Senior / Adult Race') | (df_raw['School year'].isin(['Year 10', 'Year 11', 'Year 12', 'Year 13']))
            pre_reg_adults = df_raw[adult_mask][['Forename', 'Surname', 'Gender', 'Team name', 'School year', 'Ticket']].copy()
            
            bib_allocs = load_gsheet_cached("BibAllocations")
            if not bib_allocs.empty:
                pre_reg_adults = pre_reg_adults.merge(bib_allocs[['Forename', 'Surname', 'Race Number']], on=['Forename', 'Surname'], how='left')
            else:
                pre_reg_adults['Race Number'] = ""
            
            ed_bibs = st.data_editor(pre_reg_adults.sort_values('Surname'), key="bib_ed", hide_index=True)

            if st.button("Process & Generate Pack"):
                conn.update(worksheet="Schools", data=ed_sch)
                conn.update(worksheet="Teams", data=ed_tm)
                conn.update(worksheet="BibAllocations", data=ed_bibs)
                
                sch_dict = dict(zip(ed_sch["Raw Name"], ed_sch["Cleaned Name"]))
                tm_dict = dict(zip(ed_tm["Raw Name"], ed_tm["Cleaned Name"]))
                df_export = df_raw.copy()
                df_export['School name'] = df_export['School name'].replace(sch_dict)
                df_export['Team name'] = df_export['Team name'].replace(tm_dict)
                
                if 'Race Number' in df_export.columns: df_export = df_export.drop(columns=['Race Number'])
                df_export = df_export.merge(ed_bibs[['Forename', 'Surname', 'Race Number']], on=['Forename', 'Surname'], how='left').fillna('')

                output = io.BytesIO()
                YEAR_ORDER = {'Pre-school': 0, 'Reception': 1, 'Year 1': 2, 'Year 2': 3, 'Year 3': 4, 'Year 4': 5, 'Year 5': 6, 'Year 6': 7, 'Year 7': 8, 'Year 8': 9, 'Year 9': 10, 'Year 10': 11}
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # [Formatting functions omitted for space, keep exactly as in your current file]
                    def apply_style(ws, col_count, sheet_display_name):
                        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=col_count)
                        title_cell = ws.cell(row=1, column=1); title_cell.value = f"Tring Fun Run 2026: {sheet_display_name}"
                        title_cell.font = Font(bold=True, size=16); title_cell.alignment = Alignment(horizontal='center', vertical='center')
                        ws.row_dimensions[1].height = 45
                        ws.row_dimensions[2].height = 25
                        for col_idx in range(1, col_count + 1):
                            cell = ws.cell(row=2, column=col_idx); cell.fill, cell.font, cell.border = PatternFill(start_color="333333", end_color="333333", fill_type="solid"), Font(bold=True, color="FFFFFF"), Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            header_text = str(cell.value)
                            if "School" in header_text: ws.column_dimensions[get_column_letter(col_idx)].width = 45
                            elif "Team" in header_text: ws.column_dimensions[get_column_letter(col_idx)].width = 35
                            elif "Forename" in header_text or "Surname" in header_text: ws.column_dimensions[get_column_letter(col_idx)].width = 25
                            else: ws.column_dimensions[get_column_letter(col_idx)].width = 18
                        for i, row in enumerate(ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=col_count), start=1):
                            for cell in row:
                                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                                if i % 2 == 0: cell.fill = PatternFill(start_color="EAF1FB", end_color="EAF1FB", fill_type="solid")
                                cell.alignment = Alignment(vertical='center')

                    # 1. Senior Adult Race
                    res_adult_final = df_export[adult_mask].copy().sort_values('Surname')
                    seniors_written = len(res_adult_final)
                    if not res_adult_final.empty:
                        res_adult_final[['Race Number', 'Surname', 'Forename', 'Gender', 'Team name', 'School name', 'School year']].to_excel(writer, sheet_name='Senior Adult Race', index=False, startrow=1)
                        apply_style(writer.sheets['Senior Adult Race'], 7, "Senior Adult Race")
                    
                    # 2. Kids Tabs
                    kids_mask = (df_export['Ticket'].str.strip() == 'Pre-school to Year 9') & ~df_export.index.isin(res_adult_final.index)
                    kids_df = df_export[kids_mask].copy()
                    kids_written = len(kids_df)
                    years = sorted([y for y in kids_df['School year'].unique() if str(y).strip() != ''], key=lambda x: YEAR_ORDER.get(x, 99))
                    for y in years:
                        y_df = kids_df[kids_df['School year'] == y].sort_values('Surname')
                        y_df[['Race Number', 'Surname', 'Forename', 'Gender', 'School name']].to_excel(writer, sheet_name=str(y)[:31], index=False, startrow=1)
                        apply_style(writer.sheets[str(y)[:31]], 5, y)
                
                # --- NEW RECONCILIATION SUMMARY ---
                st.success("✅ Race Pack Processed Successfully!")
                
                recon_data = {
                    "Category": ["Seniors / Adults", "Kids (Pre-school to Yr 9)", "Total Runners"],
                    "Count in CSV": [len(df_raw[adult_mask]), len(df_raw[kids_mask]), len(df_raw)],
                    "Count in Excel": [seniors_written, kids_written, seniors_written + kids_written]
                }
                st.subheader("🏁 Data Integrity Check")
                recon_df = pd.DataFrame(recon_data)
                
                # Highlight discrepancies
                def highlight_mismatch(s):
                    return ['background-color: #ffcccc' if s['Count in CSV'] != s['Count in Excel'] else 'background-color: #ccffcc' for _ in s]
                
                st.table(recon_df.style.apply(highlight_mismatch, axis=1))

                st.download_button("📥 Download Race Pack", output.getvalue(), "Tring_Race_Pack_2026.xlsx")
                st.session_state['processed_reg'] = df_export
                st.cache_data.clear()

# --- TABS 2, 3, & 4 REMAIN THE SAME ---
