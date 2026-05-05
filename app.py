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
    try:
        return conn.read(worksheet=worksheet_name, ttl=ttl_val).fillna('')
    except:
        return pd.DataFrame(columns=["Raw Name", "Cleaned Name", "Race Number"])

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
                        "School name": school.title(), "Race Number": bib, "Ticket": tkt
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
                pre_reg_adults = pre_reg_adults.merge(existing_bibs[['Forename', 'Surname', 'Race Number']], on=['Forename', 'Surname'], how='left')
            else:
                pre_reg_adults['Race Number'] = ""
            
            edited_bibs = st.data_editor(pre_reg_adults.sort_values('Surname'), key="bib_ed", hide_index=True)

            if st.button("Process Race Entries & Generate Pack"):
                conn.update(worksheet="Schools", data=edited_schools)
                conn.update(worksheet="Teams", data=edited_teams)
                conn.update(worksheet="BibAllocations", data=edited_bibs[['Forename', 'Surname', 'Race Number']])
                
                school_dict = dict(zip(edited_schools["Raw Name"], edited_schools["Cleaned Name"]))
                team_dict = dict(zip(edited_teams["Raw Name"], edited_teams["Cleaned Name"]))
                
                df_export = df.copy()
                df_export['School name'] = df_export['School name'].replace(school_dict)
                df_export['Team name'] = df_export['Team name'].replace(team_dict)

                output = io.BytesIO()
                # Comprehensive Year Order for Tab Sorting
                YEAR_ORDER = {
                    'Pre-school': 0, 'Reception': 1, 'Year 1': 2, 'Year 2': 3, 'Year 3': 4,
                    'Year 4': 5, 'Year 5': 6, 'Year 6': 7, 'Year 7': 8, 'Year 8': 9,
                    'Year 9': 10, 'Year 10': 11, 'Year 11': 12, 'Year 12': 13, 'Year 13': 14
                }
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    header_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
                    alt_fill = PatternFill(start_color="EAF1FB", end_color="EAF1FB", fill_type="solid")
                    header_font = Font(bold=True, color="FFFFFF")
                    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

                    def apply_style(ws, col_count, sheet_display_name):
                        # Row 1: Merged Title Header
                        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=col_count)
                        title_cell = ws.cell(row=1, column=1)
                        title_cell.value = f"Tring Fun Run 2026: {sheet_display_name}"
                        title_cell.font = Font(bold=True, size=16)
                        title_cell.alignment = Alignment(horizontal='center', vertical='center')
                        ws.row_dimensions[1].height = 45 # Increased height for padding

                        # Row 2: Headers
                        ws.row_dimensions[2].height = 25
                        for col_idx in range(1, col_count + 1):
                            cell = ws.cell(row=2, column=col_idx)
                            cell.fill, cell.font, cell.border = header_fill, header_font, border
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            
                            # Targeting specific column widths based on header name
                            header_text = str(cell.value)
                            if "School" in header_text:
                                ws.column_dimensions[get_column_letter(col_idx)].width = 45
                            elif "Team" in header_text:
                                ws.column_dimensions[get_column_letter(col_idx)].width = 35
                            elif "Forename" in header_text or "Surname" in header_text:
                                ws.column_dimensions[get_column_letter(col_idx)].width = 25
                            else:
                                ws.column_dimensions[get_column_letter(col_idx)].width = 18
                        
                        # Data Rows
                        for i, row in enumerate(ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=col_count), start=1):
                            for cell in row:
                                cell.border = border
                                if i % 2 == 0: cell.fill = alt_fill
                                cell.alignment = Alignment(vertical='center')

                    # Senior Race Tab
                    res_adult_final = df_export[adult_mask].copy().sort_values('Surname')
                    s_cols = ['Race Number', 'Surname', 'Forename', 'Gender', 'Team name', 'School name', 'School year']
                    res_adult_final[s_cols].to_excel(writer, sheet_name='Senior Adult Race', index=False, startrow=1)
                    apply_style(writer.sheets['Senior Adult Race'], len(s_cols), "Senior Adult Race")

                    # Kids Year Group Tabs
                    kids_mask = (df_export['Ticket'].str.strip() == 'Pre-school to Year 9') & ~df_export.index.isin(res_adult_final.index)
                    kids_df = df_export[kids_mask].copy()
                    # Sorting the list of unique years by our dictionary
                    years = sorted([y for y in kids_df['School year'].unique() if str(y).strip() != ''], key=lambda x: YEAR_ORDER.get(x, 99))
                    
                    for y in years:
                        y_df = kids_df[kids_df['School year'] == y].sort_values('Surname')
                        k_cols = ['Race Number', 'Surname', 'Forename', 'Gender', 'School name']
                        sheet_name = str(y)[:31]
                        y_df[k_cols].to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
                        apply_style(writer.sheets[sheet_name], len(k_cols), y)

                st.success("Tring Race Pack Created!")
                st.download_button("📥 Download Race Pack", output.getvalue(), "Tring_Race_Pack_2026.xlsx")
                st.session_state['processed_reg'] = df_export

# --- TABS 2, 3, & 4 LOGIC (REMAINS CONSISTENT) ---
# ... [Keeping Timer, Results Marriage, and Stats logic exactly as before] ...
