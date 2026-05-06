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

# --- NEW CACHED LOADING LOGIC ---
@st.cache_data(ttl=600) # Cache data for 10 mins to save API quota
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

# Manual Refresh Button in Sidebar
if st.sidebar.button("🔄 Sync with Google Sheets"):
    st.cache_data.clear()
    st.toast("Cloud cache cleared. Fetching fresh data...")

# Pre-load the main data ONCE per rerun
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
            bib = st.text_input("Assigned Bib")
            
            if st.form_submit_button("Submit & Save"):
                if f_name and s_name and bib:
                    tkt = "Senior / Adult Race" if yr == "Adult" or yr in ["Year 10", "Year 11", "Year 12", "Year 13"] else "Pre-school to Year 9"
                    new_entry = pd.DataFrame([{
                        "Forename": f_name.title(), "Surname": s_name.title(),
                        "Gender": gender, "School year": yr, "Team name": team.title(),
                        "School name": school.title(), "Race Number": str(bib).strip(), "Ticket": tkt
                    }])
                    updated = pd.concat([late_entries_master, new_entry], ignore_index=True)
                    conn.update(worksheet="LateEntries", data=updated)
                    st.cache_data.clear() # Clear cache so the new entry shows up
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

            # Merge Memory with local edits
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
            st.subheader("3. Bib Assignment")
            adult_mask = (df['Ticket'].str.strip() == 'Senior / Adult Race') | (df['School year'].isin(['Year 10', 'Year 11', 'Year 12', 'Year 13']))
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
                
                # Apply Cleanup
                sch_dict = dict(zip(ed_sch["Raw Name"], ed_sch["Cleaned Name"]))
                tm_dict = dict(zip(ed_tm["Raw Name"], ed_tm["Cleaned Name"]))
                df_export = df.copy()
                df_export['School name'] = df_export['School name'].replace(sch_dict)
                df_export['Team name'] = df_export['Team name'].replace(tm_dict)

                output = io.BytesIO()
                YEAR_ORDER = {'Pre-school': 0, 'Reception': 1, 'Year 1': 2, 'Year 2': 3, 'Year 3': 4, 'Year 4': 5, 'Year 5': 6, 'Year 6': 7, 'Year 7': 8, 'Year 8': 9, 'Year 9': 10, 'Year 10': 11}
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # [Formatting Logic remains same as previous version]
                    # ... (skipping formatting code for brevity, but it should be included in your file)
                    res_adult_final = df_export[adult_mask].copy().sort_values('Surname')
                    res_adult_final[['Race Number', 'Surname', 'Forename', 'Gender', 'Team name', 'School name', 'School year']].to_excel(writer, sheet_name='Senior Adult Race', index=False, startrow=1)
                
                st.success("Race Pack Ready!")
                st.download_button("📥 Download", output.getvalue(), "Tring_Race_Pack.xlsx")
                st.session_state['processed_reg'] = df_export
                st.cache_data.clear() # Refresh cloud data after update

# --- TAB 2 & 3: RECONCILIATION & MARRIAGE ---
with tab_timer:
    st.header("Timer Reconciliation")
    # [Timer Logic remains same - no cloud reads here]

with tab_results:
    st.header("Results Marriage")
    if 'master_timer' in st.session_state:
        # We use our pre-loaded cloud data here!
        runners = pd.concat([late_entries_master, bib_allocs_master], ignore_index=True)
        runners['Race Number'] = runners['Race Number'].astype(str).str.strip()
        
        # [Marriage merge and preview logic remains same]

# --- TAB 4: STATS ---
with tab_stats:
    st.header("📊 Participation Stats")
    # We use our pre-loaded cloud data here too!
    pre_reg = st.session_state.get('processed_reg', pd.DataFrame())
    df_all = pd.concat([late_entries_master, pre_reg], ignore_index=True)
    # [Stats calculation remains same]
