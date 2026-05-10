import streamlit as st
import pandas as pd
import io
import os
import datetime
from streamlit_gsheets import GSheetsConnection
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# 1. Page Configuration
st.set_page_config(page_title="Tring Fun Run Tools", layout="wide")
st.title("🏃‍♂️ Tring Fun Run: 2026 Registration & Results")

# --- 2. MODES & CONNECTION ---
conn = st.connection("gsheets", type=GSheetsConnection)
LOCAL_FILENAME = "tring_offline_backup.xlsx"

st.sidebar.header("⚙️ System Settings")
app_mode = st.sidebar.radio("Data Mode", ["Cloud (Google Sheets)", "Local (Offline Excel)"])

if app_mode == "Local (Offline Excel)":
    st.sidebar.warning("⚠️ Running in OFFLINE mode. Data is saving to your laptop, NOT the cloud.")

# --- 3. UNIVERSAL DATA HANDLERS ---

def get_default_df(worksheet_name):
    """Returns a blank DataFrame with correct headers for the specified tab."""
    if worksheet_name in ["Schools", "Teams"]:
        return pd.DataFrame(columns=["Raw Name", "Cleaned Name"])
    elif worksheet_name == "SchoolRolls":
        return pd.DataFrame(columns=["School Name", "Infants Roll", "Juniors Roll"])
    return pd.DataFrame(columns=["Forename", "Surname", "Gender", "School name", "Team name", "School year", "Race Number", "Ticket"])

@st.cache_data(ttl=600)
def load_data(worksheet_name):
    if app_mode == "Cloud (Google Sheets)":
        try:
            df = conn.read(worksheet=worksheet_name, ttl=0).fillna('')
            return df if not df.empty else get_default_df(worksheet_name)
        except:
            return get_default_df(worksheet_name)
    else:
        # LOCAL MODE
        if os.path.exists(LOCAL_FILENAME):
            try:
                with pd.ExcelFile(LOCAL_FILENAME) as xls:
                    if worksheet_name in xls.sheet_names:
                        return pd.read_excel(xls, sheet_name=worksheet_name).fillna('')
            except:
                pass
        return get_default_df(worksheet_name)

def save_data(df, worksheet_name):
    """Saves data based on the current mode."""
    if app_mode == "Cloud (Google Sheets)":
        conn.update(worksheet=worksheet_name, data=df)
        st.cache_data.clear()
    else:
        # LOCAL MODE: Save to the local Excel workbook
        mode = 'a' if os.path.exists(LOCAL_FILENAME) else 'w'
        if_sheet_exists = 'replace' if mode == 'a' else None
        
        with pd.ExcelWriter(LOCAL_FILENAME, engine='openpyxl', mode=mode, if_sheet_exists=if_sheet_exists) as writer:
            df.to_excel(writer, sheet_name=worksheet_name, index=False)
        st.cache_data.clear()
        st.toast(f"Saved to {LOCAL_FILENAME}")

# Sidebar Utilities
if st.sidebar.button("🔄 Clear Cache / Refresh"):
    st.cache_data.clear()
    st.rerun()

if st.sidebar.button("💾 Create Local Backup Snapshot"):
    # This pulls everything from cloud and saves it into the local Excel file
    st.toast("Creating full local backup...")
    tabs = ["LateEntries", "BibAllocations", "Schools", "Teams", "SchoolRolls"]
    with pd.ExcelWriter(LOCAL_FILENAME, engine='openpyxl') as writer:
        for tab in tabs:
            try:
                temp_df = conn.read(worksheet=tab, ttl=0).fillna('')
                temp_df.to_excel(writer, sheet_name=tab, index=False)
            except:
                get_default_df(tab).to_excel(writer, sheet_name=tab, index=False)
    st.sidebar.success("Local backup file updated!")

# --- 4. PRE-LOAD MASTER DATA ---
late_entries_master = load_data("LateEntries")
bib_allocs_master = load_data("BibAllocations")
school_mem_master = load_data("Schools")
team_mem_master = load_data("Teams")
rolls_master = load_data("SchoolRolls")

# --- 5. INTERFACE TABS ---
tab_entry, tab_timer, tab_results, tab_stats = st.tabs([
    "📋 Registration & Bibs", "⏱️ Timer Reconciliation", "🏁 Final Results Marriage", "📊 Participation Stats"
])

# --- TAB 1: REGISTRATION ---
with tab_entry:
    col_late, col_pre = st.columns([1, 2])
    
    with col_late:
        st.header("1. On-The-Day Entry")
        with st.form("otd_form", clear_on_submit=True):
            f_name = st.text_input("Forename")
            s_name = st.text_input("Surname")
            gender = st.selectbox("Gender", ["Male", "Female", "Other"])
            yr = st.selectbox("Year Group", ["Adult", "Pre-school", "Reception"] + [f"Year {i}" for i in range(1, 14)])
            team, school, bib = st.text_input("Team"), st.text_input("School"), st.text_input("Assigned Bib")
            
            if st.form_submit_button("Submit"):
                if f_name and s_name and bib:
                    tkt = "Senior / Adult Race" if yr == "Adult" or yr in ["Year 10", "Year 11", "Year 12", "Year 13"] else "Pre-school to Year 9"
                    new_entry = pd.DataFrame([{
                        "Forename": f_name.title(), "Surname": s_name.title(), "Gender": gender, 
                        "School year": yr, "Team name": team.title(), "School name": school.title(), 
                        "Race Number": str(bib).strip(), "Ticket": tkt
                    }])
                    updated = pd.concat([late_entries_master, new_entry], ignore_index=True)
                    save_data(updated, "LateEntries")
                    st.success(f"Bib {bib} saved.")

    with col_pre:
        st.header("2. Pre-Reg & Packs")
        csv_file = st.sidebar.file_uploader("Upload CSV", type=['csv'], key="pre_reg_upload")
        if csv_file:
            df_raw = pd.read_csv(csv_file).fillna('')
            if 'Full name' in df_raw.columns:
                name_split = df_raw['Full name'].str.split(' ', n=1, expand=True)
                df_raw['Forename'] = name_split[0].str.strip().str.title()
                df_raw['Surname'] = name_split[1].fillna('').str.strip().str.title()

            # Mapping Logic
            new_sch = [s for s in df_raw['School name'].unique() if str(s).strip() != '']
            sch_to_add = [s for s in new_sch if s not in school_mem_master['Raw Name'].values]
            if sch_to_add:
                school_mem_master = pd.concat([school_mem_master, pd.DataFrame({"Raw Name": sch_to_add, "Cleaned Name": sch_to_add})], ignore_index=True)
            
            ed_sch = st.data_editor(school_mem_master.sort_values("Raw Name"), key="sch_ed", hide_index=True)
            
            # Bib Assignment
            df_raw['Ticket'] = df_raw['Ticket'].str.strip()
            adult_mask = (df_raw['Ticket'].str.contains('Senior', case=False, na=False)) | (df_raw['School year'].isin(['Year 10', 'Year 11', 'Year 12', 'Year 13']))
            pre_reg_adults = df_raw[adult_mask][['Forename', 'Surname', 'Gender', 'Team name', 'School year', 'Ticket']].copy()
            
            if not bib_allocs_master.empty:
                pre_reg_adults = pre_reg_adults.merge(bib_allocs_master[['Forename', 'Surname', 'Race Number']], on=['Forename', 'Surname'], how='left')
            else:
                pre_reg_adults['Race Number'] = ""
            
            ed_bibs = st.data_editor(pre_reg_adults.sort_values('Surname'), key="bib_ed", hide_index=True)

            if st.button("Process & Generate Pack"):
                save_data(ed_sch, "Schools")
                save_data(ed_bibs, "BibAllocations")
                
                # Excel Generation [Standard formatting logic applies]
                st.success("✅ Processed. Download link generated below.")
                # (Excel Export Logic omitted for brevity, but same as previous working version)

# --- TAB 2, 3, & 4 ---
# These tabs use the 'late_entries_master' and 'bib_allocs_master' defined above.
# Because those masters use the 'load_data' function, they automatically switch 
# between Cloud and Local depending on your sidebar choice.

with tab_results:
    st.header("🏁 Results Marriage")
    if 'master_timer' in st.session_state:
        runners = pd.concat([late_entries_master, bib_allocs_master], ignore_index=True)
        # Results logic continues...
