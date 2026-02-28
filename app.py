import streamlit as st
import pandas as pd
import io
import os
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment

st.set_page_config(page_title="Race Entry Processor", layout="wide")
st.title("Fun Run Entry Formatter")

st.sidebar.header("1. Upload Data")
uploaded_file = st.sidebar.file_uploader("Upload Raw Race CSV", type=['csv'])

def load_memory(filename):
    if os.path.exists(filename):
        try:
            mem_df = pd.read_csv(filename)
            mem_df = mem_df.fillna('')
            return dict(zip(mem_df['Raw Name'], mem_df['Cleaned Name']))
        except:
            return {}
    return {}

# The Master Sort Order for School Years
YEAR_ORDER = {
    'Pre-school': 0, 'Reception': 1, 'Year 1': 2, 'Year 2': 3, 'Year 3': 4, 
    'Year 4': 5, 'Year 5': 6, 'Year 6': 7, 'Year 7': 8, 'Year 8': 9, 'Year 9': 10
}

if uploaded_file is not None:
    df = pd.read_csv(uploaded_file).fillna('')
    total_input_rows = len(df)
    
    name_split = df['Full name'].str.split(' ', n=1, expand=True)
    df.insert(2, 'Forename', name_split[0].str.strip().str.title())
    df.insert(3, 'Surname', name_split[1].fillna('').str.strip().str.title())
    df.insert(0, 'Race Number', '')

    if 'School name' in df.columns:
        df['School name'] = df['School name'].astype(str).str.strip().str.title()
    if 'Team name' in df.columns:
        df['Team name'] = df['Team name'].astype(str).str.strip().str.title()

    school_memory = load_memory('school_memory.csv')
    team_memory = load_memory('team_memory.csv')

    st.header("2. Clean Names")
    st.write("Edit the Cleaned Name columns. The app will save these for next time!")
    
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("School Names")
        unique_schools = [s for s in df['School name'].unique() if s.strip() != '' and s.lower() != 'nan']
        school_mapped = [school_memory.get(s, s) for s in unique_schools]
        school_df = pd.DataFrame({"Raw Name": unique_schools, "Cleaned Name": school_mapped}).sort_values(by='Raw Name')
        edited_schools = st.data_editor(school_df, hide_index=True, use_container_width=True, key="school_editor")

    with col2:
        st.subheader("Team Names")
        unique_teams = [t for t in df.get('Team name', pd.Series([])).unique() if str(t).strip() != '' and str(t).lower() != 'nan']
        team_mapped = [team_memory.get(t, t) for t in unique_teams]
        team_df = pd.DataFrame({"Raw Name": unique_teams, "Cleaned Name": team_mapped}).sort_values(by='Raw Name')
        edited_teams = st.data_editor(team_df, hide_index=True, use_container_width=True, key="team_editor")

    st.header("3. Generate Excel File")
    if st.button("Process And Create Excel"):
        
        edited_schools.to_csv('school_memory.csv', index=False)
        edited_teams.to_csv('team_memory.csv', index=False)
        
        school_dict = dict(zip(edited_schools["Raw Name"], edited_schools["Cleaned Name"]))
        df['Cleaned School Name'] = df['School name'].replace(school_dict)
        
        team_dict = dict(zip(edited_teams["Raw Name"], edited_teams["Cleaned Name"]))
        df['Cleaned Team Name'] = df['Team name'].replace(team_dict) if 'Team name' in df.columns else ''
        
        df['Cleaned Team Name'] = df['Cleaned Team Name'].replace(['Nan', 'nan'], '')
        df['Cleaned School Name'] = df['Cleaned School Name'].replace(['Nan', 'nan'], '')

        output = io.BytesIO()
        audit_records = []
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            header_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
            alt_fill = PatternFill(start_color="EAF1FB", end_color="EAF1FB", fill_type="solid")
            title_font = Font(bold=True, size=16)
            header_font = Font(bold=True, color="FFFFFF")

            def apply_style(ws, col_count):
                for cell in ws[2]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.border = thin_border
                for i, row in enumerate(ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=col_count), start=1):
                    for cell in row:
                        cell.border = thin_border
                        if i % 2 == 0: cell.fill = alt_fill
                for col in ws.columns:
                    max_length = max((len(str(cell.value)) for cell in col), default=0)
                    ws.column_dimensions[col[0].column_letter].width = max_length + 4

            # --- Categorization ---
            adult_mask = (df['Ticket'] == 'Senior / Adult Race') | (df['School year'].isin(['Year 10', 'Year 11', 'Year 12', 'Year 13']))
            kids_mask = (df['Ticket'] == 'Pre-school to Year 9')
            donation_mask = ~(adult_mask | kids_mask)

            # 1. Adult Race
            adult_df = df[adult_mask].copy().sort_values(by='Surname')
            adult_cols = ['Race Number', 'Surname', 'Forename', 'Full name', 'Gender', 'Cleaned Team Name', 'School name (and house)', 'School year']
            actual_adult_cols = [c for c in adult_cols if c in adult_df.columns]
            adult_df[actual_adult_cols].to_excel(writer, sheet_name='Senior Adult Race', index=False, startrow=1)
            ws_adult = writer.sheets['Senior Adult Race']
            ws_adult['A1'], ws_adult['A1'].font = "Senior / Adult Race", title_font
            apply_style(ws_adult, len(actual_adult_cols))
            audit_records.append({'Sort': -1, 'Category': 'Senior / Adult Race', 'Count': len(adult_df)})

            # 2. Kids Tabs
            kids_df = df[kids_mask].copy()
            kids_cols = ['Race Number', 'Surname', 'Forename', 'Full name', 'Gender', 'Cleaned School Name']
            actual_kids_cols = [c for c in kids_cols if c in kids_df.columns]
            
            years = sorted([y for y in kids_df['School year'].unique() if y != ''], 
                           key=lambda x: YEAR_ORDER.get(x, 99))
            
            for year in years:
                y_df = kids_df[kids_df['School year'] == year].sort_values(by='Surname')
                y_df[actual_kids_cols].to_excel(writer, sheet_name=str(year)[:31], index=False, startrow=1)
                ws = writer.sheets[str(year)[:31]]
                ws['A1'], ws['A1'].font = f"Race Entries: {year}", title_font
                apply_style(ws, len(actual_kids_cols))
                audit_records.append({'Sort': YEAR_ORDER.get(year, 99), 'Category': f'Kids: {year}', 'Count': len(y_df)})

            # 3. Donation Audit
            donation_df = df[donation_mask]
            audit_records.append({'Sort': 100, 'Category': 'Donations / Other (Excluded)', 'Count': len(donation_df)})

            # --- Audit Summary ---
            audit_df = pd.DataFrame(audit_records).sort_values(by='Sort')
            total_output = audit_df['Count'].sum()
            final_audit_display = audit_df[['Category', 'Count']].copy()
            final_audit_display.loc[len(final_audit_display)] = {'Category': 'GRAND TOTAL', 'Count': total_output}
            
            final_audit_display.to_excel(writer, sheet_name='Audit Summary', index=False, startrow=3)
            ws_audit = writer.sheets['Audit Summary']
            ws_audit['A1'], ws_audit['A1'].font = "Data Reconciliation Report", title_font
            ws_audit['A2'] = f"Original CSV Rows: {total_input_rows} | Total Accounted For: {total_output}"
            ws_audit['A2'].font = Font(color="008000" if total_input_rows == total_output else "FF0000", bold=True)
            apply_style(ws_audit, 2)

        st.divider()
        st.subheader("Audit Results")
        if total_input_rows == total_output:
            st.success(f"Reconciliation Perfect! All {total_input_rows} rows accounted for.")
        else:
            st.error(f"Variance! Input: {total_input_rows} | Accounted: {total_output}")
        st.table(final_audit_display)

        st.download_button(label="Download Processed_Race_Entries.xlsx", data=output.getvalue(), file_name="Processed_Race_Entries.xlsx")