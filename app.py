import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment, Border, Side, Font

# ==========================================
# 1. CONFIGURATION
# ==========================================
st.set_page_config(page_title="Roster Automation", layout="wide")

DRIVE_FOLDER_ID = "1pcZWYGXCC1axVDXWtXp1YyQJ79WVeivr" 
TEMPLATE_FILE = "Template.xlsx"

SEQ = ['C', 'C', 'B', 'B', 'A', 'A', 'W/O']
MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

def get_drive_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build('drive', 'v3', credentials=creds)

def get_latest_roster(service):
    query = f"'{DRIVE_FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
    results = service.files().list(q=query, orderBy="createdTime desc", pageSize=1, includeItemsFromAllDrives=True, supportsAllDrives=True).execute()
    items = results.get('files', [])
    if not items: return None
    request = service.files().get_media(fileId=items[0]['id'])
    return io.BytesIO(request.execute()), items[0]['name']

def get_state(row):
    last_val, prev_val = None, None
    for d in range(31, 1, -1):
        col = str(d)
        if col in row and pd.notna(row[col]):
            if last_val is None: last_val = row[col]
            elif prev_val is None: 
                prev_val = row[col]
                break
    if last_val == 'C': return 2 if prev_val == 'C' else 1
    if last_val == 'B': return 4 if prev_val == 'B' else 3
    if last_val == 'A': return 6 if prev_val == 'A' else 5
    return 0

# ==========================================
# 2. WEB UI & SIDEBAR
# ==========================================
st.title("📅 Monthly Roster Generator")

service = get_drive_service()
latest_file = get_latest_roster(service)

if not latest_file:
    st.error("Upload a baseline file to Drive first.")
    st.stop()

st.sidebar.header("1. Settings")
target_month_name = st.sidebar.selectbox("Month", MONTH_NAMES, index=3)
target_year = st.sidebar.number_input("Year", min_value=2024, max_value=2050, value=2026)
target_month_num = MONTH_NAMES.index(target_month_name) + 1
days_in_month = pd.Period(f'{target_year}-{target_month_num:02d}-01').days_in_month

st.sidebar.markdown("---")
st.sidebar.header("2. Leaves")
if 'leaves' not in st.session_state: st.session_state.leaves = {}

df_prev = pd.read_excel(latest_file[0], skiprows=2)
df_prev.columns = ['S No', 'NAME'] + [str(i) for i in range(1, len(df_prev.columns)-1)]

# FIX: Filter out A, B, C from employee logic
names_list = [n for n in df_prev['NAME'].dropna().unique().tolist() if str(n).strip() not in ["A", "B", "C"]]

sel_name = st.sidebar.selectbox("Select Employee", names_list)
sel_days = st.sidebar.text_input("Enter Days (e.g., 5, 12)")

if st.sidebar.button("Register Leave"):
    st.session_state.leaves[sel_name] = [int(d.strip()) for d in sel_days.split(',') if d.strip().isdigit()]
    st.sidebar.success("Added!")

# ==========================================
# 3. GENERATION
# ==========================================
if st.button(f"Generate Roster ({days_in_month} Days)", type="primary"):
    with st.spinner("Aligning with Expected Snip..."):
        employees = names_list
        emp_state = {name: get_state(row) for name, row in zip(employees, df_prev[df_prev['NAME'].isin(employees)].to_dict('records'))}
        
        roster = {emp: {d: None for d in range(1, days_in_month + 1)} for emp in employees}
        for d in range(1, days_in_month + 1):
            for emp in employees:
                if emp in st.session_state.leaves and d in st.session_state.leaves[emp]:
                    roster[emp][d] = 'L'
                else:
                    shift = SEQ[emp_state[emp]]
                    roster[emp][d] = shift
                    emp_state[emp] = (emp_state[emp] + 1) % 7

        wb = load_workbook(TEMPLATE_FILE)
        ws = wb.active
        
        # 1. CLEAN SLATE
        for m in list(ws.merged_cells.ranges): ws.unmerge_cells(str(m))
        for r in ws.iter_rows(min_row=1, max_row=120, min_col=1, max_col=65):
            for cell in r:
                if cell.column > 2: cell.value = None
                cell.fill = PatternFill(fill_type=None)
                cell.border = Border()

        # 2. STYLES
        thin = Side(border_style="thin", color="000000")
        thick_blue = Side(border_style="thick", color="0000FF")
        all_border = Border(left=thin, right=thin, top=thin, bottom=thin)
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        peach_fill = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid")
        center = Alignment(horizontal='center', vertical='center')
        header_font = Font(bold=True, size=16)
        title_font = Font(bold=True, size=20)

        # 3. LAYOUT
        start_totals = days_in_month + 3
        end_totals = start_totals + 7
        ws.column_dimensions['A'].width = 6.43

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_totals)
        title_cell = ws.cell(row=1, column=1, value=f"DUTY ROSTER FOR THE MONTH OF {target_month_name[:3].upper()} {target_year}")
        title_cell.alignment = center; title_cell.font = title_font

        ws.merge_cells('A2:A3'); ws['A2'] = "S No"; ws['A2'].alignment = center; ws['A2'].font = header_font
        ws.merge_cells('B2:B3'); ws['B2'] = "NAME"; ws['B2'].alignment = center; ws['B2'].font = header_font
        
        ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=days_in_month+2)
        ws.cell(row=2, column=3, value="ATTENDANCE").alignment = center; ws.cell(row=2, column=3).font = header_font
        
        ws.merge_cells(start_row=2, start_column=start_totals, end_row=2, end_column=end_totals)
        ws.cell(row=2, column=start_totals, value="TOTAL SHIFTS").alignment = center; ws.cell(row=2, column=start_totals).font = header_font

        # Day Headers (Width 5)
        for d in range(1, days_in_month + 1):
            col = d + 2
            ws.cell(row=3, column=col, value=d).alignment = center
            ws.column_dimensions[get_column_letter(col)].width = 5
            ws.cell(row=3, column=col).border = Border(left=thin, right=thin, top=thin, bottom=thin)
            
        headers = ['TOTAL', 'A', 'B', 'C', 'W/O', 'X', 'L', 'G']
        for i, h in enumerate(headers):
            col = start_totals + i
            ws.cell(row=3, column=col, value=h).alignment = center
            ws.column_dimensions[get_column_letter(col)].width = 10 if h == 'TOTAL' else 5

        # 4. EMPLOYEE DATA
        num_emp = len(employees)
        for idx, emp in enumerate(employees):
            r = idx + 4
            ws.cell(row=r, column=1, value=idx+1)
            ws.cell(row=r, column=2, value=emp)
            
            for d in range(1, days_in_month + 1):
                ws.cell(row=r, column=d+2, value=roster[emp][d]).alignment = center
            
            # Formulas
            last_d_ltr = get_column_letter(days_in_month + 2)
            tot_ltr = get_column_letter(start_totals)
            ws[f'{tot_ltr}{r}'] = f'=SUM({get_column_letter(start_totals+1)}{r}:{get_column_letter(start_totals+3)}{r})'
            for i_h, h_code in enumerate(['A*', 'B*', 'C*', 'W/O*', 'X*', 'L*', 'G*']):
                ws.cell(row=r, column=start_totals+1+i_h, value=f'=COUNTIF(C{r}:{last_d_ltr}{r},"{h_code}")')

            # Paint
            ws[f'{tot_ltr}{r}'].fill = yellow_fill
            ws[f'{tot_ltr}{r}'].font = Font(bold=True)
            for c_paint in range(start_totals + 1, end_totals + 1):
                ws.cell(row=r, column=c_paint).fill = peach_fill

            # Apply Main Table Borders (Blue Outlines)
            for c in range(1, end_totals + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = Border(left=thick_blue if c==1 else thin, right=thick_blue if c==end_totals else thin, top=thin, bottom=thin)

        # Apply Header Borders (Rows 1-3)
        for r_h in range(1, 4):
            for c_h in range(1, end_totals + 1):
                ws.cell(row=r_h, column=c_h).border = Border(left=thick_blue if c_h==1 else thin, right=thick_blue if c_h==end_totals else thin, top=thick_blue if r_h==1 else thin, bottom=thin)

        # Apply Bottom Thick Blue Border to last employee row
        for c in range(1, end_totals + 1):
            ws.cell(row=num_emp+3, column=c).border = Border(left=thick_blue if c==1 else thin, right=thick_blue if c==end_totals else thin, top=thin, bottom=thick_blue)

        # 5. EXPECTED YELLOW SUMMARY BOX (A, B, C counts)
        summary_row = num_emp + 6
        for i, s_type in enumerate(["A", "B", "C"]):
            r_sum = summary_row + i
            # Label Cell (Yellow, Bold, Center)
            l_cell = ws.cell(row=r_sum, column=2, value=s_type)
            l_cell.fill = yellow_fill; l_cell.font = Font(bold=True); l_cell.alignment = center; l_cell.border = all_border
            
            # Daily Count Cells (Yellow, Center)
            for d in range(1, days_in_month + 1):
                col = d + 2
                col_ltr = get_column_letter(col)
                v_cell = ws.cell(row=r_sum, column=col, value=f'=COUNTIF({col_ltr}4:{col_ltr}{num_emp+3},"{s_type}*")')
                v_cell.fill = yellow_fill; v_cell.alignment = center; v_cell.border = all_border

        out = io.BytesIO()
        wb.save(out)
        st.balloons()
        st.download_button("Download Expected Roster", out.getvalue(), f"ROSTER_{target_month_name}.xlsx")
