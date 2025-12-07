import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Border, Side

file_path = '매출상세내역 목록_20251207_styled.xlsx'
source_sheet = '매출정리'

try:
    # Load the source data
    df = pd.read_excel(file_path, sheet_name=source_sheet)
    
    # Remove the '합계' row if it exists
    # The previous script added a row where '객실' == '합계'
    df = df[df['객실'] != '합계']
    
    # Ensure numeric columns are numbers
    numeric_cols = ['객실료', '부가세', '총금액']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Convert '일자' to datetime and extract Month
    df['일자'] = pd.to_datetime(df['일자'])
    df['월'] = df['일자'].dt.strftime('%Y-%m')

    # 1. Monthly Sales
    monthly_sales = df.groupby('월')[numeric_cols].sum().reset_index()
    
    # 2. Room Sales
    room_sales = df.groupby('객실')[numeric_cols].sum().reset_index()
    
    # 3. Monthly & Room Sales (Pivot)
    # Using pivot_table to show Rooms as rows and Months as columns for Total Amount
    monthly_room_sales = df.pivot_table(index='객실', columns='월', values='총금액', aggfunc='sum', fill_value=0).reset_index()

    # Load existing workbook
    wb = openpyxl.load_workbook(file_path)
    
    # Helper function to add sheet if not exists and write data
    def write_dataframe_to_sheet(workbook, sheet_name, dataframe):
        if sheet_name in workbook.sheetnames:
            print(f"Sheet '{sheet_name}' already exists. Skipping.")
            return workbook[sheet_name]
        else:
            ws = workbook.create_sheet(sheet_name)
            for r in dataframe_to_rows(dataframe, index=False, header=True):
                ws.append(r)
            print(f"Created sheet '{sheet_name}'.")
            return ws

    # Add sheets
    ws_monthly = write_dataframe_to_sheet(wb, '월별매출', monthly_sales)
    ws_room = write_dataframe_to_sheet(wb, '룸별매출', room_sales)
    ws_monthly_room = write_dataframe_to_sheet(wb, '월별_룸별매출', monthly_room_sales)

    # Formatting function
    def apply_styles(ws):
        thin_border = Border(left=Side(style='thin'), 
                             right=Side(style='thin'), 
                             top=Side(style='thin'), 
                             bottom=Side(style='thin'))
        comma_format = '#,##0'
        
        for row in ws.iter_rows():
            for cell in row:
                cell.border = thin_border
                # Apply comma format to numeric values (excluding header)
                if cell.row > 1 and isinstance(cell.value, (int, float)):
                    cell.number_format = comma_format

    # Apply styles to new sheets
    if '월별매출' in wb.sheetnames: apply_styles(wb['월별매출'])
    if '룸별매출' in wb.sheetnames: apply_styles(wb['룸별매출'])
    if '월별_룸별매출' in wb.sheetnames: apply_styles(wb['월별_룸별매출'])

    wb.save(file_path)
    print(f"Analysis completed and saved to {file_path}")

except Exception as e:
    print(f"An error occurred: {e}")
