import pandas as pd

excel_path = r'c:\Users\HP\Desktop\호텔\부킹매출자료\청구금액.xlsx'

def check_status():
    df = pd.read_excel(excel_path)
    df['Reservation number'] = df['Reservation number'].astype(str)
    
    # Check the mismatched item
    mismatch_res = '4378603680'
    row = df[df['Reservation number'] == mismatch_res]
    if not row.empty:
        print(f"--- Mismatched Item ({mismatch_res}) ---")
        print(row[['Reservation number', 'Guest name', 'Final amount', 'Original amount', 'Commission amount', 'Status', 'Guest request']].to_string())
    
    # Check 0 amount items (sample)
    print("\n--- Sample 0 Amount Items ---")
    zero_rows = df[df['Final amount'] == 0].head(5)
    print(zero_rows[['Reservation number', 'Guest name', 'Final amount', 'Status']].to_string())
    
    # Count statuses for 0 amount items
    zero_df = df[df['Final amount'] == 0]
    print("\n--- Status Counts for 0 Amount Items ---")
    print(zero_df['Status'].value_counts())

if __name__ == "__main__":
    check_status()
