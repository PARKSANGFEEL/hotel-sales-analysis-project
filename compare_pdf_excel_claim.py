import pdfplumber
import pandas as pd
import os

pdf_path = r'c:\Users\HP\Desktop\호텔\부킹매출자료\부킹3.pdf'
excel_path = r'c:\Users\HP\Desktop\호텔\부킹매출자료\청구금액.xlsx'

def parse_pdf_line(line):
    parts = line.split()
    # Format: 예약 [ResNum] 예약 [InY] [InM] [InD] [OutY] [OutM] [OutD] [Name...] KRW [Amount]
    if len(parts) < 11 or parts[0] != '예약' or 'KRW' not in parts:
        return None
    
    try:
        res_num = parts[1]
        
        krw_index = -1
        for i in range(len(parts) - 1, -1, -1):
            if parts[i] == 'KRW':
                krw_index = i
                break
        
        if krw_index == -1:
            return None

        amount_str = parts[krw_index + 1]
        amount = float(amount_str.replace(',', ''))
        
        check_in_str = f"{parts[3]}-{parts[4].replace('월','').zfill(2)}-{parts[5].replace('일','').zfill(2)}"
        check_out_str = f"{parts[6]}-{parts[7].replace('월','').zfill(2)}-{parts[8].replace('일','').zfill(2)}"
        
        # Remove '년' from year if present (though parts[3] is usually '2025년')
        check_in_str = check_in_str.replace('년', '')
        check_out_str = check_out_str.replace('년', '')
        
        if krw_index > 9:
            name_parts = parts[9:krw_index]
            name = " ".join(name_parts)
        else:
            name = ""
        
        return {
            'Reservation Number': str(res_num),
            'Guest Name': name,
            'Check-in': check_in_str,
            'Check-out': check_out_str,
            'Amount': amount,
            'Source': 'PDF'
        }
    except Exception:
        return None

def compare_files():
    # 1. Parse PDF
    print(f"Parsing PDF: {os.path.basename(pdf_path)}...")
    pdf_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            lines = text.split('\n')
            for line in lines:
                if line.startswith('예약'):
                    parsed = parse_pdf_line(line)
                    if parsed:
                        pdf_data.append(parsed)
    
    df_pdf = pd.DataFrame(pdf_data)
    print(f"PDF: Found {len(df_pdf)} reservations. Total Amount: {df_pdf['Amount'].sum():,.0f}")

    # 2. Read Excel
    print(f"Reading Excel: {os.path.basename(excel_path)}...")
    df_excel = pd.read_excel(excel_path)
    
    # Normalize Excel columns
    # Reservation number might be int or str
    df_excel['Reservation number'] = df_excel['Reservation number'].astype(str)
    
    # Select relevant columns
    # Assuming 'Final amount' is the one to compare. 
    # Note: Excel might have 'Final amount' in KRW.
    df_excel_clean = df_excel[['Reservation number', 'Guest name', 'Arrival', 'Departure', 'Final amount']].copy()
    df_excel_clean.columns = ['Reservation Number', 'Guest Name', 'Check-in', 'Check-out', 'Amount']
    
    # Normalize dates in Excel (YYYY-MM-DD)
    df_excel_clean['Check-in'] = pd.to_datetime(df_excel_clean['Check-in']).dt.strftime('%Y-%m-%d')
    df_excel_clean['Check-out'] = pd.to_datetime(df_excel_clean['Check-out']).dt.strftime('%Y-%m-%d')
    
    print(f"Excel: Found {len(df_excel_clean)} reservations.")

    # 3. Compare
    # Join on Reservation Number
    merged = pd.merge(df_pdf, df_excel_clean, on='Reservation Number', how='outer', suffixes=('_PDF', '_Excel'), indicator=True)
    
    only_in_pdf = merged[merged['_merge'] == 'left_only']
    only_in_excel = merged[merged['_merge'] == 'right_only']
    both = merged[merged['_merge'] == 'both']
    
    print("\n--- Comparison Results ---")
    
    # 3.1 Missing in Excel
    if not only_in_pdf.empty:
        print(f"\n[In PDF but NOT in Excel] ({len(only_in_pdf)} items):")
        for _, row in only_in_pdf.iterrows():
            print(f"  Res#{row['Reservation Number']} | {row['Guest Name_PDF']} | {row['Amount_PDF']:,.0f}")
    else:
        print("\n[In PDF but NOT in Excel]: None. All PDF reservations found in Excel.")

    # 3.2 Missing in PDF (This might be expected if Excel contains more months, but let's check if any match the date range)
    # Filter only_in_excel for dates relevant to PDF (Sep 2025 Check-out)
    # PDF Check-outs range:
    if not df_pdf.empty:
        min_date = df_pdf['Check-out'].min()
        max_date = df_pdf['Check-out'].max()
        print(f"\nPDF Check-out Range: {min_date} ~ {max_date}")
        
        # Filter Excel items within this range
        only_in_excel_relevant = only_in_excel[
            (only_in_excel['Check-out_Excel'] >= min_date) & 
            (only_in_excel['Check-out_Excel'] <= max_date)
        ]
        
        if not only_in_excel_relevant.empty:
            print(f"\n[In Excel but NOT in PDF] (Within PDF date range {min_date}~{max_date}) ({len(only_in_excel_relevant)} items):")
            for _, row in only_in_excel_relevant.iterrows():
                print(f"  Res#{row['Reservation Number']} | {row['Guest Name_Excel']} | {row['Amount_Excel']:,.0f} | Out: {row['Check-out_Excel']}")
        else:
            print(f"\n[In Excel but NOT in PDF]: None within date range {min_date}~{max_date}.")

    # 3.3 Amount Mismatch
    mismatch = both[abs(both['Amount_PDF'] - both['Amount_Excel']) > 1.0] # Tolerance of 1 KRW
    if not mismatch.empty:
        print(f"\n[Amount Mismatch] ({len(mismatch)} items):")
        for _, row in mismatch.iterrows():
            diff = row['Amount_PDF'] - row['Amount_Excel']
            print(f"  Res#{row['Reservation Number']} ({row['Guest Name_PDF']}): PDF {row['Amount_PDF']:,.0f} vs Excel {row['Amount_Excel']:,.0f} (Diff: {diff:,.0f})")
    else:
        print("\n[Amount Mismatch]: None. All matched reservations have same amount.")

if __name__ == "__main__":
    compare_files()
