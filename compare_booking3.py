import pdfplumber
import pandas as pd
import os

pdf_path = r'c:\Users\HP\Desktop\호텔\부킹매출자료\부킹3.pdf'
excel_path = '매출상세내역 목록_20251207_final_v4.xlsx'

def parse_pdf_line(line):
    parts = line.split()
    if len(parts) < 11 or parts[0] != '예약' or 'KRW' not in parts:
        return None
    
    try:
        krw_index = -1
        for i in range(len(parts) - 1, -1, -1):
            if parts[i] == 'KRW':
                krw_index = i
                break
        
        if krw_index == -1:
            return None

        amount_str = parts[krw_index + 1]
        amount = float(amount_str.replace(',', ''))
        
        check_in_str = f"{parts[3]} {parts[4]} {parts[5]}"
        check_out_str = f"{parts[6]} {parts[7]} {parts[8]}"
        
        if krw_index > 9:
            name_parts = parts[9:krw_index]
            name = " ".join(name_parts)
        else:
            name = ""
        
        return {
            '체크인': check_in_str,
            '체크아웃': check_out_str,
            '고객명': name,
            '금액': amount,
            'Source': 'PDF'
        }
    except Exception:
        return None

def compare_data():
    # 1. Parse PDF
    pdf_data = []
    print(f"Parsing {pdf_path}...")
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
    
    # Clean dates in PDF data for comparison
    def clean_date(d):
        return d.replace('년', '-').replace('월', '-').replace('일', '').replace(' ', '')
    df_pdf['체크아웃_clean'] = df_pdf['체크아웃'].apply(clean_date)
    df_pdf['체크아웃_dt'] = pd.to_datetime(df_pdf['체크아웃_clean'])
    df_pdf['월'] = df_pdf['체크아웃_dt'].dt.strftime('%Y-%m')
    
    print(f"PDF Total Amount: {df_pdf['금액'].sum():,.0f}")
    print(f"PDF Row Count: {len(df_pdf)}")
    print("PDF Month Distribution:")
    print(df_pdf['월'].value_counts())

    # 2. Read Excel
    print(f"\nReading Excel {excel_path}...")
    df_excel = pd.read_excel(excel_path, sheet_name='부킹매출')
    df_excel = df_excel[df_excel['체크인'] != '합계']
    
    # Filter Excel for September 2025 (Target Month for Booking 3)
    # Booking 3 is paid in Oct, so it covers Sep.
    df_excel['체크아웃_dt'] = pd.to_datetime(df_excel['체크아웃'])
    df_excel['월'] = df_excel['체크아웃_dt'].dt.strftime('%Y-%m')
    
    df_excel_sep = df_excel[df_excel['월'] == '2025-09']
    
    print(f"Excel (Sep 2025 Check-out) Total Amount: {df_excel_sep['금액'].sum():,.0f}")
    print(f"Excel (Sep 2025 Check-out) Row Count: {len(df_excel_sep)}")

    # 3. Compare
    # We want to see if there are items in PDF that are NOT in Excel Sep, or items in Excel Sep NOT in PDF.
    
    # Create keys for comparison (Name + Amount + Checkout)
    df_pdf['key'] = df_pdf['고객명'] + '_' + df_pdf['금액'].astype(str) + '_' + df_pdf['체크아웃_clean']
    df_excel_sep['key'] = df_excel_sep['고객명'] + '_' + df_excel_sep['금액'].astype(str) + '_' + df_excel_sep['체크아웃'].astype(str)
    
    pdf_keys = set(df_pdf['key'])
    excel_keys = set(df_excel_sep['key'])
    
    only_in_pdf = pdf_keys - excel_keys
    only_in_excel = excel_keys - pdf_keys
    
    print(f"\n--- Comparison Results ---")
    if only_in_pdf:
        print(f"Items in PDF but NOT in Excel Sep 2025 ({len(only_in_pdf)}):")
        for k in only_in_pdf:
            row = df_pdf[df_pdf['key'] == k].iloc[0]
            print(f"  {row['고객명']} / {row['체크아웃_clean']} / {row['금액']:,.0f} (Month: {row['월']})")
    else:
        print("All items in PDF are present in Excel Sep 2025 (or accounted for).")

    if only_in_excel:
        print(f"\nItems in Excel Sep 2025 but NOT in PDF ({len(only_in_excel)}):")
        for k in only_in_excel:
            row = df_excel_sep[df_excel_sep['key'] == k].iloc[0]
            print(f"  {row['고객명']} / {row['체크아웃']} / {row['금액']:,.0f}")
    else:
        print("All items in Excel Sep 2025 are present in PDF.")

if __name__ == "__main__":
    compare_data()
