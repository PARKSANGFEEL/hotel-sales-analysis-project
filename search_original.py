import pandas as pd

file_path = '매출상세내역 목록_20251207.xls'

def search_original_file():
    try:
        df = pd.read_excel(file_path)
        print("Searching for 'Booking' in original file...")
        
        # Search in all string columns
        mask = df.apply(lambda x: x.astype(str).str.contains('Booking', case=False, na=False)).any(axis=1)
        booking_rows = df[mask]
        
        print(f"Found {len(booking_rows)} rows containing 'Booking'.")
        if not booking_rows.empty:
            print(booking_rows.head())
            
        # Also check '참조' column unique values
        if '참조' in df.columns:
            print("\nUnique values in '참조':")
            print(df['참조'].unique())

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    search_original_file()
