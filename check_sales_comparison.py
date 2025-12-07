import pandas as pd

file_path = '매출상세내역 목록_20251207_final_v4.xlsx'

def check_pms_sales():
    try:
        # Check '월별매출' (PMS Data)
        df_pms = pd.read_excel(file_path, sheet_name='월별매출')
        print("--- PMS Monthly Sales (월별매출) ---")
        print(df_pms)
        
        # Check '부킹_월별매출' (Booking Data)
        df_booking = pd.read_excel(file_path, sheet_name='부킹_월별매출')
        print("\n--- Booking Monthly Sales (부킹_월별매출) ---")
        print(df_booking)

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_pms_sales()
