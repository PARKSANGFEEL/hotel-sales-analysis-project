import pandas as pd

file_path = r'c:\Users\HP\Desktop\호텔\부킹매출자료\청구금액.xlsx'

try:
    df = pd.read_excel(file_path)
    print("Columns:", df.columns.tolist())
    print("First 5 rows:")
    print(df.head())
except Exception as e:
    print(f"Error reading excel file: {e}")
