import pandas as pd

folder = r'c:\Users\HP\Desktop\호텔\익스피디아매출자료'
files = ['statementsList.csv', 'autopayList.csv']

for f in files:
    path = f"{folder}\\{f}"
    print(f"--- Inspecting {f} ---")
    try:
        # Try reading with default encoding first, then utf-8, then cp949 (common for Korean CSVs)
        try:
            df = pd.read_csv(path)
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(path, encoding='utf-8')
            except UnicodeDecodeError:
                df = pd.read_csv(path, encoding='cp949')
                
        print("Columns:", df.columns.tolist())
        print(df.head())
    except Exception as e:
        print(f"Error reading {f}: {e}")
