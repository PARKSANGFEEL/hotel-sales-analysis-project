import pandas as pd

file_path = '매출상세내역 목록_20251207_final_v10.xlsx'
sheet_name = '점유율_및_평균단가'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
    print("Columns:", df.columns.tolist())
    print("\nFirst 5 rows:")
    print(df.head())
    
    print("\nNon-zero check for Single Room:")
    # Check unique room types and room numbers
    df_source = pd.read_excel('매출상세내역 목록_20251207_final_v5.xlsx', sheet_name='매출정리')
    print("\nUnique Room Numbers in Source:")
    print(df_source['객실'].unique())
    
    # Check how they are mapped
    room_mapping = {
        '302': '이코노미더블룸', '305': '이코노미더블룸', '306': '이코노미더블룸',
        '205': '싱글룸', '206': '싱글룸', '405': '싱글룸', '406': '싱글룸',
        '505': '싱글룸', '506': '싱글룸', '605': '싱글룸', '606': '싱글룸',
        '705': '싱글룸', '706': '싱글룸',
        '203': '트윈룸', '204': '트윈룸', '303': '트윈룸', '304': '트윈룸',
        '403': '트윈룸', '404': '트윈룸', '503': '트윈룸', '504': '트윈룸',
        '603': '트윈룸', '604': '트윈룸', '703': '트윈룸', '704': '트윈룸',
        '202': '트리플룸', '402': '트리플룸', '502': '트리플룸', 
        '602': '트리플룸', '702': '트리플룸', 
        '801': '패밀리4',
        '803': '패밀리5'
    }
    def get_room_type(room):
        return room_mapping.get(str(room).strip(), '기타')
        
    df_source['객실타입'] = df_source['객실'].apply(get_room_type)
    print("\nUnique Room Types Mapped:")
    print(df_source['객실타입'].value_counts())

except Exception as e:
    print(f"Error reading file: {e}")

