
import pandas as pd
import os

files = [
    r"D:\GITHUB\CBLand\data\행정구역_시군구_별_주민등록세대수_20251207114706.csv",
    r"D:\GITHUB\CBLand\original_data\충청북도_지적통계_20250630.csv"
]

for f in files:
    print(f"--- Analyzing {os.path.basename(f)} ---")
    try:
        df = pd.read_csv(f, encoding='cp949')
        print("Columns:", df.columns.tolist())
        print("First row:", df.iloc[0].tolist())
    except Exception as e:
        print(f"Error reading with cp949: {e}")
        try:
            df = pd.read_csv(f, encoding='euc-kr')
            print("Columns (euc-kr):", df.columns.tolist())
        except Exception as e2:
            print(f"Error reading with euc-kr: {e2}")
