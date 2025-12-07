
import pandas as pd
import sys

pop_file = r"D:\GITHUB\CBLand\data\행정구역_시군구_별_주민등록세대수_20251207114706.csv"
land_file = r"D:\GITHUB\CBLand\original_data\충청북도_지적통계_20250630.csv"

# Function to read and print unique names
def check_names(f, encoding):
    try:
        df = pd.read_csv(f, encoding=encoding)
        if '행정구역(시군구)별' in df.columns:
            print(f"\n--- Population File ({encoding}) ---")
            print(df['행정구역(시군구)별'].unique())
        elif '토지소재명' in df.columns:
            print(f"\n--- Land File ({encoding}) ---")
            print(df['토지소재명'].unique())
        else:
            print(f"\nCould not find name column in {f}")
    except Exception as e:
        print(f"\nError reading {f} with {encoding}: {e}")

# Try standard encodings
encodings = ['cp949', 'euc-kr', 'utf-8']

with open("d:/GITHUB/CBLand/raw_names_debug.txt", "w", encoding="utf-8") as out:
    sys.stdout = out
    for enc in encodings:
        check_names(pop_file, enc)
        check_names(land_file, enc)
