
import pandas as pd

files = [
    r"D:\GITHUB\CBLand\data\행정구역_시군구_별_주민등록세대수_20251207114706.csv",
    r"D:\GITHUB\CBLand\original_data\충청북도_지적통계_20250630.csv"
]

with open("d:/GITHUB/CBLand/column_names.txt", "w", encoding="utf-8") as out:
    for f in files:
        out.write(f"--- Analyzing {f} ---\n")
        try:
            df = pd.read_csv(f, encoding='cp949')
            out.write("Columns:\n")
            for i, c in enumerate(df.columns):
                out.write(f"{i}: {c}\n")
            out.write("\nFirst Row:\n")
            # Convert values to string to avoid encoding issues if they are mixed
            vals = [str(x) for x in df.iloc[0].tolist()] 
            out.write(", ".join(vals) + "\n\n")
        except Exception as e:
            out.write(f"Error: {e}\n")
