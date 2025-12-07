
import pandas as pd
import sys

csv_path = r"D:\GITHUB\CBLand\result\final_comprehensive_analysis.csv"

try:
    df = pd.read_csv(csv_path, encoding='utf-8-sig')
    print("Successfully read CSV.")
    print("Regions found:")
    print(df['Region'].unique())
    
    expected_count = 14
    current_count = len(df)
    print(f"Row count: {current_count}")
    
    if current_count < expected_count:
        print("WARNING: Missing regions.")
    
except Exception as e:
    print(f"Error reading CSV: {e}")
