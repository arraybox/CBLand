
import pandas as pd
import os
import re

# Config
DATA_DIR = r"D:\GITHUB\CBLand\original_data"
RESULT_DIR = r"D:\GITHUB\CBLand\result_v2"
POP_FILE = r"D:\GITHUB\CBLand\data\행정구역_시군구_별_주민등록세대수_20251207114706.csv"

# File Mapping
FILES = {
    '2017': '충청북도 지적통계 2017년 4분기.csv',
    '2018': '충청북도 지적통계_20181001.csv',
    '2019': '충청북도 지적통계_20190401.csv',
    '2020': '충청북도 지적통계_20200109.csv',
    # 2021: MISSING
    '2022': '충청북도_지적통계_20220701.xlsx',
    '2023': '충청북도_지적통계_20230701.csv',
    '2024': '충청북도_지적통계_20240701.csv',
    '2025': '충청북도_지적통계_20250630.csv'
}

# Standard Columns to Extract
TARGET_COLS = {
    'Region': ['토지소재명', '행정구역', '구분', '시군구'], # Potential names
    'Road': ['도로 면적', '도로'],
    'Factory': ['공장용지 면적', '공장용지', '공장'],
    'House': ['대 면적', '대'],
    'Farm': ['전 면적', '답 면적', '과수원 면적', '전', '답', '과수원'], # Need to sum
    'Forest': ['임야 면적', '임야'],
    'Total': ['계 면적', '면적', '총면적'] # Sometimes needs calculation
}

def clean_region_name(name):
    if not isinstance(name, str): return str(name)
    name = name.strip()
    # Normalize '청주 상당구' -> '청주시 상당구' or keep consistent
    # Target: '청주 상당구', '충주시' etc as per 2025 standard
    
    # Handle '청주시상당구' -> '청주 상당구'
    if '청주시' in name and ('구' in name) and len(name) > 3:
        return name.replace('청주시', '청주 ')
    if '청원군' in name: return '청주 청원구' # Old name check (though 2017 is post-integration)
    
    return name

def load_year_data(year, filename):
    path = os.path.join(DATA_DIR, filename)
    print(f"Loading {year}: {filename}")
    
    try:
        if filename.endswith('.xlsx'):
            df = pd.read_excel(path)
        else:
            # Try encodings
            try:
                df = pd.read_csv(path, encoding='cp949')
            except:
                df = pd.read_csv(path, encoding='euc-kr')
                
        # 1. Identify Region Column
        region_col = None
        for cand in TARGET_COLS['Region']:
            matches = [c for c in df.columns if cand in c]
            if matches:
                region_col = matches[0]
                break
        
        if not region_col:
            # Fallback: First column
            region_col = df.columns[0]
            print(f"  Warning: Region col not found, using {region_col}")
            
        df = df.rename(columns={region_col: 'Region'})
        df['Region'] = df['Region'].apply(clean_region_name)
        
        # Filter for summary rows if mixed (remove '계' or '합계')
        # Actually we need specific 14 regions.
        
        # 2. Extract Metrics
        # Helper to find column
        def get_col(candidates):
            for c in df.columns:
                for cand in candidates:
                    # Specific match preferred
                    if cand == c: return c
                    if cand in c and '지번수' not in c and '비율' not in c: return c
            return None

        # Build Data Dict
        data = []
        for idx, row in df.iterrows():
            region = row['Region']
            # Skip invalid regions (Sum rows etc)
            if '합계' in region or '총계' in region: continue
            
            # Extract Values
            def val(candidates):
                c = get_col(candidates)
                if c: return pd.to_numeric(row[c], errors='coerce')
                return 0
            
            # Farm is Sum
            farm_area = 0
            for f_type in ['전', '답', '과수원']:
                # Find col for specific farm type
                c = get_col([f_type + ' 면적', f_type])
                if c: farm_area += pd.to_numeric(row[c], errors='coerce')
            
            record = {
                'Year': year,
                'Region': region,
                'Road_Area': val(TARGET_COLS['Road']),
                'Factory_Area': val(TARGET_COLS['Factory']),
                'House_Area': val(TARGET_COLS['House']),
                'Forest_Area': val(TARGET_COLS['Forest']),
                'Farm_Area': farm_area
            }
            data.append(record)
            
        return pd.DataFrame(data)
        
    except Exception as e:
        print(f"  Error loading {year}: {e}")
        return pd.DataFrame()

def main():
    all_data = []
    
    # 1. Load Land Data (2017-2025 except 2021)
    for year, fname in FILES.items():
        if fname:
            df = load_year_data(year, fname)
            if not df.empty:
                all_data.append(df)
    
    if not all_data:
        print("No data loaded!")
        return
        
    full_df = pd.concat(all_data, ignore_index=True)
    
    # 2. Load Population Data (2017-2025)
    print("Loading Population...")
    pop_df = pd.read_csv(POP_FILE, encoding='cp949')
    # Reshape Pop: Wide to Long
    # Cols: Region, 2017, 2018...
    pop_long = []
    pop_years = [str(y) for y in range(2017, 2026)]
    
    region_map = {
        '상당구': '청주 상당구', '서원구': '청주 서원구', 
        '흥덕구': '청주 흥덕구', '청원구': '청주 청원구'
    }
    
    for idx, row in pop_df.iterrows():
        r_name = row['행정구역(시군구)별']
        r_name = region_map.get(r_name, r_name) # Fix names
        
        for y in pop_years:
            if y in pop_df.columns:
                pop_long.append({
                    'Year': y, # String to match Land
                    'Region': r_name,
                    'Population': row[y]
                })
    
    pop_final = pd.DataFrame(pop_long)
    
    # 3. Merge Land & Pop
    # Outer join to see what matches
    merged = pd.merge(full_df, pop_final, on=['Year', 'Region'], how='left')
    
    # Filter for core 14 regions
    core_regions = [
        '청주 상당구', '청주 서원구', '청주 흥덕구', '청주 청원구',
        '충주시', '제천시', '보은군', '옥천군', '영동군', 
        '증평군', '진천군', '괴산군', '음성군', '단양군'
    ]
    
    # Clean up Region names slightly if needed (remove extra spaces)
    merged['Region'] = merged['Region'].str.replace('  ', ' ').str.strip()
    
    final_df = merged[merged['Region'].isin(core_regions)].copy()
    
    # 4. Fill Missing 2021 (Interpolation)
    # Sort for interpolation
    final_df.sort_values(by=['Region', 'Year'], inplace=True)
    
    # We have gaps in Land Data for 2021. Pop data exists.
    # Group by Region and reindex to full year range to force NaNs for 2021 Land
    # Then interpolate
    
    df_list = []
    for region in core_regions:
        r_df = final_df[final_df['Region'] == region].copy()
        r_df.set_index('Year', inplace=True)
        # Reindex to all years
        all_y = [str(y) for y in range(2017, 2026)]
        r_df = r_df.reindex(all_y)
        r_df['Region'] = region # Restore region name
        r_df['Population'] = r_df['Population'].interpolate() # Pop shouldn't miss but good safety
        
        # Interpolate Land Columns
        land_cols = ['Road_Area', 'Factory_Area', 'House_Area', 'Forest_Area', 'Farm_Area']
        for c in land_cols:
            r_df[c] = r_df[c].interpolate(method='linear')
            
        r_df['Year'] = r_df.index
        df_list.append(r_df)
        
    interpolated_df = pd.concat(df_list, ignore_index=True)
    
    # Save
    out_path = os.path.join(RESULT_DIR, "chungbuk_timeseries_2017_2025.csv")
    interpolated_df.to_csv(out_path, index=False, encoding='utf-8-sig')
    print(f"Time Series Data Saved: {out_path}")
    print("Sample:\n", interpolated_df.head())

if __name__ == "__main__":
    main()
