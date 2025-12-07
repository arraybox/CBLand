
import pandas as pd
import folium
import os
import matplotlib.pyplot as plt
from matplotlib import font_manager, rc
import base64
from io import BytesIO

# --- 1. CONFIGURATION & SETUP ---

# File Paths
POP_FILE = r"D:\GITHUB\CBLand\data\행정구역_시군구_별_주민등록세대수_20251207114706.csv"
LAND_FILE = r"D:\GITHUB\CBLand\original_data\충청북도_지적통계_20250630.csv"
OUTPUT_HTML = r"D:\GITHUB\CBLand\result\new_comprehensive_analysis.html"
OUTPUT_README = r"D:\GITHUB\CBLand\README.md"

# Ensure result directory exists
os.makedirs(os.path.dirname(OUTPUT_HTML), exist_ok=True)

# Coordinates for Chungbuk regions (Approximate centroids)
COORDINATES = {
    '청주시 상당구': [36.5746, 127.5252],
    '청주시 서원구': [36.5786, 127.4597],
    '청주시 흥덕구': [36.6346, 127.4278],
    '청주시 청원구': [36.7118, 127.5255],
    '충주시': [36.9915, 127.9259],
    '제천시': [37.1326, 128.1909],
    '보은군': [36.4893, 127.7290],
    '옥천군': [36.3063, 127.5714],
    '영동군': [36.1748, 127.7769],
    '증평군': [36.7853, 127.5814],
    '진천군': [36.8553, 127.4355],
    '괴산군': [36.8153, 127.7865],
    '음성군': [36.9403, 127.6905],
    '단양군': [36.9845, 128.3655]
}

# --- 2. DATA LOADING & CLEANING ---

def load_data():
    # Load Population Data
    print("Loading Population Data...")
    pop_df = pd.read_csv(POP_FILE, encoding='cp949')
    # Filter for relevant columns: Region and Years
    # Expected columns: '행정구역(시군구)별', '2025', '2024', ... '2017'
    # Rename for clarity
    pop_df.rename(columns={'행정구역(시군구)별': 'Region'}, inplace=True)
    
    # Load Land Data
    print("Loading Land Data...")
    land_df = pd.read_csv(LAND_FILE, encoding='cp949')
    land_df.rename(columns={'토지소재명': 'Region_Land'}, inplace=True)
    
    return pop_df, land_df

def clean_and_merge(pop_df, land_df):
    print("Cleaning and Merging Data...")
    
    # 1. Standardize Region Names in Land Data to match Pop Data
    # Pop Data uses '청주시 상당구', Land Data uses '청주 상당구'
    def standardize_land_region(name):
        return name.replace('청주 ', '청주시 ')
    
    land_df['Region'] = land_df['Region_Land'].apply(standardize_land_region)
    
    # 2. Filter Population Data for relevant regions (Chungbuk specific)
    # Actually, just merge on Region.
    
    # 3. Merge
    merged_df = pd.merge(pop_df, land_df, on='Region', how='inner')
    
    return merged_df

# --- 3. ANALYSIS ---

def perform_analysis(df):
    print("Performing Analysis...")
    
    # Calculate Total Area (Sum of all land use types? Or just assume it covers most)
    # There isn't a "Total Area" column identified explicitly in the header list, 
    # so we should sum the key components or use the sum of all columns if they are exhaustive.
    # To be safe and accurate based on "Jijeok", we sum the area columns.
    # However, listing ALL area columns is tedious. 
    # Looking at the pattern: Columns 1, 3, 5, ... are Areas. 
    # Column 0 is Name.
    # Let's verify if there's a "Total" column. Usually column 1 in strict gov stats is Total.
    # Wait, looking at the header list from inspection:
    # 0: 토지소재명, 1: 전 면적...
    # There is NO explicitly named "Total" column in the captured 54 columns.
    # So we must sum them up.
    
    area_cols = [c for c in df.columns if '면적' in c]
    df['Total_Area'] = df[area_cols].sum(axis=1)
    
    # 1. Road Ratio
    df['Road_Area'] = df['도로 면적']
    df['Road_Ratio'] = (df['Road_Area'] / df['Total_Area']) * 100
    
    # 2. Land Use Ratios
    df['Residential_Area'] = df['대 면적']
    df['Industrial_Area'] = df['공장용지 면적']
    df['Farmland_Area'] = df['전 면적'] + df['답 면적'] + df['과수원 면적']
    df['Forest_Area'] = df['임야 면적']
    
    df['Residential_Ratio'] = (df['Residential_Area'] / df['Total_Area']) * 100
    df['Industrial_Ratio'] = (df['Industrial_Area'] / df['Total_Area']) * 100
    df['Farmland_Ratio'] = (df['Farmland_Area'] / df['Total_Area']) * 100
    df['Forest_Ratio'] = (df['Forest_Area'] / df['Total_Area']) * 100
    
    # 3. Population Trends
    # Calculate Change Rate (2017 -> 2025)
    # Ensure columns are numeric
    for col in ['2017', '2025']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
        
    df['Pop_Change_Rate'] = ((df['2025'] - df['2017']) / df['2017']) * 100
    
    return df

# --- 4. VISUALIZATION ---

def create_folium_map(df):
    print("Creating Map...")
    
    # Center map on Chungbuk
    m = folium.Map(location=[36.8, 127.7], zoom_start=9)
    
    # Create feature groups for layers (optional, but requested "Comprehensive")
    # We will put markers with popups containing detailed info
    
    for idx, row in df.iterrows():
        region = row['Region']
        if region not in COORDINATES:
            continue
            
        coords = COORDINATES[region]
        
        # Create Popup Content
        html = f"""
        <div style="width:300px">
            <h4>{region}</h4>
            <hr>
            <b>인구 (2025):</b> {row['2025']:,} 세대<br>
            <b>인구 증감 (2017-2025):</b> {row['Pop_Change_Rate']:.2f}%<br>
            <br>
            <b>총 면적:</b> {row['Total_Area']:,.0f} m²<br>
            <br>
            <b>토지 이용 비율:</b>
            <ul>
                <li>도로: {row['Road_Ratio']:.2f}% ({row['Road_Area']:,.0f} m²)</li>
                <li>대지(주거): {row['Residential_Ratio']:.2f}%</li>
                <li>산업(공장): {row['Industrial_Ratio']:.2f}%</li>
                <li>농지: {row['Farmland_Ratio']:.2f}%</li>
                <li>임야: {row['Forest_Ratio']:.2f}%</li>
            </ul>
        </div>
        """
        
        popup = folium.Popup(html, max_width=350)
        
        # Color marker based on Population Change
        color = 'blue' if row['Pop_Change_Rate'] > 0 else 'red'
        
        # Circle Marker for Population Size
        folium.CircleMarker(
            location=coords,
            radius=10 + (row['2025'] / 50000), # Size scaling
            popup=popup,
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.6,
            tooltip=f"{region}: {row['2025']:,} 세대"
        ).add_to(m)
        
    return m

# --- 5. MAIN EXECUTION ---

def main():
    try:
        pop_df, land_df = load_data()
        merged_df = clean_and_merge(pop_df, land_df)
        final_df = perform_analysis(merged_df)
        
        # Export Analysis to CSV
        output_csv = r"D:\GITHUB\CBLand\result\final_comprehensive_analysis.csv"
        final_df.to_csv(output_csv, index=False, encoding='utf-8-sig')
        print(f"Analysis saved to {output_csv}")
        
        # Create Map
        m = create_folium_map(final_df)
        m.save(OUTPUT_HTML)
        print(f"Map saved to {OUTPUT_HTML}")
        
    except Exception as e:
        print(f"An error occurred: {e}")
        raise e

if __name__ == "__main__":
    main()
