
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
import folium
from matplotlib import font_manager, rc

# --- 1. SETUP ---
pd.set_option('display.max_columns', None)

# Paths
POP_FILE = r"D:\GITHUB\CBLand\data\행정구역_시군구_별_주민등록세대수_20251207114706.csv"
LAND_FILE = r"D:\GITHUB\CBLand\original_data\충청북도_지적통계_20250630.csv"
RESULT_DIR = r"D:\GITHUB\CBLand\result"
os.makedirs(RESULT_DIR, exist_ok=True)

# Font Setup for Windows
try:
    font_path = "C:/Windows/Fonts/malgun.ttf"
    font_name = font_manager.FontProperties(fname=font_path).get_name()
    rc('font', family=font_name)
    plt.rcParams['axes.unicode_minus'] = False
except:
    print("Korean font not found, using default.")

# Coordinates (Updated keys to match Land File format: '청주 상당구' etc.)
COORDINATES = {
    '청주 상당구': [36.5746, 127.5252],
    '청주 서원구': [36.5786, 127.4597],
    '청주 흥덕구': [36.6346, 127.4278],
    '청주 청원구': [36.7118, 127.5255],
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

# --- 2. DATA PROCESSING ---

def load_and_merge():
    # Load
    pop = pd.read_csv(POP_FILE, encoding='cp949')
    land = pd.read_csv(LAND_FILE, encoding='cp949')
    
    # Check loaded basic info
    print(f"Pop Rows: {len(pop)}, Land Rows: {len(land)}")
    
    # Rename columns for easier access
    pop.rename(columns={'행정구역(시군구)별': 'Pop_Region'}, inplace=True)
    land.rename(columns={'토지소재명': 'Land_Region'}, inplace=True)
    
    # Clean Pop Data
    # Mapping Dictionary for Pop_Region -> Land_Region style
    pop_mapping = {
        '상당구': '청주 상당구',
        '서원구': '청주 서원구',
        '흥덕구': '청주 흥덕구',
        '청원구': '청주 청원구'
    }
    
    # Apply mapping: If key exists in mapping, use it. Else use original.
    pop['Region_Key'] = pop['Pop_Region'].apply(lambda x: pop_mapping.get(x, x))
    
    # Filter Pop Data to include only the 14 targets
    target_regions = list(COORDINATES.keys())
    pop = pop[pop['Region_Key'].isin(target_regions)].copy()
    
    # Clean Land Data
    # Land regions should mostly match target_regions exactly
    land = land[land['Land_Region'].isin(target_regions)].copy()
    
    # Merge
    merged = pd.merge(pop, land, left_on='Region_Key', right_on='Land_Region', how='inner')
    
    print(f"Merged Rows: {len(merged)}")
    if len(merged) != 14:
        missing = set(target_regions) - set(merged['Region_Key'])
        print(f"CRITICAL WARNING: Missing regions: {missing}")
    
    return merged

def calculate_metrics(df):
    # Numeric Conversion for Pop
    cols_to_num = ['2017', '2025']
    for c in cols_to_num:
        df[c] = pd.to_numeric(df[c], errors='coerce')
        
    df['Pop_Change_Rate'] = ((df['2025'] - df['2017']) / df['2017']) * 100
    
    # Land Area Calculations
    # Based on inspected headers:
    # 25: 도로 면적 (Column name might be '도로 면적')
    # 13: 대 면적
    # 15: 공장용지 면적
    # 01: 전 면적, 03: 답 면적, 05: 과수원 면적
    # 09: 임야 면적
    
    # However, let's use the explicit column names which we saw in the text file earlier.
    # We must trust pandas column matching.
    
    # Total Area Estimate: Sum of all numeric columns in Land part is risky.
    # It's better to calculate Total as Sum of the specific components if they are dominant, 
    # OR sum all '.. 면적' columns.
    area_cols = [c for c in df.columns if '면적' in c]
    df['Total_Area'] = df[area_cols].sum(axis=1)
    
    # Specific Areas
    df['Road_Area'] = df['도로 면적']
    df['Res_Area'] = df['대 면적']
    df['Ind_Area'] = df['공장용지 면적']
    df['Farm_Area'] = df['전 면적'] + df['답 면적'] + df['과수원 면적']
    df['Forest_Area'] = df['임야 면적']
    
    # Ratios
    df['Road_Ratio'] = (df['Road_Area'] / df['Total_Area']) * 100
    df['Res_Ratio'] = (df['Res_Area'] / df['Total_Area']) * 100
    df['Ind_Ratio'] = (df['Ind_Area'] / df['Total_Area']) * 100
    df['Farm_Ratio'] = (df['Farm_Area'] / df['Total_Area']) * 100
    df['Forest_Ratio'] = (df['Forest_Area'] / df['Total_Area']) * 100
    
    return df

# --- 3. ANALYSIS & VISUALIZATION ---

def perform_correlation_analysis(df):
    # Selected Variables
    analysis_vars = [
        'Pop_Change_Rate', 
        'Road_Ratio', 
        'Res_Ratio', 
        'Ind_Ratio', 
        'Farm_Ratio', 
        'Forest_Ratio'
    ]
    
    # Rename for cleaner plot labels
    rename_map = {
        'Pop_Change_Rate': '인구증감률',
        'Road_Ratio': '도로율',
        'Res_Ratio': '대지비율',
        'Ind_Ratio': '공장용지비율',
        'Farm_Ratio': '농지비율',
        'Forest_Ratio': '임야비율'
    }
    
    corr_df = df[analysis_vars].rename(columns=rename_map)
    corr_matrix = corr_df.corr()
    
    # Save Matrix
    corr_matrix.to_csv(os.path.join(RESULT_DIR, 'integrated_correlation_matrix.csv'), encoding='utf-8-sig')
    print("Correlation matrix saved.")
    
    # 1. Heatmap
    plt.figure(figsize=(10, 8))
    sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', fmt=".2f", linewidths=.5)
    plt.title('충청북도 인구-토지-인프라 통합 상관관계 분석')
    plt.tight_layout()
    plt.savefig(os.path.join(RESULT_DIR, 'correlation_heatmap.png'), dpi=300)
    plt.close()
    
    # 2. Scatter Plots (Key Insights)
    # Figure with subplots
    fig, axes = plt.subplots(1, 3, figsize=(18, 5))
    
    # Plot 1: Road vs Pop Change
    sns.regplot(x='Road_Ratio', y='Pop_Change_Rate', data=df, ax=axes[0], color='blue')
    axes[0].set_title('도로율과 인구증감률의 관계')
    axes[0].set_xlabel('도로율 (%)')
    axes[0].set_ylabel('인구증감률 (2017-2025, %)')
    
    # Plot 2: Industry vs Pop Change
    sns.regplot(x='Ind_Ratio', y='Pop_Change_Rate', data=df, ax=axes[1], color='red')
    axes[1].set_title('공장용지비율과 인구증감률의 관계')
    axes[1].set_xlabel('공장용지비율 (%)')
    axes[1].set_ylabel('인구증감률 (%)')
    
    # Plot 3: Forest vs Pop Change (Negative correlation expected)
    sns.regplot(x='Forest_Ratio', y='Pop_Change_Rate', data=df, ax=axes[2], color='green')
    axes[2].set_title('임야비율과 인구증감률의 관계')
    axes[2].set_xlabel('임야비율 (%)')
    axes[2].set_ylabel('인구증감률 (%)')
    
    # Label Points
    for i, row in df.iterrows():
        # Label specifically for Plot 1
        axes[0].text(row['Road_Ratio'], row['Pop_Change_Rate'], row['Region_Key'], 
                     fontsize=8, alpha=0.7)
        axes[1].text(row['Ind_Ratio'], row['Pop_Change_Rate'], row['Region_Key'], 
                     fontsize=8, alpha=0.7)
        axes[2].text(row['Forest_Ratio'], row['Pop_Change_Rate'], row['Region_Key'], 
                     fontsize=8, alpha=0.7)
    
    plt.tight_layout()
    plt.savefig(os.path.join(RESULT_DIR, 'scatter_pop_road_industry.png'), dpi=300)
    plt.close()

def create_folium_map(df):
    m = folium.Map(location=[36.8, 127.7], zoom_start=9)
    
    for idx, row in df.iterrows():
        region = row['Region_Key']
        if region not in COORDINATES:
            continue
        
        coords = COORDINATES[region]
        
        # Color Logic
        pop_rate = row['Pop_Change_Rate']
        if pop_rate >= 10: color = 'darkblue'
        elif pop_rate > 0: color = 'blue'
        elif pop_rate > -5: color = 'orange'
        else: color = 'red'
        
        html = f"""
        <div style="font-family: Malgun Gothic; width:280px">
            <h4>{region}</h4>
            <hr>
            <b>인구 (2025):</b> {int(row['2025']):,}세대<br>
            <b>인구 증감 (8년):</b> <span style="color:{color}">{pop_rate:.2f}%</span><br>
            <br>
            <b>핵심 지표:</b>
            <ul>
                <li>도로율: {row['Road_Ratio']:.2f}%</li>
                <li>공장용지: {row['Ind_Ratio']:.2f}%</li>
                <li>대지(주거): {row['Res_Ratio']:.2f}%</li>
            </ul>
        </div>
        """
        
        folium.CircleMarker(
            location=coords,
            radius=12 + (pop_rate/2 if pop_rate > 0 else 0), # Size bonus for growth
            popup=folium.Popup(html, max_width=300),
            color=color,
            fill=True,
            fill_opacity=0.7,
            tooltip=f"{region} ({pop_rate:.1f}%)"
        ).add_to(m)
        
    m.save(os.path.join(RESULT_DIR, 'chungbuk_integrated_map.html'))
    print("Map saved.")

# --- 4. MAIN ---

def main():
    print("Starting Comprehensive Analysis...")
    df = load_and_merge()
    df = calculate_metrics(df)
    
    # Save the CLEAN Data
    out_csv = os.path.join(RESULT_DIR, 'final_clean_14regions.csv')
    df.to_csv(out_csv, index=False, encoding='utf-8-sig')
    print(f"Clean Data Saved: {out_csv}")
    
    perform_correlation_analysis(df)
    create_folium_map(df)
    print("All tasks completed.")

if __name__ == "__main__":
    main()
