
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import numpy as np
from matplotlib import font_manager, rc

# 1. Config & Setup
INPUT_FILE = r"D:\GITHUB\CBLand\result_v2\chungbuk_timeseries_2017_2025.csv"
OUTPUT_DIR = r"D:\GITHUB\CBLand\result_v2"

# Font Setup
try:
    font_path = "C:/Windows/Fonts/malgun.ttf"
    font_name = font_manager.FontProperties(fname=font_path).get_name()
    rc('font', family=font_name)
    plt.rcParams['axes.unicode_minus'] = False
except:
    pass

# Defined Clusters
CLUSTER_MAP = {
    # Urban/Industrial
    '청주 흥덕구': '도시/산업형',
    '청주 청원구': '도시/산업형',
    '진천군': '도시/산업형',
    '음성군': '도시/산업형',
    
    # Agricultural/Forest
    '청주 상당구': '농업/산림형',
    '충주시': '농업/산림형',
    '제천시': '농업/산림형',
    '보은군': '농업/산림형',
    '옥천군': '농업/산림형',
    '영동군': '농업/산림형',
    '괴산군': '농업/산림형',
    '단양군': '농업/산림형',
    
    # Balanced
    '청주 서원구': '균형형',
    '증평군': '균형형'
}

# Column Renaming Map
COL_RENAME = {
    'Factory_Area': '공장용지면적',
    'Forest_Area': '임야면적',
    'House_Area': '대지면적',
    'Farm_Area': '농경지면적',
    'Population': '인구수'
}

def load_and_prep():
    df = pd.read_csv(INPUT_FILE)
    df['Year'] = df['Year'].astype(int)
    
    # Rename Columns
    df = df.rename(columns=COL_RENAME)
    
    # Map Clusters
    df['Cluster'] = df['Region'].map(CLUSTER_MAP)
    
    # Validate mapping coverage
    missing = df[df['Cluster'].isna()]['Region'].unique()
    if len(missing) > 0:
        print(f"Warning: Regions not mapped to clusters: {missing}")
        
    return df

def generate_cluster_stats(df):
    # Aggregated by Cluster & Year
    # Sum the areas/population for the group
    group_cols = ['Cluster', 'Year']
    target_cols = ['공장용지면적', '임야면적', '대지면적', '농경지면적', '인구수']
    
    grouped = df.groupby(group_cols)[target_cols].sum().reset_index()
    
    # Calculate Ratios per Cluster (e.g., Factory Area %)
    # We need Total Area for each cluster-year to get ratio
    # Assuming these 4 types cover most developed/undeveloped land, but not all (Roads etc are missing in target_cols list but exist in df)
    # Let's sum just these 4 for the "Analyzed Land Area"
    grouped['분석토지합계'] = grouped['공장용지면적'] + grouped['임야면적'] + grouped['대지면적'] + grouped['농경지면적']
    
    for col in ['공장용지면적', '임야면적', '대지면적', '농경지면적']:
        grouped[f'{col}_비율'] = (grouped[col] / grouped['분석토지합계']) * 100
        
    # Save Grouped Data
    grouped.to_csv(os.path.join(OUTPUT_DIR, 'cluster_timeseries_stats.csv'), index=False, encoding='utf-8-sig')
    return grouped

def visualize_clusters(df, grouped):
    # 1. Cluster Composition (2025 Snapshot)
    snapshot = grouped[grouped['Year'] == 2025].copy()
    
    # Plot Stacked Bar for Ratios
    ratio_cols = ['공장용지면적_비율', '대지면적_비율', '농경지면적_비율', '임야면적_비율']
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#9467bd'] # Blue, Orange, Green, Purple
    labels = ['공장용지(산업)', '대지(주거)', '농경지(농업)', '임야(산림)']
    
    plt.figure(figsize=(10, 6))
    bottom = np.zeros(len(snapshot))
    
    for i, col in enumerate(ratio_cols):
        plt.bar(snapshot['Cluster'], snapshot[col], bottom=bottom, label=labels[i], color=colors[i], width=0.5)
        bottom += snapshot[col].values

    plt.title('2025년 군집별 토지 이용 구성 비율', fontsize=15)
    plt.ylabel('비율 (%)')
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, 'Viz_Cluster_Composition_2025.png'), dpi=300)
    plt.close()
    
    # 2. Trend Analysis (Factory Growth Normalized)
    # Index 2017 = 100
    pivoted = grouped.pivot(index='Year', columns='Cluster', values='공장용지면적')
    normalized = pivoted.div(pivoted.iloc[0]) * 100
    
    plt.figure(figsize=(10, 6))
    sns.lineplot(data=normalized, dashes=False, markers=True, linewidth=2.5)
    plt.title('군집별 공장용지 면적 성장 추이 (2017년=100)', fontsize=15)
    plt.ylabel('성장 지수 (2017=100)')
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, 'Viz_Cluster_Factory_Trend.png'), dpi=300)
    plt.close()

    # 3. Population Trend by Cluster
    # Index 2017 = 100
    pop_piv = grouped.pivot(index='Year', columns='Cluster', values='인구수')
    pop_norm = pop_piv.div(pop_piv.iloc[0]) * 100
    
    plt.figure(figsize=(10, 6))
    sns.lineplot(data=pop_norm, dashes=False, markers=True, linewidth=2.5)
    plt.title('군집별 인구 변화 추이 (2017년=100)', fontsize=15)
    plt.ylabel('인구 지수 (2017=100)')
    plt.axhline(100, color='grey', linestyle='--')
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, 'Viz_Cluster_Pop_Trend.png'), dpi=300)
    plt.close()

def main():
    print("Starting Cluster Analysis...")
    df = load_and_prep()
    
    grouped = generate_cluster_stats(df)
    print("Stats Generated.")
    
    visualize_clusters(df, grouped)
    print("Visualization Complete.")
    
    # Quick Text Summary
    print("\n--- Cluster Summary (2025) ---")
    print(grouped[grouped['Year'] == 2025][['Cluster', '공장용지면적', '인구수']])

if __name__ == "__main__":
    main()
