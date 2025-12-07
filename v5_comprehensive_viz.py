
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import folium
import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
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
    print("Font error: Using default")

# Coordinates (For Map)
COORDINATES = {
    '청주 상당구': [36.5746, 127.5252], '청주 서원구': [36.5786, 127.4597],
    '청주 흥덕구': [36.6346, 127.4278], '청주 청원구': [36.7118, 127.5255],
    '충주시': [36.9915, 127.9259], '제천시': [37.1326, 128.1909],
    '보은군': [36.4893, 127.7290], '옥천군': [36.3063, 127.5714],
    '영동군': [36.1748, 127.7769], '증평군': [36.7853, 127.5814],
    '진천군': [36.8553, 127.4355], '괴산군': [36.8153, 127.7865],
    '음성군': [36.9403, 127.6905], '단양군': [36.9845, 128.3655]
}

# Cluster Mapping
CLUSTER_MAP = {
    '청주 흥덕구': '도시/산업형', '청주 청원구': '도시/산업형',
    '진천군': '도시/산업형', '음성군': '도시/산업형',
    '청주 상당구': '농업/산림형', '충주시': '농업/산림형',
    '제천시': '농업/산림형', '보은군': '농업/산림형',
    '옥천군': '농업/산림형', '영동군': '농업/산림형',
    '괴산군': '농업/산림형', '단양군': '농업/산림형',
    '청주 서원구': '균형형', '증평군': '균형형'
}

COL_RENAME = {
    'Factory_Area': '공장용지면적', 'Forest_Area': '임야면적',
    'House_Area': '대지면적', 'Farm_Area': '농경지면적', 'Population': '인구수'
}

def load_prep():
    df = pd.read_csv(INPUT_FILE)
    df = df.rename(columns=COL_RENAME)
    df['Cluster'] = df['Region'].map(CLUSTER_MAP)
    return df

# --- Analysis 1: PCA (Why these clusters?) ---
def perform_pca(df):
    # Use 2025 data for static comparison of current state
    data_2025 = df[df['Year'] == 2025].copy()
    
    # Features for PCA: Ratios of land use
    data_2025['Total_Area'] = data_2025['공장용지면적'] + data_2025['임야면적'] + data_2025['대지면적'] + data_2025['농경지면적']
    features = ['공장용지비율', '임야비율', '대지비율', '농경지비율']
    
    for f_col, raw_col in zip(features, ['공장용지면적', '임야면적', '대지면적', '농경지면적']):
        data_2025[f_col] = data_2025[raw_col] / data_2025['Total_Area']
        
    x = data_2025[features].values
    x = StandardScaler().fit_transform(x)
    
    pca = PCA(n_components=2)
    principalComponents = pca.fit_transform(x)
    pca_df = pd.DataFrame(data=principalComponents, columns=['PC1', 'PC2'])
    final_pca = pd.concat([pca_df, data_2025[['Region', 'Cluster']].reset_index(drop=True)], axis=1)
    
    # Visualization
    plt.figure(figsize=(10, 8))
    sns.scatterplot(data=final_pca, x='PC1', y='PC2', hue='Cluster', style='Cluster', s=150)
    
    # Loading vectors (Arrows)
    loadings = pca.components_.T * np.sqrt(pca.explained_variance_)
    for i, feature in enumerate(features):
        plt.arrow(0, 0, loadings[i, 0]*3, loadings[i, 1]*3, color='r', alpha=0.5)
        plt.text(loadings[i, 0]*3.2, loadings[i, 1]*3.2, feature, color='red', ha='center')
        
    # Labels
    for i, row in final_pca.iterrows():
        plt.text(row['PC1']+0.1, row['PC2']+0.1, row['Region'], fontsize=9)
        
    plt.title('2025년 토지 이용 특성 PCA 분석 (군집 분류 근거)', fontsize=14)
    plt.xlabel('PC1 (도시/산업화 ↔ 농림보존)')
    plt.ylabel('PC2 (주거밀집 ↔ 산림)')
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, 'Viz_PCA_Analysis.png'), dpi=300)
    plt.close()

# --- Analysis 2: Heatmap (Change over time) ---
def create_heatmap(df):
    # X: Year, Y: Region, Value: Factory Area Growth Index (2017=100)
    pivoted = df.pivot(index='Region', columns='Year', values='공장용지면적')
    # Normalize by 2017 value
    normalized = pivoted.div(pivoted[2017], axis=0) * 100
    
    # Ordered by growth
    normalized['Total_Growth'] = normalized[2025] - 100
    normalized = normalized.sort_values('Total_Growth', ascending=False)
    normalized = normalized.drop(columns=['Total_Growth'])
    
    plt.figure(figsize=(12, 8))
    sns.heatmap(normalized, cmap='YlOrRd', annot=True, fmt='.0f', linewidths=.5)
    plt.title('시군구별 공장용지 성장 지수 히트맵 (2017년=100)', fontsize=14)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, 'Viz_Factory_Growth_Heatmap.png'), dpi=300)
    plt.close()

# --- Analysis 3: Donut Charts (Composition) ---
def create_donut_charts(df):
    data_2025 = df[df['Year'] == 2025]
    
    # Aggregate by Cluster
    cluster_sum = data_2025.groupby('Cluster')[['공장용지면적', '임야면적', '대지면적', '농경지면적']].sum()
    
    fig, axes = plt.subplots(1, 3, figsize=(18, 6))
    clusters = ['도시/산업형', '농업/산림형', '균형형']
    colors = ['#ff9999', '#66b3ff', '#99ff99', '#ffcc99'] # Factory, Forest, House, Farm
    labels = ['공장', '임야', '대지', '농지']
    
    for i, cluster in enumerate(clusters):
        if cluster not in cluster_sum.index: continue
        
        row = cluster_sum.loc[cluster]
        # Reorder to match labels
        values = [row['공장용지면적'], row['임야면적'], row['대지면적'], row['농경지면적']]
        
        axes[i].pie(values, labels=labels, autopct='%1.1f%%', startangle=90, colors=colors, pctdistance=0.85)
        # Draw circle
        centre_circle = plt.Circle((0,0),0.70,fc='white')
        axes[i].add_artist(centre_circle)
        axes[i].set_title(cluster)
        
    plt.suptitle('군집별 토지 이용 구성비 (2025)', fontsize=16)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, 'Viz_Cluster_Donut.png'), dpi=300)
    plt.close()

# --- Analysis 4: Map Visualization ---
def create_comprehensive_map(df):
    m = folium.Map(location=[36.8, 127.7], zoom_start=9)
    
    data_2025 = df[df['Year'] == 2025].copy()
    # Calculate Growth (2017-2025)
    data_2017 = df[df['Year'] == 2017].set_index('Region')
    data_2025 = data_2025.set_index('Region')
    
    data_2025['Pop_Growth_Rate'] = ((data_2025['인구수'] - data_2017['인구수']) / data_2017['인구수']) * 100
    data_2025.reset_index(inplace=True)
    
    for idx, row in data_2025.iterrows():
        region = row['Region']
        if region not in COORDINATES: continue
        
        coords = COORDINATES[region]
        cluster = row['Cluster']
        
        # Color by Cluster
        if cluster == '도시/산업형': color = 'red'
        elif cluster == '균형형': color = 'blue'
        else: color = 'green'
        
        html = f"""
        <div style="font-family: Malgun Gothic; width:250px">
            <h4>{region} ({cluster})</h4>
            <hr>
            <b>인구 증감:</b> {row['Pop_Growth_Rate']:.1f}%<br>
            <br>
            <b>토지 면적 (2025):</b><br>
            - 공장: {row['공장용지면적']:,.0f}<br>
            - 임야: {row['임야면적']:,.0f}<br>
            - 대지: {row['대지면적']:,.0f}<br>
        </div>
        """
        
        folium.CircleMarker(
            location=coords,
            radius=10 + (row['공장용지면적']/2000000), # Scale by Factory size
            popup=folium.Popup(html, max_width=300),
            color=color,
            fill=True,
            fill_opacity=0.6,
            tooltip=f"{region} ({cluster})"
        ).add_to(m)
        
    m.save(os.path.join(OUTPUT_DIR, 'Final_Chungbuk_Map.html'))

def main():
    print("Starting Comprehensive Viz...")
    df = load_prep()
    
    perform_pca(df)
    create_heatmap(df)
    create_donut_charts(df)
    create_comprehensive_map(df)
    
    print("All visual assets created.")

if __name__ == "__main__":
    main()
