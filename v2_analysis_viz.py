
import pandas as pd
import matplotlib.pyplot as plt
import os
import seaborn as sns
from matplotlib import font_manager, rc

# Config
input_file = r"D:\GITHUB\CBLand\result_v2\chungbuk_timeseries_2017_2025.csv"
output_dir = r"D:\GITHUB\CBLand\result_v2"

# Font Setup
try:
    font_path = "C:/Windows/Fonts/malgun.ttf"
    font_name = font_manager.FontProperties(fname=font_path).get_name()
    rc('font', family=font_name)
    plt.rcParams['axes.unicode_minus'] = False
except:
    pass

def load_data():
    df = pd.read_csv(input_file)
    # Ensure Year is int for plotting
    df['Year'] = df['Year'].astype(int)
    return df

def analyze_changes(df):
    # Calculate Changes (2025 - 2017)
    df_start = df[df['Year'] == 2017].set_index('Region')
    df_end = df[df['Year'] == 2025].set_index('Region')
    
    metrics = ['Road_Area', 'Factory_Area', 'House_Area', 'Farm_Area', 'Forest_Area', 'Population']
    changes = pd.DataFrame()
    
    for m in metrics:
        # Absolute Change
        changes[f'{m}_Delta'] = df_end[m] - df_start[m]
        # Pct Change
        changes[f'{m}_Pct'] = ((df_end[m] - df_start[m]) / df_start[m]) * 100
        
    changes.reset_index(inplace=True)
    
    # Save Classification Data
    changes.to_csv(os.path.join(output_dir, 'region_change_metrics_2017_2025.csv'), index=False, encoding='utf-8-sig')
    return changes

def classify_regions(changes):
    # Cluster Logic based on 'Factory_Area_Pct' and 'Population_Pct'
    # 1. Industrial Growth: Factory > 20% & Pop > 5%
    # 2. Urban Center: House > 5% & Pop > 0
    # 3. Rural/Forest: Forest_Delta approx 0 (preserved) or Low Pop Growth
    # Simple manual classification for report
    
    def get_type(row):
        fac = row['Factory_Area_Pct']
        pop = row['Population_Pct']
        house = row['House_Area_Pct']
        
        if fac > 15 and pop > 5:
            return "산업성장형"
        elif house > 3 and pop > -3:
            return "도시확장형" 
        elif pop < -5:
            return "인구소멸위험형"
        else:
            return "농림보존/정체형"
            
    changes['Type'] = changes.apply(get_type, axis=1)
    
    # Save with Type
    changes.to_csv(os.path.join(output_dir, 'region_classification.csv'), index=False, encoding='utf-8-sig')
    return changes

def plot_trends(df):
    regions = df['Region'].unique()
    
    # 1. Factory Area Trend
    plt.figure(figsize=(12, 6))
    sns.lineplot(data=df, x='Year', y='Factory_Area', hue='Region', marker='o')
    plt.title('연도별 공장용지 면적 변화 (2017-2025)')
    plt.ylabel('면적 (m²)')
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'Trend_Factory.png'))
    plt.close()
    
    # 2. Population Trend
    plt.figure(figsize=(12, 6))
    sns.lineplot(data=df, x='Year', y='Population', hue='Region', marker='o')
    plt.title('연도별 인구 변화 (2017-2025)')
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'Trend_Population.png'))
    plt.close()

def plot_scatter(changes):
    plt.figure(figsize=(10, 8))
    
    # Scatter: Factory Growth (%) vs Population Growth (%)
    sns.scatterplot(data=changes, x='Factory_Area_Pct', y='Population_Pct', 
                    hue='Type', s=100, style='Type')
    
    # Add textual labels
    for i, row in changes.iterrows():
        plt.text(row['Factory_Area_Pct']+0.5, row['Population_Pct']+0.5, 
                 row['Region'], fontsize=9)
        
    plt.axhline(0, color='grey', linestyle='--')
    plt.axvline(0, color='grey', linestyle='--')
    plt.title('공장용지 증가율 vs 인구 증가율 상관관계 매트릭스')
    plt.xlabel('공장용지 증가율 (%)')
    plt.ylabel('인구 증가율 (%)')
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'Scatter_Growth_Matrix.png'))
    plt.close()

def main():
    df = load_data()
    changes = analyze_changes(df)
    changes = classify_regions(changes)
    
    plot_trends(df)
    plot_scatter(changes)
    print("Analysis & Visualization Complete.")

if __name__ == "__main__":
    main()
