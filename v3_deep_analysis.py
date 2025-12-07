
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import numpy as np
from matplotlib import font_manager, rc

# Config
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

def load_data():
    df = pd.read_csv(INPUT_FILE)
    df['Year'] = df['Year'].astype(int)
    return df

def generate_individual_analysis(df):
    """
    Generate detailed metrics for each region and save to CSV.
    Calculates: CAGR, Efficiency, Density, Correlation between Factory/Road and Pop.
    """
    regions = df['Region'].unique()
    results = []
    
    for r in regions:
        sub = df[df['Region'] == r].copy().sort_values('Year')
        
        # 1. CAGR (Compound Annual Growth Rate) calculation
        years = 8 # 2025 - 2017
        
        def get_cagr(col):
            start_val = sub[col].iloc[0]
            end_val = sub[col].iloc[-1]
            if start_val == 0: return 0
            return ((end_val / start_val) ** (1/years)) - 1

        # 2. Correlation Analysis (Does X drive Pop?)
        corr_fac_pop = sub['Factory_Area'].corr(sub['Population'])
        corr_road_pop = sub['Road_Area'].corr(sub['Population'])
        
        # 3. Efficiency Metric (Pop per Developed Land)
        # Developed Land = House + Factory
        sub['Developed_Area'] = sub['House_Area'] + sub['Factory_Area']
        sub['Eff_Index'] = sub['Population'] / sub['Developed_Area']
        
        eff_start = sub['Eff_Index'].iloc[0]
        eff_end = sub['Eff_Index'].iloc[-1]
        eff_change = ((eff_end - eff_start) / eff_start) * 100
        
        results.append({
            'Region': r,
            'Pop_CAGR': get_cagr('Population') * 100,
            'Factory_CAGR': get_cagr('Factory_Area') * 100,
            'Road_CAGR': get_cagr('Road_Area') * 100,
            'House_CAGR': get_cagr('House_Area') * 100,
            'Forest_Loss_Pct': ((sub['Forest_Area'].iloc[-1] - sub['Forest_Area'].iloc[0]) / sub['Forest_Area'].iloc[0]) * 100,
            'Corr_Factory_Pop': corr_fac_pop, # Correlation Index
            'Corr_Road_Pop': corr_road_pop,
            'Efficiency_Change_Pct': eff_change, # Did they use land more efficiently?
            'Startup_Factory_Area': sub['Factory_Area'].iloc[0],
            'Final_Factory_Area': sub['Factory_Area'].iloc[-1],
            'Startup_Pop': sub['Population'].iloc[0],
            'Final_Pop': sub['Population'].iloc[-1]
        })
        
    res_df = pd.DataFrame(results)
    res_df.to_csv(os.path.join(OUTPUT_DIR, 'detailed_region_metrics.csv'), index=False, encoding='utf-8-sig')
    return res_df

def create_advanced_visualizations(df, metrics_df):
    """
    Create high-quality comparative visualizations.
    """
    
    # 1. Heatmap of Industrial Growth Speed (Factory CAGR)
    # Sort by Factory Growth
    sorted_df = metrics_df.sort_values('Factory_CAGR', ascending=False)
    
    plt.figure(figsize=(12, 10))
    # We will use a Bar Chart for CAGR comparison instead of simple heatmap for better readability
    sns.barplot(data=sorted_df, y='Region', x='Factory_CAGR', palette='Blues_r')
    plt.title('시군구별 연평균 공장용지 증가율 (CAGR, %)', fontsize=15)
    plt.xlabel('연평균 증가율 (%)')
    plt.axvline(0, color='grey', linewidth=0.8)
    plt.grid(axis='x', alpha=0.3)
    
    # Add Value Labels
    for i, v in enumerate(sorted_df['Factory_CAGR']):
        plt.text(v + 0.1, i, f"{v:.1f}%", va='center')
        
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, 'Viz_Factory_Growth_Ranking.png'), dpi=300)
    plt.close()
    
    # 2. Scatter: Efficiency Change vs Population Growth
    # X: Pop Growth, Y: Efficiency Change
    # If Y is positive, Pop grew faster than land consumption (Smart Growth)
    # If Y is negative, Land consumption outpaced Pop (Sprawl)
    
    plt.figure(figsize=(10, 8))
    sns.scatterplot(data=metrics_df, x='Pop_CAGR', y='Efficiency_Change_Pct', 
                    size='Final_Pop', sizes=(50, 500), hue='Region', legend=False)
    
    # Quadrant Lines
    plt.axhline(0, color='red', linestyle='--', alpha=0.5)
    plt.axvline(0, color='red', linestyle='--', alpha=0.5)
    
    # Labels
    for i, row in metrics_df.iterrows():
        plt.text(row['Pop_CAGR'], row['Efficiency_Change_Pct'] + 0.2, 
                 row['Region'], fontsize=9, ha='center')
        
    plt.title('성장성 vs 효율성 매트릭스 (Efficiency Analysis)')
    plt.xlabel('연평균 인구 성장률 (CAGR, %)')
    plt.ylabel('토지 효율성 변화율 (%) (양수=고밀도, 음수=저밀도)')
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, 'Viz_Efficiency_Matrix.png'), dpi=300)
    plt.close()

    # 3. Stacked Area Chart for Land Use Change (Overall Province)
    # Aggregated by Year
    yearly_sum = df.groupby('Year')[['Road_Area', 'Factory_Area', 'House_Area']].sum().reset_index()
    
    plt.figure(figsize=(12, 6))
    plt.stackplot(yearly_sum['Year'], 
                  yearly_sum['House_Area'], 
                  yearly_sum['Factory_Area'], 
                  yearly_sum['Road_Area'],
                  labels=['대지(주거)', '공장용지', '도로'],
                  colors=['#ff9999', '#66b3ff', '#99ff99'], alpha=0.8)
    plt.legend(loc='upper left')
    plt.title('충청북도 전체 도시기반시설 면적 변화 추이 (Cumulative)')
    plt.ylabel('면적 합계 (m²)')
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, 'Viz_Total_LandUse_Change.png'), dpi=300)
    plt.close()
    
    # 4. Correlation Bar Chart (What drives Population?)
    # Compare Avg Correlation of Factory vs Pop and Road vs Pop across all regions
    avg_fac_corr = metrics_df['Corr_Factory_Pop'].mean()
    avg_road_corr = metrics_df['Corr_Road_Pop'].mean()
    
    # Also individual regions strong correlations
    # Let's plot the Correlation Coefficient for each region
    
    plt.figure(figsize=(12, 6))
    sns.barplot(data=metrics_df, x='Region', y='Corr_Factory_Pop', color='skyblue', label='공장-인구 상관계수')
    plt.axhline(0, color='black')
    plt.title('지역별 [공장용지 증가-인구 증가] 상관관계 강도')
    plt.ylabel('피어슨 상관계수 (1.0 = 완전 비례)')
    plt.ylim(-1, 1)
    plt.xticks(rotation=45)
    plt.grid(axis='y', alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, 'Viz_Correlation_Factory_Pop.png'), dpi=300)
    plt.close()

def main():
    print("Starting Deep Analysis...")
    df = load_data()
    
    # 1. Calculate Detailed Metrics
    metrics_df = generate_individual_analysis(df)
    print("Metrics calculated.")
    
    # 2. Generate Visualizations
    create_advanced_visualizations(df, metrics_df)
    print("Visualizations created.")
    
    # 3. Generate Summary Text for Report
    print("\n--- Insight Summary ---")
    top_growth = metrics_df.sort_values('Factory_CAGR', ascending=False).iloc[0]
    print(f"충북 최대 산업 성장 지역: {top_growth['Region']} (연평균 {top_growth['Factory_CAGR']:.2f}% 성장)")
    
    top_pop_corr = metrics_df.sort_values('Corr_Factory_Pop', ascending=False).iloc[0]
    print(f"산업-인구 연동성 최강 지역: {top_pop_corr['Region']} (상관계수 {top_pop_corr['Corr_Factory_Pop']:.2f})")

if __name__ == "__main__":
    main()
