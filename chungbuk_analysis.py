import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.decomposition import PCA

plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

def plot_donut_by_year(df, year, outname='donut_grid.png'):
    data = df[df['연도']==year]
    areas = data['행정구역명']
    vals = data[['임야면적_비율','농경지면적_비율','대지면적_비율','공장용지면적_비율']].values
    labels = ['임야','농경지','대지','공장']
    colors = ['#77C478','#F3D670','#54A6D8','#E67C73']
    grid = (4,4)
    fig, axes = plt.subplots(grid[0],grid[1],figsize=(18,14))
    axes = axes.flatten()
    def _pct_format(val):
        return '' if val < 8 else f'{val:.1f}%'
    for i, (ax, area, v) in enumerate(zip(axes,areas,vals)):
        wedges, texts, autotexts = ax.pie(
            v, labels=None, colors=colors, startangle=90,
            autopct=_pct_format, pctdistance=1.07, labeldistance=1.18, textprops={'fontsize':10})
        for j, wedge in enumerate(wedges):
            if v[j] >= 8:
                ang = (wedge.theta2 + wedge.theta1)/2.
                x = np.cos(np.deg2rad(ang)) * 1.27
                y = np.sin(np.deg2rad(ang)) * 1.27
                ax.text(x, y, labels[j], ha='center', va='center', fontsize=11, weight='bold')
        plt.setp(wedges, width=0.36)
        ax.set_title(str(area),fontsize=12)
        ax.axis('equal')
    for j in range(len(areas), len(axes)):
        axes[j].axis('off')
    plt.suptitle(f"{year}년 시군구별 토지 용도비율(도넛차트)", fontsize=18)
    plt.tight_layout(rect=[0,0,1,0.97])
    plt.savefig(outname, dpi=140)
    plt.show()

def run_pca_plot(df, outname='PCA_산점도_최소겹침.png'):
    cols=['임야면적_비율','농경지면적_비율','대지면적_비율','공장용지면적_비율']
    X = df[cols].values
    pca = PCA(n_components=2)
    X_pca = pca.fit_transform(X)
    result = df[['연도','행정구역명']].copy()
    result['PCA1'] = X_pca[:,0]
    result['PCA2'] = X_pca[:,1]
    print('PCA 분산 설명력:', pca.explained_variance_ratio_)

    plt.figure(figsize=(9,7))
    color_map = {2017:'tab:blue', 2018:'tab:orange', 2019:'tab:green', 2020:'tab:red',
                 2023:'tab:purple', 2024:'tab:brown', 2025:'tab:gray'}
    for y in sorted(result['연도'].unique()):
        idx = result['연도']==y
        plt.scatter(result.loc[idx,'PCA1'], result.loc[idx,'PCA2'], label=f'{y}', alpha=0.65, s=36)
    shown = set()
    for y in sorted(result['연도'].unique()):
        part = result[result['연도']==y]
        for row in pd.concat([part.nlargest(1,'PCA1'), part.nsmallest(1,'PCA1'),
                              part.nlargest(1,'PCA2'), part.nsmallest(1,'PCA2')]).itertuples():
            if (row.행정구역명, row.연도) not in shown:
                plt.text(row.PCA1,row.PCA2, row.행정구역명, fontsize=11, alpha=0.65,
                        bbox=dict(boxstyle='round,pad=0.16', fc='white', ec='0.7', alpha=0.04))
                shown.add((row.행정구역명, row.연도))
    plt.xlabel('PCA1')
    plt.ylabel('PCA2')
    plt.title('충북 시군구 토지 용도비율 PCA (라벨 최소 겹침)')
    plt.legend()
    plt.tight_layout()
    plt.savefig(outname, dpi=160)
    plt.show()

def plot_trend_summary(df_sum, outname='시계열추이.png'):
    fig,ax = plt.subplots(figsize=(9,5))
    x = df_sum['연도']
    col_list = [
        ('임야면적_비율', '임야', '#77C478'),
        ('농경지면적_비율', '농경지', '#F3D670'),
        ('대지면적_비율', '대지', '#54A6D8'),
        ('공장용지면적_비율', '공장', '#E67C73')
    ]
    for cname, label, color in col_list:
        ax.plot(x, df_sum[cname], marker='o', label=label, color=color)
        # 각 점 위에 비율 값 표기 (소수점 2자리, 위쪽)
        for xi, yi in zip(x, df_sum[cname]):
            # 점 위(아래는 va='top'), 옆 정렬 조정
            ax.text(xi, yi+0.75, f"{yi:.2f}", fontsize=10, color=color, ha='center', va='bottom', fontweight='bold')
    ax.legend()
    ax.set_xlabel('연도')
    ax.set_ylabel('비율(%)')
    ax.set_title('충북 전체 토지 용도비율 시계열')
    plt.tight_layout()
    plt.savefig(outname, dpi=130)
    plt.show()

def show_extreme_changes(df_ext):
    print(df_ext.head(8))

def plot_heatmap_by_area(df, outname='heatmap_지역별연도별비율.png'):
    # 데이터 피벗 (index=행정구역명, columns=연도, values=특정 용도비율)
    landuses = ['임야면적_비율','농경지면적_비율','대지면적_비율','공장용지면적_비율']
    for col in landuses:
        pivot = df.pivot(index='행정구역명', columns='연도', values=col)
        plt.figure(figsize=(10,6))
        # 색상은 예시: 녹-노랑-빨 강약(필요시 cmap 자유조정)
        ax = sns.heatmap(pivot, annot=True, fmt=".1f",
                         cmap='YlOrRd', cbar_kws={'label':'% 비율'})
        plt.title(f'연도별 {col[:-3]} 비율 히트맵')
        plt.tight_layout()
        plt.savefig(f"{outname.split('.')[0]}_{col[:-3]}.png", dpi=140)
        plt.show()

def plot_heatmap_trend(df_sum, outname='heatmap_도전체추이.png'):
    # df_sum: 08_chungbuk_yearly_use_trend.csv 사용
    pivot = df_sum[['연도','임야면적_비율','농경지면적_비율','대지면적_비율','공장용지면적_비율']]
    pivot = pivot.set_index('연도').T
    plt.figure(figsize=(8,2.5))
    ax = sns.heatmap(pivot, annot=True, fmt=".2f",
                   cmap='YlGnBu', cbar_kws={'label':'% 비율'})
    plt.title('충청북도 전체 용도별 시계열 히트맵')
    plt.ylabel('토지 용도')
    plt.xlabel('연도')
    plt.tight_layout()
    plt.savefig(outname, dpi=140)
    plt.show()

if __name__ == "__main__":
    df = pd.read_csv(os.path.join(os.getcwd(), 'data', '01_chungbuk_yearly_full_data.csv'))
    df_sum = pd.read_csv(os.path.join(os.getcwd(), 'data', '08_chungbuk_yearly_use_trend.csv'))
    df_ext = pd.read_csv(os.path.join(os.getcwd(), 'data', '09_chungbuk_extreme_change_areas.csv'))

    plot_donut_by_year(df, 2025, outname='도넛차트_2025.png')
    run_pca_plot(df, outname='PCA_격자라벨.png')
    plot_trend_summary(df_sum, outname='연도별시계열.png')
    plot_heatmap_by_area(df)
    plot_heatmap_trend(df_sum)
    show_extreme_changes(df_ext)
