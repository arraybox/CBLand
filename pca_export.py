
import os
import pandas as pd
from sklearn.decomposition import PCA

# print(os.getcwd())  # 현재 파이썬 작업 디렉토리
# print(os.listdir('.'))  # 현재 폴더 내 파일 목록

# 파일은 반드시 제공 데이터 파일(01_chungbuk_yearly_full_data.csv)을 사용할 것
df = pd.read_csv(os.path.join(os.getcwd(), '01_chungbuk_yearly_full_data.csv'))
cols = ['임야면적_비율','농경지면적_비율','대지면적_비율','공장용지면적_비율']

X = df[cols].values
pca = PCA(n_components=2)
X_pca = pca.fit_transform(X)

out = df[['연도','행정구역명']].copy()
out['PCA1'] = X_pca[:,0]
out['PCA2'] = X_pca[:,1]
out.to_csv(os.path.join(os.getcwd(), '10_chungbuk_landuse_PCA.csv'), index=False, encoding='utf-8-sig')
print('10_chungbuk_landuse_PCA.csv 파일이 생성되었습니다.')
