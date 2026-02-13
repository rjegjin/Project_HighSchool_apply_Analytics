import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# 한글 폰트 설정 (Linux 환경 대응)
plt.rcParams['font.family'] = 'NanumGothic' if os.path.exists('/usr/share/fonts/truetype/nanum/NanumGothic.ttf') else 'DejaVu Sans'
plt.rcParams['axes.unicode_minus'] = False

BASE_DIR = "Project_HighSchool_apply_Analytics"
INPUT_EXCEL = os.path.join(BASE_DIR, "data", "processed", "Step4_대학원수준_심층분석.xlsx")
OUTPUT_DIR = os.path.join(BASE_DIR, "output", "advanced_plots")

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

def visualize_results():
    print("🎨 고급 통계 지표 시각화를 시작합니다...")
    xls = pd.ExcelFile(INPUT_EXCEL)
    
    # 1. PCA & GMM Clustering Scatter Plot
    df_school = pd.read_excel(xls, '1_학교_고급유형화')
    plt.figure(figsize=(10, 7))
    sns.scatterplot(data=df_school, x='PCA_1', y='PCA_2', hue='GMM_Cluster', palette='viridis', s=100, alpha=0.7)
    
    # 학교 이름 라벨링 (일부 핵심 학교만)
    for i, row in df_school.iterrows():
        if abs(row['PCA_1']) > 1.5 or abs(row['PCA_2']) > 1.5:
            plt.text(row['PCA_1'], row['PCA_2'], row['배정고등학교'], fontsize=9)
            
    plt.title('PCA-GMM 기반 학교 유형 다차원 분석')
    plt.xlabel('PC1: 학교 규모 및 인지도 지표')
    plt.ylabel('PC2: 선호도 및 만족도 지표')
    plt.grid(True, linestyle='--', alpha=0.6)
    plt.savefig(os.path.join(OUTPUT_DIR, '1_PCA_GMM_Cluster.png'))
    plt.close()

    # 2. 지역별 엔트로피 (배정 다양성) - 하위 10개 (쏠림 지역)
    df_dong = pd.read_excel(xls, '2_지역_배정다양성')
    plt.figure(figsize=(12, 6))
    df_dong_sorted = df_dong.sort_values('엔트로피_지수', ascending=True).head(10)
    sns.barplot(data=df_dong_sorted, x='엔트로피_지수', y='행정동', palette='Reds_r')
    plt.title('지역별 배정 엔트로피 (지수가 낮을수록 특정 학교 쏠림 강함)')
    plt.savefig(os.path.join(OUTPUT_DIR, '2_Dong_Entropy_Top10.png'))
    plt.close()

    # 3. 네트워크 중심성 Top 10
    df_centrality = pd.read_excel(xls, '3_네트워크_중심성')
    plt.figure(figsize=(12, 6))
    df_top_centrality = df_centrality.sort_values('중심성_지수', ascending=False).head(10)
    sns.barplot(data=df_top_centrality, x='중심성_지수', y='ID', palette='magma')
    plt.title('네트워크 중심성 지수 (배정 흐름의 허브 역할)')
    plt.savefig(os.path.join(OUTPUT_DIR, '3_Network_Centrality.png'))
    plt.close()

    # 4. 공간 상호작용 Heatmap (일부 상위 데이터만)
    df_inter = pd.read_excel(xls, '4_공간상호작용_강도', index_col=0)
    plt.figure(figsize=(14, 10))
    # 데이터가 너무 크면 일부만 슬라이싱
    sns.heatmap(df_inter.iloc[:15, :15], annot=True, fmt=".1f", cmap='YlGnBu')
    plt.title('지역-학교 공간 상호작용 강도 (1.0 기준 상회 시 밀접 관계)')
    plt.savefig(os.path.join(OUTPUT_DIR, '4_Spatial_Interaction_Heatmap.png'))
    plt.close()

    print(f"✨ 시각화 완료! 결과물이 '{OUTPUT_DIR}' 폴더에 저장되었습니다.")

if __name__ == "__main__":
    visualize_results()
