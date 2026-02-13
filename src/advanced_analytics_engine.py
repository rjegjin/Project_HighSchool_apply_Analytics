import pandas as pd
import numpy as np
import os
from sklearn.decomposition import PCA
from sklearn.mixture import GaussianMixture
from sklearn.preprocessing import StandardScaler
from scipy.stats import entropy
import networkx as nx
from statsmodels.multivariate.factor import Factor

# ==========================================
# [설정] 입력 및 출력 경로
# ==========================================
BASE_DIR = "Project_HighSchool_apply_Analytics"
INPUT_EXCEL = os.path.join(BASE_DIR, "data", "processed", "Step2_지망선호도_및_지역흐름.xlsx")
OUTPUT_EXCEL = os.path.join(BASE_DIR, "data", "processed", "Step4_대학원수준_심층분석.xlsx")

def load_data():
    """Step 2에서 생성된 연구 데이터를 로드합니다."""
    if not os.path.exists(INPUT_EXCEL):
        raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {INPUT_EXCEL}")
    
    with pd.ExcelFile(INPUT_EXCEL) as xls:
        df_school = pd.read_excel(xls, '연구1_학교별_인기도')
        df_matrix = pd.read_excel(xls, '부록_동네_학교_전체매트릭스', index_col=0)
    return df_school, df_matrix

def analysis_pca_factor(df_school):
    """1. 다변량 차원 축소 및 잠재 요인 분석 (PCA/Factor Analysis)"""
    print("🔬 [1/5] 다변량 잠재 요인 분석 수행 중...")
    
    # 분석용 지표 선택
    features = ['실제배정인원', '일지망_배정된_사람', '총_1지망_지원자수', '실질경쟁률', '배정만족도(%)']
    x = df_school[features].fillna(0)
    
    # 데이터 표준화
    scaler = StandardScaler()
    x_scaled = scaler.fit_transform(x)
    
    # PCA 수행 (2개 주성분)
    pca = PCA(n_components=2)
    pca_result = pca.fit_transform(x_scaled)
    
    df_school['PCA_1'] = pca_result[:, 0]
    df_school['PCA_2'] = pca_result[:, 1]
    
    # 요인 부하량(Factor Loadings) 확인
    loadings = pd.DataFrame(pca.components_.T, columns=['PC1', 'PC2'], index=features)
    
    return df_school, loadings

def analysis_gmm_clustering(df_school):
    """2. 확률적 모델 기반 군집화 (Gaussian Mixture Model)"""
    print("🔬 [2/5] GMM 기반 확률적 학교 유형화 수행 중...")
    
    features = ['PCA_1', 'PCA_2']
    x = df_school[features]
    
    # GMM 수행 (최적 군집 수는 BIC로 결정할 수 있으나 여기선 4개로 가정)
    gmm = GaussianMixture(n_components=4, random_state=42)
    df_school['GMM_Cluster'] = gmm.fit_predict(x)
    df_school['GMM_Probability'] = gmm.predict_proba(x).max(axis=1) # 소속 확률
    
    return df_school

def analysis_entropy_diversity(df_matrix):
    """3. 정보 엔트로피를 이용한 지망/배정 다양성 분석"""
    print("🔬 [3/5] 정보 엔트로피(Shannon Entropy) 다양성 지수 산출 중...")
    
    # 행별(동네별) 엔트로피 계산: 특정 동네 학생들이 얼마나 다양한 학교로 흩어지는가?
    # 값이 낮을수록 특정 학교로의 배정 쏠림(Segregation)이 강함
    dong_entropy = df_matrix.apply(lambda x: entropy(x + 1e-9), axis=1) # 0 방지
    
    # 열별(학교별) 엔트로피 계산: 특정 학교가 얼마나 다양한 동네에서 학생을 받아들이는가?
    school_entropy = df_matrix.apply(lambda x: entropy(x + 1e-9), axis=0)
    
    df_dong_entropy = pd.DataFrame({'행정동': dong_entropy.index, '엔트로피_지수': dong_entropy.values})
    df_school_entropy = pd.DataFrame({'배정고등학교': school_entropy.index, '포용성_지수': school_entropy.values})
    
    return df_dong_entropy, df_school_entropy

def analysis_network_centrality(df_matrix):
    """4. 네트워크 분석 (Centrality Analysis)"""
    print("🔬 [4/5] 지역-학교 네트워크 중심성 분석 중...")
    
    # 이분 그래프(Bipartite Graph) 또는 가중치 방향 그래프 생성
    G = nx.Graph()
    
    for dong in df_matrix.index:
        for school in df_matrix.columns:
            weight = df_matrix.loc[dong, school]
            if weight > 0:
                G.add_edge(dong, school, weight=weight)
    
    # 고유벡터 중심성(Eigenvector Centrality): 중요도 전이 모델
    try:
        centrality = nx.eigenvector_centrality(G, weight='weight', max_iter=1000)
    except:
        centrality = nx.degree_centrality(G) # 수렴 실패 시 차수 중심성
        
    df_centrality = pd.DataFrame(list(centrality.items()), columns=['ID', '중심성_지수'])
    
    return df_centrality

def analysis_gravity_proxy(df_matrix):
    """5. 공간 상호작용 프록시 분석 (Interaction Intensity)"""
    print("🔬 [5/5] 공간 상호작용 강도 모델링 중...")
    
    # 실제 거리는 데이터에 없으므로, 배정 인원의 집중도를 통해 '공간적 배타성'을 추정
    # 특정 동네 i와 학교 j 사이의 상호작용 강도 I_ij = T_ij / (sum_i * sum_j)
    total_students = df_matrix.values.sum()
    row_sums = df_matrix.sum(axis=1)
    col_sums = df_matrix.sum(axis=0)
    
    expected = np.outer(row_sums, col_sums) / total_students
    interaction_ratio = df_matrix.values / (expected + 1e-9)
    
    df_interaction = pd.DataFrame(interaction_ratio, index=df_matrix.index, columns=df_matrix.columns)
    
    return df_interaction

def main():
    print("🚀 [Advanced Analytics Engine] 대학원 수준 심층 분석 프로세스를 시작합니다.")
    
    try:
        df_school, df_matrix = load_data()
        
        # 1 & 2. PCA + GMM
        df_school, loadings = analysis_pca_factor(df_school)
        df_school = analysis_gmm_clustering(df_school)
        
        # 3. Entropy
        df_dong_entropy, df_school_entropy = analysis_entropy_diversity(df_matrix)
        
        # 4. Network
        df_centrality = analysis_network_centrality(df_matrix)
        
        # 5. Gravity/Interaction
        df_interaction = analysis_gravity_proxy(df_matrix)
        
        # 결과 저장
        print(f"💾 결과를 저장 중입니다: {OUTPUT_EXCEL}")
        with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
            df_school.to_excel(writer, sheet_name='1_학교_고급유형화', index=False)
            loadings.to_excel(writer, sheet_name='1_PCA_부하량')
            df_dong_entropy.to_excel(writer, sheet_name='2_지역_배정다양성', index=False)
            df_school_entropy.to_excel(writer, sheet_name='2_학교_수용다양성', index=False)
            df_centrality.to_excel(writer, sheet_name='3_네트워크_중심성', index=False)
            df_interaction.to_excel(writer, sheet_name='4_공간상호작용_강도')
            
        print("\n✨ 모든 분석이 완료되었습니다. 고차원 통계 지표가 Step 4 파일에 반영되었습니다.")
        
    except Exception as e:
        print(f"❌ 분석 도중 오류 발생: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
