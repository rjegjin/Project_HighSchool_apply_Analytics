import pandas as pd
import numpy as np
from scipy.stats import chi2_contingency, pearsonr
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
import os

# ==========================================
# [설정] 입력 양식 수정 (단일 엑셀 파일 로드)
# ==========================================
# 여러 개의 csv 파일 대신, 앞 단계에서 만든 엑셀 파일 하나만 있으면 됩니다.
INPUT_EXCEL = "2025_후기고_심층연구보고서.xlsx"
OUTPUT_FILE = "2025_후기고_통계적_심층분석_신뢰도검증.xlsx"

# [중요] 최소 표본 기준 (이 숫자보다 적으면 통계 분석에서 제외)
MIN_SAMPLE_SCHOOL = 10  # 학교별 최소 배정 인원
MIN_SAMPLE_DONG = 10    # 동네별 최소 거주 학생 수

def run_advanced_stats_v2():
    print("🔬 신뢰도 검증이 포함된 심층 통계 연구를 시작합니다...")

    # 1. 데이터 로드 (수정된 부분: read_csv -> read_excel)
    if not os.path.exists(INPUT_EXCEL):
        print(f"❌ 오류: '{INPUT_EXCEL}' 파일이 폴더에 없습니다.")
        return

    try:
        # 엑셀 파일 하나에서 필요한 '시트(Sheet)'를 쏙쏙 뽑아옵니다.
        df_school = pd.read_excel(INPUT_EXCEL, sheet_name='연구1_학교별_인기도')
        df_matrix = pd.read_excel(INPUT_EXCEL, sheet_name='부록_동네_학교_전체매트릭스', index_col=0)
        print("✔ 엑셀 파일 데이터 로드 성공!")
    except Exception as e:
        print(f"❌ 데이터 로드 실패: {e}")
        print("   -> 엑셀 파일이 열려있다면 닫고 다시 실행해주세요.")
        return

    # ---------------------------------------------------------
    # [Pre-Step] 데이터 신뢰도 필터링
    # ---------------------------------------------------------
    print(f"\n🔍 데이터 충분성 검사 (기준: 학교 {MIN_SAMPLE_SCHOOL}명, 동네 {MIN_SAMPLE_DONG}명 이상)")
    
    # 학교 필터링
    valid_schools = df_school[df_school['실제배정인원'] >= MIN_SAMPLE_SCHOOL].copy()
    dropped_schools = len(df_school) - len(valid_schools)
    print(f"   - 학교: 전체 {len(df_school)}개 중 {len(valid_schools)}개 분석 포함 ({dropped_schools}개 제외됨)")

    # 행정동 필터링 (매트릭스에서 행 합계 계산)
    dong_counts = df_matrix.sum(axis=1)
    valid_dongs_idx = dong_counts[dong_counts >= MIN_SAMPLE_DONG].index
    
    # 매트릭스 재구성 (유효한 동네 x 유효한 학교)
    # 학교 컬럼명이 매트릭스 컬럼과 일치한다고 가정
    valid_school_names = valid_schools['배정고등학교'].unique()
    # 매트릭스에 있는 학교만 교집합으로 선택
    common_schools = [s for s in valid_school_names if s in df_matrix.columns]
    
    filtered_matrix = df_matrix.loc[valid_dongs_idx, common_schools]
    print(f"   - 행정동: 전체 {len(df_matrix)}개 중 {len(valid_dongs_idx)}개 분석 포함")

    if len(valid_schools) < 3 or filtered_matrix.empty:
        print("⚠ 경고: 분석할 수 있는 유효 데이터가 너무 적습니다. 기준(MIN_SAMPLE)을 낮춰보세요.")
        return

    # ---------------------------------------------------------
    # [연구 1] K-Means 군집 분석 (유효 학교 대상)
    # ---------------------------------------------------------
    print("\n📊 1. 학교 유형화 (Clustering) - 유효 데이터만")
    
    features = valid_schools[['실질경쟁률', '배정만족도(%)']].fillna(0)
    scaler = StandardScaler()
    scaled_features = scaler.fit_transform(features)
    
    # 데이터가 적으면 클러스터 수도 줄임
    n_clusters = 3 if len(valid_schools) > 10 else 2
    kmeans = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
    valid_schools['군집_Label'] = kmeans.fit_predict(scaled_features)
    
    cluster_summary = valid_schools.groupby('군집_Label')[['실질경쟁률', '배정만족도(%)']].mean().reset_index()
    
    # 군집 이름 부여
    def name_cluster(row):
        comp = row['실질경쟁률']
        sat = row['배정만족도(%)']
        mean_comp = cluster_summary['실질경쟁률'].mean()
        mean_sat = cluster_summary['배정만족도(%)'].mean()
        
        if comp > mean_comp and sat < mean_sat: return "유형A: 고경쟁_아쉬움"
        elif comp < mean_comp and sat > mean_sat: return "유형B: 안정_만족형"
        elif sat > 90: return "유형C: 고만족_적정형"
        else: return "유형D: 복합형"

    cluster_summary['유형_특성'] = cluster_summary.apply(name_cluster, axis=1)
    label_map = dict(zip(cluster_summary['군집_Label'], cluster_summary['유형_특성']))
    valid_schools['분석_학교유형'] = valid_schools['군집_Label'].map(label_map)

    # ---------------------------------------------------------
    # [연구 2] 카이제곱 검정 (유효 동네 x 유효 학교)
    # ---------------------------------------------------------
    print("📊 2. 거주지-배정학교 종속성 검정 (Filtered)")
    
    # 빈도가 0인 컬럼/행 제거 (오류 방지)
    filtered_matrix = filtered_matrix.loc[:, (filtered_matrix != 0).any(axis=0)]
    
    if filtered_matrix.size > 0:
        chi2, p, dof, expected = chi2_contingency(filtered_matrix)
        n = filtered_matrix.sum().sum()
        min_dim = min(filtered_matrix.shape) - 1
        cramer_v = np.sqrt(chi2 / (n * min_dim)) if min_dim > 0 else 0
        
        chi_msg = "통계적 유의함 (믿을만함)" if p < 0.05 else "우연일 가능성 높음"
    else:
        chi2, p, cramer_v, chi_msg = 0, 1, 0, "데이터 부족으로 분석 불가"

    chi_result = pd.DataFrame({
        "분석 대상": [f"동네 {len(valid_dongs_idx)}개 x 학교 {len(common_schools)}개"],
        "P-value": [p],
        "결론": [chi_msg],
        "연관성 강도(V)": [cramer_v]
    })

    # ---------------------------------------------------------
    # [연구 3] 상관관계 분석 (가중치 고려 없이 유효 데이터만)
    # ---------------------------------------------------------
    print("📊 3. 경쟁률-만족도 상관관계 (Filtered)")
    
    if len(valid_schools) > 2:
        corr, p_val = pearsonr(valid_schools['실질경쟁률'], valid_schools['배정만족도(%)'])
        corr_msg = "강한 음의 상관관계" if corr < -0.5 else "약한 상관관계"
    else:
        corr, p_val, corr_msg = 0, 1, "데이터 부족"

    corr_result = pd.DataFrame({
        "분석 학교 수": [len(valid_schools)],
        "상관계수(r)": [corr],
        "P-value": [p_val],
        "해석": [corr_msg]
    })

    # ---------------------------------------------------------
    # 결과 저장
    # ---------------------------------------------------------
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        valid_schools.sort_values('군집_Label').to_excel(writer, sheet_name='1_유형화(신뢰데이터)', index=False)
        cluster_summary.to_excel(writer, sheet_name='1_군집요약', index=False)
        chi_result.to_excel(writer, sheet_name='2_종속성검정_결과')
        corr_result.to_excel(writer, sheet_name='3_상관관계_결과')
        
        # 제외된 학교 목록도 별도 저장 (참고용)
        excluded = df_school[~df_school.index.isin(valid_schools.index)]
        if not excluded.empty:
            excluded.to_excel(writer, sheet_name='부록_제외된_소수데이터', index=False)

    print(f"\n✅ 신뢰도 검증 완료! 파일 생성됨: {OUTPUT_FILE}")
    print(f"   -> 분석에 사용된 학교 수: {len(valid_schools)} (제외 {dropped_schools})")

if __name__ == "__main__":
    run_advanced_stats_v2()