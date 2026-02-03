import pandas as pd
import os

# ==========================================
# [설정] 파일명 (마스킹 등 전처리가 끝난 파일 권장하지만 원본도 가능)
# ==========================================
INPUT_FILE = "2025_후기고_배정결과.xlsx" # 또는 마스킹된 파일
OUTPUT_FILE = "2025_후기고_심층연구보고서.xlsx"

# 분석용 핵심 컬럼 키워드
KEY_DONG = "행정동"       # 주소(동)
KEY_DISTRICT = "자치구"   # 주소(구)
KEY_ASSIGNED = "배정"     # 배정 학교
KEY_CHOICE_1 = "1지망"    # 1지망
KEY_CHOICE_2 = "2지망"    # 2지망
# ==========================================

def find_col(df, keyword):
    """키워드가 포함된 첫 번째 컬럼명을 반환"""
    cols = [c for c in df.columns if keyword in str(c)]
    return cols[0] if cols else None

def find_all_cols(df, keyword):
    """키워드가 포함된 모든 컬럼명을 반환"""
    return [c for c in df.columns if keyword in str(c)]

def run_research():
    print("🔬 고교 배정 영향 요인 심층 연구를 시작합니다...")

    if not os.path.exists(INPUT_FILE):
        print(f"❌ 파일 없음: {INPUT_FILE}")
        return

    try:
        df = pd.read_excel(INPUT_FILE)
    except Exception as e:
        print(f"❌ 로드 실패: {e}")
        return

    # 공백 제거
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].astype(str).str.strip()

    # 컬럼 매핑
    col_dong = find_col(df, KEY_DONG)
    col_assigned = find_col(df, KEY_ASSIGNED)
    cols_1st = find_all_cols(df, KEY_CHOICE_1)
    
    if not (col_dong and col_assigned and cols_1st):
        print("⚠ 필수 컬럼(행정동, 배정학교, 1지망)을 찾을 수 없습니다.")
        return

    print(f"   - 행정동 기준: {col_dong}")
    print(f"   - 배정학교 기준: {col_assigned}")

    # ---------------------------------------------------------
    # [연구 1] 학교별 '인기 지수' vs '기피 지수' 분석
    # ---------------------------------------------------------
    print("📊 1. 학교별 선호도 및 배정 성격 분석 중...")
    
    # 1지망 지원 건수 계산 (단일/일반 등 모든 1지망 합산)
    choice_counts = pd.Series(dtype=int)
    for col in cols_1st:
        counts = df[col].value_counts()
        choice_counts = choice_counts.add(counts, fill_value=0)
    
    # 배정 유형 판별 (1지망/2지망/미지망) 로직 재사용
    def get_assign_type(row):
        school = row[col_assigned]
        # 1지망 군에 있는가?
        for c in cols_1st:
            if row[c] == school: return "1지망"
        # (2지망 생략하고 미지망 여부만 판단)
        return "기타(2지망/미지망)"

    df['배정_성격'] = df.apply(get_assign_type, axis=1)
    
    # 학교별 통계 집계
    school_stats = df.groupby(col_assigned).agg(
        실제배정인원=(col_assigned, 'count'),
        일지망_배정된_사람=('배정_성격', lambda x: (x == '1지망').sum())
    )
    
    # 지표 계산
    school_stats['총_1지망_지원자수'] = choice_counts
    school_stats['실질경쟁률'] = (school_stats['총_1지망_지원자수'] / school_stats['실제배정인원']).round(2)
    school_stats['배정만족도(%)'] = (school_stats['일지망_배정된_사람'] / school_stats['실제배정인원'] * 100).round(1)
    
    # 인사이트: 경쟁률은 높은데 만족도가 낮으면 -> 너무 많이 몰려서 다 튕기고 2지망/임의배정자가 섞임
    # 인사이트: 경쟁률은 낮은데 만족도가 낮으면 -> 1지망 쓴 사람이 거의 없어서 임의배정자가 채워짐 (기피학교)
    
    school_stats = school_stats.sort_values('실질경쟁률', ascending=False)


    # ---------------------------------------------------------
    # [연구 2] 행정동(거주지)별 배정 만족도 (Happiness Index)
    # ---------------------------------------------------------
    print("📊 2. 동네별 배정 만족도(1지망 성공률) 분석 중...")
    
    dong_stats = df.groupby(col_dong).agg(
        거주학생수=(col_dong, 'count'),
        일지망_성공수=('배정_성격', lambda x: (x == '1지망').sum())
    )
    
    dong_stats['1지망_성공률(%)'] = (dong_stats['일지망_성공수'] / dong_stats['거주학생수'] * 100).round(1)
    
    # 학생 수가 너무 적은 동네(5명 미만)는 통계적 의미가 없으므로 필터링 가능
    dong_stats = dong_stats[dong_stats['거주학생수'] >= 5].sort_values('1지망_성공률(%)', ascending=True) # 낮은 순 = 불만 지역


    # ---------------------------------------------------------
    # [연구 3] 행정동 -> 배정학교 흐름 (Geographical Flow)
    # ---------------------------------------------------------
    print("📊 3. 거주지-학교 배정 흐름 매트릭스 생성 중...")
    
    # 행: 행정동, 열: 학교, 값: 인원수
    flow_matrix = pd.crosstab(df[col_dong], df[col_assigned])
    
    # 보기 좋게: 특정 동네에서 가장 많이 간 학교 TOP 3 찾기
    dong_top_schools = []
    for dong in flow_matrix.index:
        top_schools = flow_matrix.loc[dong].nlargest(3)
        row_data = {"행정동": dong}
        for i, (school, count) in enumerate(top_schools.items(), 1):
            row_data[f"Top{i}_학교"] = school
            row_data[f"Top{i}_인원"] = count
        dong_top_schools.append(row_data)
        
    dong_flow_summary = pd.DataFrame(dong_top_schools)


    # ---------------------------------------------------------
    # 결과 저장
    # ---------------------------------------------------------
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        school_stats.to_excel(writer, sheet_name='연구1_학교별_인기도')
        dong_stats.to_excel(writer, sheet_name='연구2_동네별_만족도')
        dong_flow_summary.to_excel(writer, sheet_name='연구3_동네별_주요배정학교')
        flow_matrix.to_excel(writer, sheet_name='부록_동네_학교_전체매트릭스')

    print(f"\n✅ 연구 완료! 파일 생성됨: {OUTPUT_FILE}")
    print("1. [학교별_인기도]: 어떤 학교가 'Wannabe'인지, 어디가 '기피'인지 확인하세요.")
    print("2. [동네별_만족도]: 배정이 유독 안 되는 '불운의 동네'가 어디인지 확인하세요.")
    print("3. [동네별_주요배정]: 우리 동네 애들은 주로 어디로 가는지 확인하세요.")

if __name__ == "__main__":
    run_research()