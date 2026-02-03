import pandas as pd
import uuid
import os

# ==========================================
# [설정] 파일명
# ==========================================
INPUT_FILE = "2025_후기고_배정결과.xlsx"
OUTPUT_FILE = "2025_후기고_최종분석결과.xlsx"

# 마스킹 대상 키워드 (헤더에 이 글자가 포함되면 마스킹)
MASK_KEYWORDS = ['성명', '이름', '생년월일', '접수번호', '전화', '연락처', '☎']
# ==========================================

def get_random_token(prefix):
    return f"{prefix}_{uuid.uuid4().hex[:4].upper()}"

def find_columns_by_keyword(df, keyword):
    """특정 키워드가 포함된 모든 컬럼명을 찾습니다."""
    return [col for col in df.columns if keyword in str(col)]

def run_process():
    print("🚀 고교 배정 데이터 심층 분석을 시작합니다...")

    # 1. 파일 로드
    if not os.path.exists(INPUT_FILE):
        print(f"❌ 오류: '{INPUT_FILE}' 파일이 없습니다.")
        return

    try:
        df = pd.read_excel(INPUT_FILE)
        print(f"✔ 파일 로드 성공: 총 {len(df)}명")
    except Exception as e:
        print(f"❌ 엑셀 읽기 실패: {e}")
        return

    # 2. 데이터 전처리 (공백 제거)
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].astype(str).str.strip()

    # 3. 핵심 컬럼 자동 탐색
    # (이미지를 기반으로 '성별', '배정', '1지망', '2지망'이 포함된 컬럼을 찾음)
    cols_gender = find_columns_by_keyword(df, '성별')
    cols_assigned = find_columns_by_keyword(df, '배정고등학교') # 혹은 그냥 '배정'
    
    # 1지망, 2지망 컬럼은 여러 개일 수 있음 (단일학교군, 일반학교군 등)
    cols_choice_1 = find_columns_by_keyword(df, '1지망')
    cols_choice_2 = find_columns_by_keyword(df, '2지망')

    # 컬럼 검증
    if not cols_assigned:
        print("❌ '배정고등학교' 관련 컬럼을 찾을 수 없습니다.")
        print(f"   현재 헤더 목록: {list(df.columns)}")
        return
    if not cols_gender:
        print("⚠ '성별' 컬럼을 찾지 못해 성비 분석이 제한될 수 있습니다.")
    
    main_assigned_col = cols_assigned[0] # 배정고등학교 컬럼 (보통 1개)
    main_gender_col = cols_gender[0] if cols_gender else None

    print(f"   - 배정 컬럼: {main_assigned_col}")
    print(f"   - 1지망 컬럼들: {cols_choice_1}")
    print(f"   - 2지망 컬럼들: {cols_choice_2}")
    
    # 4. 마스킹 (개인정보 보호)
    masked_df = df.copy()
    print("\n🔒 개인정보 마스킹 진행 중...")
    for col in masked_df.columns:
        for keyword in MASK_KEYWORDS:
            if keyword in col:
                masked_df[col] = [get_random_token("MASK") for _ in range(len(masked_df))]
                break # 한 번 마스킹하면 다음 키워드 검사 생략

    # 5. [심층 분석 1] 배정 유형 분류 (1지망/2지망/미지망)
    print("\n📊 배정 적합성 분석 중...")
    
    def classify_assignment(row):
        assigned_school = row[main_assigned_col]
        
        # 1. 1지망 군(단일/일반)에 배정학교가 있는지 확인
        for col in cols_choice_1:
            if row[col] == assigned_school:
                return "1지망 배정"
        
        # 2. 2지망 군에 배정학교가 있는지 확인
        for col in cols_choice_2:
            if row[col] == assigned_school:
                return "2지망 배정"
                
        # 3. 어디에도 없으면
        return "미지망(임의) 배정"

    masked_df['분석_배정유형'] = masked_df.apply(classify_assignment, axis=1)

    # 6. [결과 집계 1] 전체 요약 (요청하신 4가지 지표)
    total_count = len(masked_df)
    type_counts = masked_df['분석_배정유형'].value_counts()
    
    count_1st = type_counts.get("1지망 배정", 0)
    count_2nd = type_counts.get("2지망 배정", 0)
    count_none = type_counts.get("미지망(임의) 배정", 0)
    count_any = count_1st + count_2nd # 지망 내 배정 (1+2)

    summary_data = {
        '구분': [
            '1. 지망 내 배정 (1지망+2지망)', 
            '2. 1지망 배정', 
            '3. 2지망 배정', 
            '4. 미지망(임의) 배정',
            '총 학생 수'
        ],
        '학생 수': [count_any, count_1st, count_2nd, count_none, total_count],
        '비율(%)': [
            round(count_any / total_count * 100, 1),
            round(count_1st / total_count * 100, 1),
            round(count_2nd / total_count * 100, 1),
            round(count_none / total_count * 100, 1),
            100.0
        ]
    }
    summary_df = pd.DataFrame(summary_data)

    # 7. [결과 집계 2] 학교별 배정 인원 및 성비 (남/녀 구분)
    print("📊 학교별 세부 현황 집계 중...")
    
    # 기본 학교별 카운트
    if main_gender_col:
        school_stats = pd.crosstab(
            masked_df[main_assigned_col], 
            masked_df[main_gender_col],
            margins=True,
            margins_name="합계"
        )
    else:
        school_stats = masked_df[main_assigned_col].value_counts().to_frame(name='배정인원')

    # 8. [결과 집계 3] 학교별 배정 유형 상세 (A학교에 온 애들이 1지망 써서 왔나?)
    school_quality = pd.crosstab(
        masked_df[main_assigned_col],
        masked_df['분석_배정유형'],
        margins=True,
        margins_name="합계"
    )

    # 9. 엑셀 저장
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        # 시트1: 요약표 (가장 중요한 4가지 지표)
        summary_df.to_excel(writer, sheet_name='종합_요약', index=False)
        
        # 시트2: 학교별 성비
        school_stats.to_excel(writer, sheet_name='학교별_성비')
        
        # 시트3: 학교별 배정 만족도 (1지망으로 왔는지, 튕겨서 왔는지)
        school_quality.to_excel(writer, sheet_name='학교별_배정유형')
        
        # 시트4: 마스킹된 원본 데이터 (검증용)
        masked_df.to_excel(writer, sheet_name='보안_RawData', index=False)

    print(f"\n✅ 분석 완료! 결과 파일: {OUTPUT_FILE}")
    print("   -> '종합_요약' 시트에서 전체 퍼센트를 확인하세요.")
    print("   -> '학교별_배정유형' 시트에서 학교별 선호도를 확인하세요.")

if __name__ == "__main__":
    run_process()