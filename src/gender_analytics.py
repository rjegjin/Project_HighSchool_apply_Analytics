import pandas as pd
import os

# ==========================================
# [설정]
# ==========================================
INPUT_FILE = os.path.join("data", "processed", "Step1_전처리_익명화_마스터.xlsx")
OUTPUT_FILE = os.path.join("data", "processed", "Experimental_Gender_Analysis.xlsx")

def run_gender_analysis():
    print("👫 [시나리오 2] 성별 선호도 및 배정 격차 분석을 시작합니다...")

    if not os.path.exists(INPUT_FILE):
        print(f"❌ 파일 없음: {INPUT_FILE}")
        return

    try:
        df = pd.read_excel(INPUT_FILE, sheet_name='보안_RawData')
    except Exception as e:
        print(f"❌ 데이터 로드 실패: {e}")
        return

    # 컬럼 클리닝 (공백 제거 등)
    df.columns = [c.strip() for c in df.columns]
    
    # 핵심 컬럼 식별
    col_gender = '성별'
    col_assigned = '배정고등학교'
    # 1지망 컬럼들
    cols_1st = [c for c in df.columns if '1지망' in c]
    col_assign_type = '분석_배정유형'

    print(f"   - 분석 대상 인원: {len(df)}명")

    # ---------------------------------------------------------
    # 1. 성별 학교별 1지망 선호도
    # ---------------------------------------------------------
    gender_pref = []
    unique_genders = df[col_gender].dropna().unique()
    
    for gender in unique_genders:
        gender_df = df[df[col_gender] == gender]
        pref_counts = pd.Series(dtype=int)
        for col in cols_1st:
            counts = gender_df[col].value_counts()
            pref_counts = pref_counts.add(counts, fill_value=0)
        
        pref_df = pref_counts.to_frame(name=f'{gender}_1지망_지원수')
        gender_pref.append(pref_df)

    if not gender_pref:
        print("❌ 성별 데이터를 찾을 수 없습니다.")
        return
        
    pref_summary = pd.concat(gender_pref, axis=1).fillna(0)
    
    # 남/녀 데이터가 모두 있을 때만 격차 계산
    if '남자_1지망_지원수' in pref_summary.columns and '여자_1지망_지원수' in pref_summary.columns:
        pref_summary['선호도_격차(남-여)'] = pref_summary['남자_1지망_지원수'] - pref_summary['여자_1지망_지원수']
        pref_summary = pref_summary.sort_values('선호도_격차(남-여)', ascending=False)

    # ---------------------------------------------------------
    # 2. 성별 배정 만족도 (1지망 성공률)
    # ---------------------------------------------------------
    satisfaction = df.groupby(col_gender).agg(
        총인원=(col_gender, 'count'),
        일지망_성공=(col_assign_type, lambda x: (x == '1지망 배정').sum())
    )
    satisfaction['1지망_성공률(%)'] = (satisfaction['일지망_성공'] / satisfaction['총인원'] * 100).round(1)

    # ---------------------------------------------------------
    # 3. 학교별 실제 배정 성비
    # ---------------------------------------------------------
    school_gender = pd.crosstab(df[col_assigned], df[col_gender])
    if '남자' in school_gender.columns and '여자' in school_gender.columns:
        school_gender['남초_비율(%)'] = (school_gender['남자'] / (school_gender['남자'] + school_gender['여자']) * 100).round(1)
    
    # ---------------------------------------------------------
    # 결과 저장
    # ---------------------------------------------------------
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        pref_summary.to_excel(writer, sheet_name='1_성별_선호학교_순위')
        satisfaction.to_excel(writer, sheet_name='2_성별_배정만족도')
        school_gender.to_excel(writer, sheet_name='3_학교별_실제성비')

    print(f"\n✅ 분석 완료! 파일 생성됨: {OUTPUT_FILE}")

if __name__ == "__main__":
    run_gender_analysis()
