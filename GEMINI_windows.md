# Project HighSchool Apply Analytics

## ⚠️ 작업 연속성 (Work Continuity)
**작업을 시작하기 전에 반드시 `dev_log.md`를 확인하십시오.**
이전 작업 내역, 변경 사항, 및 현재 프로젝트 상태를 파악하여 작업의 연속성을 유지해야 합니다.

## 프로젝트 개요
이 프로젝트는 2025학년도 후기고등학교 배정 결과를 분석하고 통계적 심층 연구를 수행하기 위한 데이터 분석 도구입니다.

## 폴더 구조
- **data/**: 데이터 파일 저장소
  - `input/`: 원본 데이터 (예: 2025_후기고_배정결과.xlsx)
  - `processed/`: 처리된 데이터 및 중간 분석 결과
  - `reference/`: 참고용 데이터 (예: 전체학생명렬표)
- **output/**: 최종 분석 리포트 및 시각화 결과물 저장
- **src/**: 분석 소스 코드
  - `final_dashboard_generator.py`: 최종 대시보드 생성
  - `pii_masking.py`: 개인정보 비식별화 처리
  - `research_analytics.py`: 연구 분석 로직
  - `stat_reliability.py`: 통계적 신뢰도 검증
  - `statistical_deep_research.py`: 심층 통계 연구

## 환경 설정
이 프로젝트는 Python 기반으로 작성되었습니다. 필요한 라이브러리를 설치하여 실행할 수 있습니다.

```bash
# 가상환경 생성 및 활성화 (예시)
python -m venv venv
# Windows
.\venv\Scripts\activate
# Mac/Linux
source venv/bin/activate

# 의존성 설치 (requirements.txt가 있다면)
# pip install -r requirements.txt
```

## 사용법
`data/input/` 폴더에 `2025_후기고_배정결과.xlsx` 파일을 위치시킨 후, `src` 폴더 내의 스크립트를 순서대로 실행합니다.

```bash
# 1. 전처리 및 익명화 (Step 1 생성)
python src/pii_masking.py

# 2. 기초 선호도 및 지역 흐름 분석 (Step 2 생성)
python src/research_analytics.py

# 3. 심층 통계 및 학교 유형화 (Step 3 생성)
python src/statistical_deep_research.py

# 4. 최종 대시보드 생성 (HTML 결과물)
python src/final_dashboard_generator.py
```
*최종 결과물은 `output/Insight_Dashboard_2025.html`에 저장됩니다.*
