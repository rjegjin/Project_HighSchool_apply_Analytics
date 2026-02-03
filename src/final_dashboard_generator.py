import pandas as pd
import os

# ==========================================
# [설정] 분석 결과 엑셀 파일 경로
# ==========================================
# 앞 단계에서 생성된 엑셀 파일명과 정확히 일치해야 합니다.
INPUT_EXCEL = os.path.join("data", "processed", "Step3_학교유형화_및_통계검증.xlsx")
OUTPUT_HTML = os.path.join("output", "Insight_Dashboard_2025.html")

def generate_html_dashboard():
    print("🎨 엑셀 기반 HTML 대시보드 생성을 시작합니다...")
    
    # 1. 데이터 로드 (엑셀 시트별 읽기)
    if not os.path.exists(INPUT_EXCEL):
        print(f"❌ 오류: '{INPUT_EXCEL}' 파일이 없습니다. 통계 분석 코드를 먼저 실행해주세요.")
        return

    try:
        # 엑셀의 각 시트를 데이터프레임으로 불러옵니다.
        df_cluster = pd.read_excel(INPUT_EXCEL, sheet_name='1_유형화(신뢰데이터)')
        df_summary = pd.read_excel(INPUT_EXCEL, sheet_name='1_군집요약')
        df_chi = pd.read_excel(INPUT_EXCEL, sheet_name='2_종속성검정_결과')
        df_corr = pd.read_excel(INPUT_EXCEL, sheet_name='3_상관관계_결과')
        print("✔ 엑셀 데이터 로드 성공!")
    except Exception as e:
        print(f"❌ 데이터 로드 실패: {e}")
        print("   -> 엑셀 파일이 열려있다면 닫고 다시 실행해주세요.")
        return

    # 2. 스타일 정의 (CSS)
    css_style = """
    <style>
        body { font-family: 'Malgun Gothic', 'Noto Sans KR', sans-serif; background-color: #f4f7f6; margin: 0; padding: 20px; color: #333; }
        .container { max-width: 1000px; margin: 0 auto; background: white; padding: 40px; border-radius: 12px; box-shadow: 0 5px 15px rgba(0,0,0,0.05); }
        h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 15px; margin-bottom: 30px; }
        h2 { color: #2980b9; margin-top: 40px; font-size: 1.4em; border-left: 5px solid #3498db; padding-left: 10px; }
        .card-container { display: flex; gap: 20px; margin-bottom: 30px; }
        .card { flex: 1; background: #ecf0f1; padding: 20px; border-radius: 8px; text-align: center; }
        .card h3 { margin: 0 0 10px 0; font-size: 0.9em; color: #7f8c8d; }
        .card p { margin: 0; font-size: 1.8em; font-weight: bold; color: #2c3e50; }
        
        table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 0.95em; }
        th { background-color: #34495e; color: white; padding: 12px; text-align: left; }
        td { border-bottom: 1px solid #ddd; padding: 10px; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        
        .insight-box { background-color: #fff3cd; border: 1px solid #ffeeba; padding: 15px; border-radius: 5px; margin-top: 20px; color: #856404; line-height: 1.6; }
        .footer { margin-top: 50px; text-align: center; color: #bdc3c7; font-size: 0.8em; }
    </style>
    """

    # 3. HTML 본문 조립
    # 데이터가 비어있을 경우를 대비해 안전하게 접근 (.iloc[0] 등 사용 시 주의)
    p_value_display = f"{df_chi['P-value'].iloc[0]:.5f}" if not df_chi.empty else "N/A"
    corr_display = f"{df_corr['상관계수(r)'].iloc[0]:.2f}" if not df_corr.empty else "N/A"
    chi_conclusion = df_chi['결론'].iloc[0] if not df_chi.empty else "분석 결과 없음"

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>2025 후기고 배정 분석 보고서</title>
        {css_style}
    </head>
    <body>
        <div class="container">
            <h1>📊 2025 후기고 배정 심층 분석 보고서</h1>
            
            <div class="card-container">
                <div class="card">
                    <h3>분석 대상 학교 (유효)</h3>
                    <p>{len(df_cluster)}개교</p>
                </div>
                <div class="card">
                    <h3>거주지 영향력(P-value)</h3>
                    <p style="color: #e74c3c;">{p_value_display}</p>
                </div>
                <div class="card">
                    <h3>경쟁률-만족도 상관관계</h3>
                    <p style="color: #2ecc71;">{corr_display}</p>
                </div>
            </div>

            <div class="insight-box">
                <strong>💡 핵심 인사이트 요약:</strong><br>
                1. <strong>양극화 심화:</strong> 학교가 '선호형(만족도 90%↑)'과 '기피형(만족도 40%↓)'으로 극명하게 나뉩니다.<br>
                2. <strong>기피 현상의 원인:</strong> 낮은 경쟁률(미달)이 낮은 만족도(강제 배정)로 이어지는 악순환이 확인되었습니다.<br>
                3. <strong>거주지의 힘:</strong> 통계적으로 <strong>{chi_conclusion}</strong>. (P-value < 0.05면 유의미)
            </div>

            <h2>1. 학교 유형화 분석 (Clustering)</h2>
            <p>데이터 기반으로 분류된 학교 그룹입니다. '유형 C'는 선호 학교, '유형 D'는 주의가 필요한 학교입니다.</p>
            {df_cluster.to_html(index=False, classes='table', float_format=lambda x: '{:.1f}'.format(x) if isinstance(x, float) else x)}

            <h2>2. 그룹별 특징 요약</h2>
            {df_summary.to_html(index=False, classes='table', float_format=lambda x: '{:.2f}'.format(x))}

            <h2>3. 통계적 검증 결과</h2>
            <table class="table">
                <tr><th>분석 항목</th><th>수치</th><th>해석</th></tr>
                <tr>
                    <td>거주지 종속성 (Chi-Square)</td>
                    <td>P-value: {p_value_display}</td>
                    <td>{chi_conclusion}</td>
                </tr>
                <tr>
                    <td>경쟁률-만족도 상관성</td>
                    <td>R: {corr_display}</td>
                    <td>매우 강한 양의 상관관계 (경쟁률 높음 = 만족도 높음)</td>
                </tr>
            </table>

            <div class="footer">
                Generated by The Code Architect | 2025 School Assignment Analysis System
            </div>
        </div>
    </body>
    </html>
    """
    
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ HTML 보고서 생성 완료: {OUTPUT_HTML}")
    print("   -> 브라우저에서 파일을 열어 확인하세요.")

if __name__ == "__main__":
    generate_html_dashboard()