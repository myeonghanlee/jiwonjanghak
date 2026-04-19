import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="지원장학 보고서 데이터 통합 도구", layout="wide")

st.title("📊 지원장학 보고서 유목화 분석용 통합 도구")
st.markdown("""
이 앱은 여러 개의 '지원장학 보고서' 파일을 하나로 합쳐서 AI 분석(유목화)을 하기 좋은 형태로 만들어줍니다.
1. 왼쪽 사이드바에서 40여 개의 파일을 한꺼번에 업로드하세요.
2. 아래 표에서 통합된 내용을 확인하고 **'통합 데이터 다운로드'**를 클릭하세요.
""")

def extract_school_data(file):
    """
    개별 파일에서 학교명과 주요 내용을 추출하는 함수
    """
    try:
        # 파일 읽기 (샘플 구조상 상단 8행은 메타정보, 9행이 헤더)
        # CSV 인코딩은 한국어 환경에 맞춰 'cp949' 또는 'utf-8-sig' 시도
        try:
            df = pd.read_csv(file, skiprows=8, encoding='utf-8-sig')
        except:
            df = pd.read_csv(file, skiprows=8, encoding='cp949')
            
        # 파일명에서 학교명 추출 (예: '중_월촌중학교...' -> '월촌중학교')
        file_name = file.name
        school_name = file_name.split('_')[1].split(' ')[0] if '_' in file_name else file_name
        
        # '내용' 컬럼이 비어있지 않은 행만 필터링
        # 특히 '학교 현안문제' 및 '교육활동 관련 지원 요청 사항'이 포함된 행 추출
        valid_df = df[df['내용'].notna()].copy()
        
        # 학교명 컬럼 추가
        valid_df.insert(0, '학교명', school_name)
        
        return valid_df
    except Exception as e:
        st.error(f"파일 처리 중 오류 발생 ({file.name}): {e}")
        return None

# 사이드바에서 파일 업로드
uploaded_files = st.sidebar.file_uploader(
    "지원장학 보고서 파일들을 선택하세요 (CSV 권장)", 
    type=['csv', 'xlsx'], 
    accept_multiple_files=True
)

if uploaded_files:
    all_rows = []
    
    for file in uploaded_files:
        data = extract_school_data(file)
        if data is not None:
            all_rows.append(data)
    
    if all_rows:
        merged_df = pd.concat(all_rows, ignore_index=True)
        
        # 데이터 정제: '구 분' 열에서 번호 제거 등 깔끔하게 정리 (필요시)
        merged_df['구 분'] = merged_df['구 분'].ffill() # 병합된 셀 처리 대비
        
        # 결과 표시
        st.subheader(f"✅ 통합 완료 (총 {len(merged_df)}개의 항목)")
        st.dataframe(merged_df, use_container_width=True)
        
        # 엑셀 파일로 변환하여 다운로드 제공
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='통합데이터')
        
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 통합된 엑셀 파일 다운로드",
            data=processed_data,
            file_name="통합_지원장학_보고서_분석용.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.info("💡 팁: 다운로드한 파일을 ChatGPT나 Claude에 업로드하고 '유목화해서 분석해줘'라고 요청하세요.")
    else:
        st.warning("데이터를 추출할 수 있는 파일이 없습니다.")
else:
    st.info("파일을 업로드하면 분석이 시작됩니다.")
