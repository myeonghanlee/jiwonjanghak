import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="지원장학 분석 도구", layout="wide")

st.title("🏫 학교별 현안 및 아쉬운 점 분석기")
st.info("파이썬 설치 없이 웹에서 바로 엑셀 파일들을 통합 분석합니다.")

uploaded_files = st.file_uploader("엑셀/CSV 파일들을 선택하세요 (최대 40개)", type=["xlsx", "csv"], accept_multiple_files=True)

if uploaded_files:
    all_rows = []
    for file in uploaded_files:
        school_name = file.name.split(" ")[0].replace("중_", "").replace("고_", "")
        try:
            # 8행을 건너뛰고 9행부터 데이터로 인식 (샘플 파일 기준)
            df = pd.read_excel(file, skiprows=8) if file.name.endswith('xlsx') else pd.read_csv(file, skiprows=8)
            
            for _, row in df.iterrows():
                gubun = str(row.get('구 분', ''))
                # '현안문제'나 '지원 요청' 등 정책적 아쉬움이 담긴 행만 추출
                if any(k in gubun for k in ['현안', '요청', '아쉬운']):
                    all_rows.append({
                        '학교명': school_name,
                        '구분': gubun,
                        '내용(아쉬운 점)': str(row.get('내용', '')),
                        '관련부서': str(row.get('관련부서', '')),
                        '부서의견': str(row.get('관련부서 의견', ''))
                    })
        except: continue
    
    master_df = pd.DataFrame(all_rows)
    if not master_df.empty:
        st.subheader("📊 통합 분석 결과")
        # 부서별 필터
        target_dept = st.multiselect("확인할 부서를 선택하세요", options=master_df['관련부서'].unique())
        display_df = master_df[master_df['관련부서'].isin(target_dept)] if target_dept else master_df
        
        st.dataframe(display_df, use_container_width=True)
        
        # 다운로드 버튼
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            display_df.to_excel(writer, index=False)
        st.download_button("결과 엑셀 다운로드", data=output.getvalue(), file_name="분석결과.xlsx")
