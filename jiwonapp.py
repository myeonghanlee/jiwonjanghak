import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="지원장학 분석 도구", layout="wide")

# --- 초기화 로직 ---
# 'clear_count'라는 변수를 세션에 저장하여, 이 값이 바뀔 때마다 업로더를 새로 고침합니다.
if 'clear_count' not in st.session_state:
    st.session_state.clear_count = 0

def reset_app():
    st.session_state.clear_count += 1
    st.rerun()

st.title("🏫 학교별 현안 및 아쉬운 점 분석기")

# 사이드바에 초기화 버튼 배치
with st.sidebar:
    st.header("설정")
    if st.button("🔄 전체 데이터 초기화"):
        reset_app()
    st.info("초기화 버튼을 누르면 업로드된 모든 파일과 분석 결과가 삭제됩니다.")

# 파일 업로더에 key를 부여하여 초기화 시 강제로 새로고침되게 함
uploaded_files = st.file_uploader(
    "엑셀/CSV 파일들을 선택하세요 (최대 40개)", 
    type=["xlsx", "csv"], 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.clear_count}"
)

if uploaded_files:
    all_rows = []
    for file in uploaded_files:
        # 파일명에서 학교명 추출 로직 (파일명: '중_월촌중학교...')
        school_name = file.name.split(" ")[0].replace("중_", "").replace("고_", "")
        try:
            # 8행 건너뛰고 데이터 읽기
            df = pd.read_excel(file, skiprows=8) if file.name.endswith('xlsx') else pd.read_csv(file, skiprows=8)
            
            for _, row in df.iterrows():
                gubun = str(row.get('구 분', ''))
                # '현안문제' 및 '아쉬운 점' 위주 추출
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
        st.subheader(f"📊 분석 결과 (총 {len(uploaded_files)}개교)")
        
        # 필터 및 결과 화면
        search_keyword = st.text_input("특정 키워드 검색 (예: 공사, 인력)")
        display_df = master_df
        if search_keyword:
            display_df = master_df[master_df['내용(아쉬운 점)'].str.contains(search_keyword, na=False)]
            
        st.dataframe(display_df, use_container_width=True, hide_index=True)
        
        # 엑셀 다운로드
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            display_df.to_excel(writer, index=False)
        st.download_button("📥 분석 결과 엑셀로 받기", data=output.getvalue(), file_name="현안분석_결과.xlsx")
