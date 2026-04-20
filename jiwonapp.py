import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="지원장학 보고서 통합 도구", layout="wide")

# 세션 상태 초기화 (파일 업로드 위젯 초기화용)
if 'file_uploader_key' not in st.state:
    st.session_state['file_uploader_key'] = 0

def reset_files():
    """업로드된 파일 목록을 초기화하는 함수"""
    st.session_state['file_uploader_key'] += 1
    st.rerun()

st.title("📊 지원장학 보고서 유목화 분석용 통합 도구")

# 파일 업로드 및 초기화 버튼 layout
col1, col2 = st.columns([4, 1])
with col1:
    uploaded_files = st.file_uploader(
        "지원장학 보고서 파일들을 선택하세요 (CSV 또는 XLSX)", 
        type=['csv', 'xlsx'], 
        accept_multiple_files=True,
        key=f"uploader_{st.session_state['file_uploader_key']}"
    )
with col2:
    st.write("---")
    if st.button("🔄 파일 초기화", help="업로드된 모든 파일을 목록에서 삭제합니다."):
        reset_files()

def extract_school_data(file):
    """
    파일에서 데이터를 추출하는 개선된 로직
    """
    try:
        # 1. 파일 읽기 (CSV/Excel 구분 및 인코딩 처리)
        if file.name.endswith('.csv'):
            try:
                # 먼저 utf-8-sig로 시도
                df_raw = pd.read_csv(file, encoding='utf-8-sig')
            except:
                # 실패 시 한국어 인코딩(cp949)으로 시도
                file.seek(0)
                df_raw = pd.read_csv(file, encoding='cp949')
        else:
            df_raw = pd.read_excel(file)

        # 2. 헤더 찾기 (샘플 파일처럼 '순'이나 '구 분'이 있는 행을 헤더로 설정)
        header_row_index = None
        for i, row in df_raw.iterrows():
            if '구 분' in row.values and '내용' in row.values:
                header_row_index = i
                break
        
        if header_row_index is None:
            return None # 헤더를 찾지 못한 경우

        # 3. 데이터 프레임 재설정
        df = df_raw.iloc[header_row_index + 1:].copy()
        df.columns = df_raw.iloc[header_row_index].values
        
        # 4. 학교명 추출 (파일명에서 추출하거나 파일 내부 특정 셀에서 추출 가능)
        # 파일명 형식: "중_월촌중학교..."에서 학교명만 추출
        file_name = file.name
        school_name = file_name.split('_')[1].split(' ')[0] if '_' in file_name else file_name
        
        # 5. 유효한 데이터 필터링 ('내용' 항목이 있는 것만)
        # '순'이 숫자가 아니거나 '일시' 같은 행은 제외하고 '학교 현안문제' 등만 남김
        target_categories = ['학교 현안문제', '교육활동 관련 지원 요청 사항']
        valid_df = df[df['구 분'].astype(str).str.contains('현안|지원|요청', na=False)].copy()
        
        if not valid_df.empty:
            valid_df.insert(0, '학교명', school_name)
            # 불필요한 열 제거 및 이름 정리
            return valid_df[['학교명', '구 분', '내용', '관련부서', '관련부서 의견']]
        
        return None

    except Exception as e:
        st.error(f"오류 발생 ({file.name}): {e}")
        return None

if uploaded_files:
    all_rows = []
    
    with st.spinner('데이터를 통합 중입니다...'):
        for file in uploaded_files:
            data = extract_school_data(file)
            if data is not None:
                all_rows.append(data)
    
    if all_rows:
        merged_df = pd.concat(all_rows, ignore_index=True)
        
        st.subheader(f"✅ 통합 완료 (추출된 항목: {len(merged_df)}건)")
        st.dataframe(merged_df, use_container_width=True)
        
        # 엑셀 다운로드 파일 생성
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='통합데이터')
        
        st.download_button(
            label="📥 통합 데이터(Excel) 다운로드",
            data=output.getvalue(),
            file_name="지원장학_분석용_통합파일.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("파일에서 데이터를 추출하지 못했습니다. 파일의 '구 분' 또는 '내용' 열 이름을 확인해주세요.")
