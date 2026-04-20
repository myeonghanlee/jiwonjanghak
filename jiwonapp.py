import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="지원장학 보고서 통합 도구", layout="wide")

# [수정] st.state -> st.session_state로 변경
if 'file_uploader_key' not in st.session_state:
    st.session_state['file_uploader_key'] = 0

def reset_files():
    """업로드된 파일 목록을 초기화하는 함수"""
    st.session_state['file_uploader_key'] += 1
    # 최신 버전에서는 st.rerun()을 사용합니다.
    st.rerun()

st.title("📊 지원장학 보고서 유목화 분석용 통합 도구")

# 파일 업로드 및 초기화 버튼 레이아웃
col1, col2 = st.columns([4, 1])
with col1:
    # key값에 session_state를 반영하여 초기화 가능하게 설정
    uploaded_files = st.file_uploader(
        "지원장학 보고서 파일들을 선택하세요 (CSV 또는 XLSX)", 
        type=['csv', 'xlsx'], 
        accept_multiple_files=True,
        key=f"uploader_{st.session_state['file_uploader_key']}"
    )
with col2:
    st.write("---")
    if st.button("🔄 파일 일괄 초기화"):
        reset_files()

def extract_school_data(file):
    try:
        # 1. 파일 읽기 (CSV/Excel 구분 및 인코딩 처리)
        if file.name.endswith('.csv'):
            try:
                df_raw = pd.read_csv(file, encoding='utf-8-sig')
            except:
                file.seek(0)
                df_raw = pd.read_csv(file, encoding='cp949')
        else:
            df_raw = pd.read_excel(file)

        # 2. 헤더 찾기 (데이터가 시작되는 행 검색)
        header_row_index = None
        for i, row in df_raw.iterrows():
            row_values = [str(val) for val in row.values]
            if '구 분' in row_values and '내용' in row_values:
                header_row_index = i
                break
        
        if header_row_index is None:
            return None

        # 3. 데이터 프레임 재설정
        df = df_raw.iloc[header_row_index + 1:].copy()
        df.columns = df_raw.iloc[header_row_index].values
        
        # 4. 학교명 추출 (파일명 규칙: 중_학교명... 또는 파일명 그대로)
        file_name = file.name
        school_name = file_name.split('_')[1].split(' ')[0] if '_' in file_name else file_name
        
        # 5. 유의미한 항목만 추출 (현안 및 지원 요청 사항)
        # 문자열로 변환 후 '현안' 또는 '지원' 단어가 포함된 행 필터링
        valid_df = df[df['구 분'].astype(str).str.contains('현안|지원|요청', na=False)].copy()
        
        if not valid_df.empty:
            valid_df.insert(0, '학교명', school_name)
            # 필요한 열만 선택 (열 이름이 정확하지 않을 수 있어 에러 방지 처리)
            cols = ['학교명', '구 분', '내용', '관련부서', '관련부서 의견']
            available_cols = [c for c in cols if c in valid_df.columns]
            return valid_df[available_cols]
        
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
        st.error("파일에서 분석 가능한 데이터를 찾지 못했습니다. 파일 구조를 확인해주세요.")
else:
    st.info("학교별 지원장학 보고서(CSV/XLSX)를 업로드해 주세요.")
