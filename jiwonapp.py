import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="지원장학 보고서 통합 도구", layout="wide")

if 'file_uploader_key' not in st.session_state:
    st.session_state['file_uploader_key'] = 0

def reset_files():
    st.session_state['file_uploader_key'] += 1
    st.rerun()

st.title("📊 지원장학 보고서 유목화 분석용 통합 도구")

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
    if st.button("🔄 파일 일괄 초기화"):
        reset_files()

def extract_school_data(file):
    try:
        # 1. 파일 읽기
        if file.name.endswith('.csv'):
            try:
                df_raw = pd.read_csv(file, encoding='utf-8-sig')
            except:
                file.seek(0)
                df_raw = pd.read_csv(file, encoding='cp949')
        else:
            df_raw = pd.read_excel(file)

        # 2. 헤더('구 분', '내용' 등)가 있는 행 찾기
        header_row_index = None
        for i, row in df_raw.iterrows():
            row_values = [str(val).strip() for val in row.values]
            if '구 분' in row_values and '내용' in row_values:
                header_row_index = i
                break
        
        if header_row_index is None:
            return None

        # 3. 데이터 프레임 재설정 및 컬럼명 지정
        df = df_raw.iloc[header_row_index + 1:].copy()
        df.columns = [str(c).strip() for c in df_raw.iloc[header_row_index].values]
        
        # [핵심 수정] 병합된 셀 처리: '구 분' 열의 빈값(NaN)을 위에서부터 채움
        if '구 분' in df.columns:
            df['구 분'] = df['구 분'].replace('', pd.NA).ffill()

        # 4. 학교명 추출
        file_name = file.name
        school_name = file_name.split('_')[1].split(' ')[0] if '_' in file_name else file_name
        
        # 5. 대상 항목 필터링 (현안문제, 지원 요청 사항)
        # '구 분' 열에서 해당 키워드가 포함된 행만 추출
        mask = df['구 분'].astype(str).str.contains('현안|지원|요청', na=False)
        valid_df = df[mask].copy()
        
        # 6. '내용'이 비어있는 행은 제외 (실제 데이터가 있는 행만)
        valid_df = valid_df[valid_df['내용'].notna() & (valid_df['내용'].astype(str).str.strip() != "")]

        if not valid_df.empty:
            valid_df.insert(0, '학교명', school_name)
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
        
        # 순번 정리 (보기 좋게 1부터 시작)
        merged_df.index = merged_df.index + 1
        
        st.subheader(f"✅ 통합 완료 (추출된 항목: {len(merged_df)}건)")
        st.dataframe(merged_df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            merged_df.to_excel(writer, index=True, index_label='No', sheet_name='통합데이터')
        
        st.download_button(
            label="📥 통합 데이터(Excel) 다운로드",
            data=output.getvalue(),
            file_name="지원장학_분석용_통합파일.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("분석 가능한 데이터를 찾지 못했습니다. '구 분' 열에 '현안' 혹은 '지원' 단어가 있는지 확인해주세요.")
else:
    st.info("학교별 보고서 파일을 업로드해 주세요. '구 분' 셀이 병합되어 있어도 모든 내용을 추출합니다.")
