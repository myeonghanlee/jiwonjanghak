import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="지원장학 보고서 통합 도구", layout="wide")

# 세션 상태 초기화
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

        # 2. 헤더('순', '구 분', '내용' 등)가 있는 행 찾기
        header_row_index = None
        for i, row in df_raw.iterrows():
            row_values = [str(val).strip() for val in row.values]
            if '순' in row_values and '구 분' in row_values and '내용' in row_values:
                header_row_index = i
                break
        
        if header_row_index is None:
            return None

        # 3. 데이터 프레임 설정 및 컬럼명 정리
        df = df_raw.iloc[header_row_index + 1:].copy()
        df.columns = [str(c).strip() for c in df_raw.iloc[header_row_index].values]
        
        # 4. 병합된 셀 처리 (구 분 열 ffill)
        if '구 분' in df.columns:
            df['구 분'] = df['구 분'].replace('', pd.NA).ffill()

        # 5. 학교명 추출
        file_name = file.name
        school_name = file_name.split('_')[1].split(' ')[0] if '_' in file_name else file_name
        
        # 6. [핵심 수정] 유효 데이터 추출 로직 개선
        # '순' 컬럼에 숫자가 있거나, '구 분'이 일시/방문자 정보가 아닌 모든 데이터 추출
        def is_valid_row(row):
            gubun = str(row['구 분'])
            content = str(row['내용'])
            # 분석에서 제외할 키워드 (일시 정보 등)
            exclude_keywords = ['일시', '방문장학', '장학사 작성']
            
            # 내용은 비어있지 않아야 함
            if pd.isna(row['내용']) or content.strip() == "" or content == "nan":
                return False
            
            # 제외 키워드가 포함된 행은 거름
            if any(key in gubun for key in exclude_keywords):
                return False
                
            return True

        valid_df = df[df.apply(is_valid_row, axis=1)].copy()

        if not valid_df.empty:
            valid_df.insert(0, '학교명', school_name)
            # 필요한 공통 열 선택
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
        
        # 전체 데이터 결과 표시
        st.subheader(f"✅ 통합 완료 (추출된 항목: {len(merged_df)}건)")
        st.dataframe(merged_df, use_container_width=True)
        
        # 엑셀 다운로드 생성
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='통합데이터')
        
        st.download_button(
            label="📥 통합 데이터(Excel) 다운로드",
            data=output.getvalue(),
            file_name="지원장학_전체항목_통합파일.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # 통계 정보 제공
        with st.expander("📊 유목화 분석 전 통계 보기"):
            st.write("항목별 빈도수 (구분 기준):")
            st.table(merged_df['구 분'].value_counts())
            
    else:
        st.error("분석 가능한 데이터를 찾지 못했습니다.")
else:
    st.info("파일을 업로드하면 '구 분'에 관계없이 실제 작성된 모든 내용을 수집합니다.")
