import streamlit as st
import pandas as pd
import io
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment

# --- 1. 페이지 설정 ---
icon_path = "blue-white.png"
st.set_page_config(page_title="JNPMEDI EDC Validation", page_icon=icon_path, layout="wide")

TEMPLATE_PATH = 'EDC Validation_template.xlsx'

st.markdown("""
    <style>
    /* 1. 전체 앱 배경 (연한 회색으로 변경하여 화이트 카드와 대비) */
    .stApp {
        background-color: #F4F7F6;
        color: #333333;
    }

    /* 2. 텍스트 가독성 강제 해결 (흰색 글씨 방지) */
    h1, h2, h3, h4, h5, h6, p, span, div, label {
        color: #2c3e50 !important; /* 진한 남색 계열의 검정 */
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* 3. 입력창 디자인 (화이트 배경 + 검정 글씨) */
    .stTextInput > div > div > input, 
    .stNumberInput > div > div > input ,
    .stSelectbox > div > div {
        background-color: #ffffff !important;
        color: #333333 !important;
        border: 1px solid #dcdcdc;
        border-radius: 8px;
    }
    /* 드롭다운 메뉴 글씨 색상 */
    ul[data-testid="stSelectboxVirtualDropdown"] li {
        color: #333333 !important;
        background-color: #ffffff !important;
    }

    /* 4. 파일 업로더 & 설정 박스 (카드 스타일) */
    .stFileUploader, div[data-testid="stExpander"], div[data-testid="stVerticalBlock"] > div {
        background-color: #ffffff;
        color: #333333 !important;
        border-radius: 10px;
        padding: 5px;
        /* box-shadow: 0 2px 5px rgba(0,0,0,0.05); */ /* 그림자는 선택사항 */
    }
    
    /* 파일 업로더 내부 텍스트 수정 */
    .stFileUploader label {
        font-weight: bold;
        font-size: 1.1em;
    }

    /* 5. 버튼 디자인 (요청하신 #008fd4 블루) */
    .stButton > button, .stDownloadButton > button {
        width: 100%;
        background-color: #008fd4;
        color: #ffffff !important; /* 글씨는 무조건 흰색 */
        font-weight: bold;
        border: none;
        padding: 0.6rem;
        border-radius: 8px;
        transition: all 0.3s ease;
        box-shadow: 0 2px 4px rgba(0,143,212, 0.3);
    }
    
    .stButton > button:hover, .stDownloadButton > button:hover {
        background-color: #006fa3;
        color: #ffffff !important;
        box-shadow: 0 4px 8px rgba(0,111,163, 0.4);
        transform: translateY(-1px);
    }
    .stButton > button:active {
        transform: translateY(0px);
    }

    /* 6. 상태 메시지 박스 커스텀 */
    .success-box {
        padding: 15px;
        background-color: #e3f2fd; /* 아주 연한 블루 */
        color: #0d47a1 !important;
        border-left: 5px solid #008fd4;
        border-radius: 4px;
        margin-bottom: 15px;
        font-weight: 600;
    }
    .error-box {
        padding: 15px;
        background-color: #ffebee; /* 아주 연한 레드 */
        color: #b71c1c !important;
        border-left: 5px solid #d32f2f;
        border-radius: 4px;
        margin-bottom: 15px;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. 핵심 로직 ---

@st.cache_resource
def load_excel_file(file):
    """파일을 메모리에 로드 (속도 향상) - Resource 캐싱 사용"""
    return pd.ExcelFile(file)

def get_dynamic_preview(excel_file, sheet_name, header_row):
    """사용자가 선택한 행을 헤더로 적용하여 미리보기 생성"""
    try:
        # 헤더를 적용하여 5행만 읽어옴
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header_row, nrows=5, dtype=str)
        return df
    except Exception as e:
        return pd.DataFrame()

def check_columns_status(df):
    """필수 컬럼이 식별되는지 진단"""
    if df.empty:
        return False, "데이터 없음", []

    current_cols = [str(c).upper().strip() for c in df.columns]
    
    # 동의어 사전 (Synonyms)
    rename_map = {
        'VAR NAME': 'ITEM ID', 'VARIABLE NAME': 'ITEM ID', 'VARIABLE': 'ITEM ID', 'OID': 'ITEM ID', 
        'ITEMOID': 'ITEM ID', 'QUESTION OID': 'ITEM ID',
        'FORM': 'PAGE', 'FORM OID': 'PAGE', 'FORM NAME': 'PAGE', 'CRF PAGE': 'PAGE',
        'FOLDER': 'VISIT', 'FOLDER OID': 'VISIT', 'EVENT': 'VISIT', 'VISIT NAME': 'VISIT',
        'DATASET': 'DOMAIN', 'LB DOMAIN': 'DOMAIN', 'DOMAIN NAME': 'DOMAIN',
        'VER.': 'VERSION', 
        'VER': 'VERSION', 
        'CRF_VERSION': 'VERSION', 
        'CRF VERSION': 'VERSION', 
    }
    
    mapped_cols = set()
    for col in current_cols:
        if col in rename_map:
            mapped_cols.add(rename_map[col])
        elif col in ['DOMAIN', 'PAGE', 'VISIT', 'ITEM ID']:
            mapped_cols.add(col)
            
    required = {'DOMAIN', 'PAGE', 'VISIT', 'ITEM ID'}
    missing = required - mapped_cols
    
    if not missing:
        return True, "✅ 필수 컬럼 자동 인식 성공!", []
    else:
        return False, f"⚠️ 필수 컬럼 미식별: {', '.join(missing)}", list(missing)

def process_data_final(excel_file, sheet_name, header_row):
    """최종 데이터 처리"""
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header_row, dtype=str)
        df.columns = [str(c).upper().strip() for c in df.columns]
        
        rename_map = {
            'VAR NAME': 'ITEM ID', 'VARIABLE NAME': 'ITEM ID', 'VARIABLE': 'ITEM ID', 'OID': 'ITEM ID', 'ITEMOID': 'ITEM ID',
            'FORM': 'PAGE', 'FORM OID': 'PAGE', 'FORM NAME': 'PAGE', 'CRF PAGE': 'PAGE',
            'FOLDER': 'VISIT', 'FOLDER OID': 'VISIT', 'EVENT': 'VISIT',
            'DATASET': 'DOMAIN', 'LB DOMAIN': 'DOMAIN',
            # 버전 관련 추가
            'VER.': 'VERSION', 'VER': 'VERSION', 'CRF_VERSION': 'VERSION', 'CRF VERSION': 'VERSION'
        }
        df = df.rename(columns=rename_map)
        
        std_cols = ['DOMAIN', 'DOMAIN LABEL', 'PAGE', 'PAGE LABEL', 'VISIT', 
                    'ITEM ID', 'ITEM LABEL', 'ITEM SEQ', 'VERSION', 'CODE', 
                    'LAYOUT', 'TYPE', 'MAX_LEN', 'MIN_VAL', 'MAX_VAL']
        
        for col in std_cols:
            if col not in df.columns: df[col] = ""
            df[col] = df[col].fillna("").astype(str).apply(lambda x: x.replace('.0', '').strip() if x.endswith('.0') else x.strip())

        # JOIN KEY 생성
        df['JOIN_KEY'] = (df['DOMAIN'] + df['PAGE'] + df['VISIT'] + df['ITEM ID']).str.replace(r'\s+', '', regex=True).str.upper()
        
        df = df[df['JOIN_KEY'].str.len() > 1]
        df = df.drop_duplicates(subset=['JOIN_KEY'])
        
        return df
    except Exception as e:
        return pd.DataFrame()

def save_to_template(template_path, df_doc, df_edc, ver_info):
    """
    템플릿 저장 함수 (버전 텍스트 덮어쓰기 수정 + 순서 유지 기능 추가)
    """
    if not os.path.exists(template_path): return None
    wb = load_workbook(template_path)
    
    # -----------------------------------------------------------
    # [수정 1] Cover Page: 텍스트 자체를 교체 (A열 덮어쓰기)
    # -----------------------------------------------------------
    if 'Cover Page' in wb.sheetnames:
        ws_cover = wb['Cover Page']
        
        # 검색할 라벨 키워드와 사용자가 입력한 값 매핑
        cover_mapping = [
            ("Blank eCRF Version", ver_info['blank']),
            ("Database Specifications Version", ver_info['db']),
            ("Annotated CRF Version", ver_info['annotated'])
        ]
        
        # A열(1열)을 훑으며 라벨이 포함된 셀을 찾으면 -> 내용을 통째로 교체
        for r in range(1, 50):
            cell = ws_cover.cell(row=r, column=1)
            cell_val = str(cell.value or "")
            
            for label, user_input in cover_mapping:
                if label in cell_val:
                    cell.value = f"{label}: {user_input}"

    # -----------------------------------------------------------
    # [수정 2] Entry Screen Validation: 병합된 A열 강제 덮어쓰기
    # -----------------------------------------------------------
    target_sheet = 'Entry Screen Validation'
    if target_sheet not in wb.sheetnames: return None
    ws = wb[target_sheet]
    
    # 병합된 셀(A2:O2)의 주인이 A2이므로, A2에 전체 텍스트를 덮어씁니다.
    ws['A2'].value = f"Blank eCRF Version: {ver_info['blank']}"
    ws['A3'].value = f"Database Specifications Version: {ver_info['db']}"
    ws['A4'].value = f"Annotated CRF Version: {ver_info['annotated']}"

    # -----------------------------------------------------------
    # [기존 로직 유지 + 순서 보존 로직 추가] 데이터 비교 및 조건부 서식
    # -----------------------------------------------------------
    
    # 스타일 정의
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # 연분홍
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # 👉 [추가된 로직 1] 기준 문서(df_doc)의 원본 행 순서 기억하기
    df_doc['ORIGINAL_ORDER'] = range(len(df_doc))

    # 데이터 병합 (how='outer')
    merged = pd.merge(df_doc, df_edc, on='JOIN_KEY', how='outer', suffixes=('_Doc', '_EDC'), indicator=True)
    
    # 👉 [추가된 로직 2] 기억해둔 원본 순서대로 재정렬 (EDC에만 있는 값은 마지막에 배치)
    merged = merged.sort_values(by=['ORIGINAL_ORDER'], na_position='last')
    
    # 👉 [추가된 로직 3] 정렬 후 불필요해진 임시 컬럼 삭제
    merged = merged.drop(columns=['ORIGINAL_ORDER'])

    cols = ['DOMAIN', 'DOMAIN LABEL', 'PAGE', 'PAGE LABEL', 'VISIT', 'ITEM ID', 'ITEM LABEL', 'ITEM SEQ', 'VERSION', 'CODE', 'LAYOUT', 'TYPE', 'MAX_LEN', 'MIN_VAL', 'MAX_VAL']

    # 데이터 시작 행: 7행
    start_row = 7
    
    # 👉 [추가된 로직 4] 정렬로 인해 섞인 인덱스를 초기화(reset_index)한 후 반복문 실행
    for i, row in merged.reset_index(drop=True).iterrows():
        curr_r = start_row + i
        status = row['_merge']
        
        doc_vals = [row[f"{c}_Doc"] for c in cols] if status != 'right_only' else [""] * 15
        edc_vals = [row[f"{c}_EDC"] for c in cols] if status != 'left_only' else [""] * 15
        
        mismatches = []
        if status == 'both':
            for idx, cname in enumerate(cols):
                if doc_vals[idx] != edc_vals[idx]: mismatches.append(cname)

        res_text = "True"
        is_false = False
        
        if status == 'left_only' or status == 'right_only':
            res_text = "False"
            is_false = True
        elif mismatches:
            res_text = "False"
            is_false = True

        # 1. Document Area (Col 1~15)
        for idx, val in enumerate(doc_vals):
            cell = ws.cell(row=curr_r, column=idx+1)
            cell.value = val
            cell.border = thin_border
            cell.alignment = align_center
            
            if is_false:
                if status == 'left_only': cell.fill = red_fill
                elif status == 'both' and cols[idx] in mismatches: cell.fill = red_fill

        # 2. EDC Area (Col 16~30 / P~AD)
        for idx, val in enumerate(edc_vals):
            cell = ws.cell(row=curr_r, column=idx+16)
            cell.value = val
            cell.border = thin_border
            cell.alignment = align_center
            
            if is_false:
                if status == 'right_only': cell.fill = red_fill
                elif status == 'both' and cols[idx] in mismatches: cell.fill = red_fill

        # 3. Result Area (Col 31 / AE)
        cell_res = ws.cell(row=curr_r, column=31)
        cell_res.value = res_text
        cell_res.border = thin_border
        cell_res.alignment = align_center

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
    
    """
    템플릿 저장 함수 (버전 텍스트 덮어쓰기 수정)
    """
    if not os.path.exists(template_path): return None
    wb = load_workbook(template_path)
    
    # -----------------------------------------------------------
    # [수정 1] Cover Page: 텍스트 자체를 교체 (A열 덮어쓰기)
    # -----------------------------------------------------------
    if 'Cover Page' in wb.sheetnames:
        ws_cover = wb['Cover Page']
        
        # 검색할 라벨 키워드와 사용자가 입력한 값 매핑
        cover_mapping = [
            ("Blank eCRF Version", ver_info['blank']),
            ("Database Specifications Version", ver_info['db']),
            ("Annotated CRF Version", ver_info['annotated'])
        ]
        
        # A열(1열)을 훑으며 라벨이 포함된 셀을 찾으면 -> 내용을 통째로 교체
        for r in range(1, 50):
            cell = ws_cover.cell(row=r, column=1)
            cell_val = str(cell.value or "")
            
            for label, user_input in cover_mapping:
                if label in cell_val:
                    # [핵심] 옆 칸이 아니라, 해당 셀의 값을 직접 변경
                    # 예: "Blank eCRF Version: V1.1" -> "Blank eCRF Version: 2.1"
                    cell.value = f"{label}: {user_input}"

    # -----------------------------------------------------------
    # [수정 2] Entry Screen Validation: 병합된 A열 강제 덮어쓰기
    # -----------------------------------------------------------
    target_sheet = 'Entry Screen Validation'
    if target_sheet not in wb.sheetnames: return None
    ws = wb[target_sheet]
    
    # 병합된 셀(A2:O2)의 주인이 A2이므로, A2에 전체 텍스트를 덮어씁니다.
    ws['A2'].value = f"Blank eCRF Version: {ver_info['blank']}"
    ws['A3'].value = f"Database Specifications Version: {ver_info['db']}"
    ws['A4'].value = f"Annotated CRF Version: {ver_info['annotated']}"

    # -----------------------------------------------------------
    # [기존 로직 유지] 데이터 비교 및 조건부 서식
    # -----------------------------------------------------------
    
    # 스타일 정의
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # 연분홍
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    merged = pd.merge(df_doc, df_edc, on='JOIN_KEY', how='outer', suffixes=('_Doc', '_EDC'), indicator=True)
    cols = ['DOMAIN', 'DOMAIN LABEL', 'PAGE', 'PAGE LABEL', 'VISIT', 'ITEM ID', 'ITEM LABEL', 'ITEM SEQ', 'VERSION', 'CODE', 'LAYOUT', 'TYPE', 'MAX_LEN', 'MIN_VAL', 'MAX_VAL']

    # 데이터 시작 행: 7행
    start_row = 7
    
    for i, row in merged.iterrows():
        curr_r = start_row + i
        status = row['_merge']
        
        doc_vals = [row[f"{c}_Doc"] for c in cols] if status != 'right_only' else [""] * 15
        edc_vals = [row[f"{c}_EDC"] for c in cols] if status != 'left_only' else [""] * 15
        
        mismatches = []
        if status == 'both':
            for idx, cname in enumerate(cols):
                if doc_vals[idx] != edc_vals[idx]: mismatches.append(cname)

        res_text = "True"
        is_false = False
        
        if status == 'left_only' or status == 'right_only':
            res_text = "False"
            is_false = True
        elif mismatches:
            res_text = "False"
            is_false = True

        # 1. Document Area (Col 1~15)
        for idx, val in enumerate(doc_vals):
            cell = ws.cell(row=curr_r, column=idx+1)
            cell.value = val
            cell.border = thin_border
            cell.alignment = align_center
            
            if is_false:
                if status == 'left_only': cell.fill = red_fill
                elif status == 'both' and cols[idx] in mismatches: cell.fill = red_fill

        # 2. EDC Area (Col 16~30 / P~AD)
        for idx, val in enumerate(edc_vals):
            cell = ws.cell(row=curr_r, column=idx+16)
            cell.value = val
            cell.border = thin_border
            cell.alignment = align_center
            
            if is_false:
                if status == 'right_only': cell.fill = red_fill
                elif status == 'both' and cols[idx] in mismatches: cell.fill = red_fill

        # 3. Result Area (Col 31 / AE)
        cell_res = ws.cell(row=curr_r, column=31)
        cell_res.value = res_text
        cell_res.border = thin_border
        cell_res.alignment = align_center

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- 3. UI 구성 ---
col1, col2 = st.columns([4, 15], vertical_alignment="center")
logo_path = "JNPMEDI_original.jpg"

with col1:
    # 아이콘 표시 (width로 크기 조절)
    st.image(logo_path, width=200) 

with col2:
    # 타이틀 표시
    st.title("EDC Validation")

st.info("실시간 프리뷰를 통해 컬럼이 올바르게 인식되는지 확인 후 검증을 시작하세요.")

col_u1, col_u2 = st.columns(2)
with col_u1:
    doc_file_up = st.file_uploader("📂 기준 문서 (DB Spec)", type=['xlsx', 'xls'], key="doc")
with col_u2:
    edc_file_up = st.file_uploader("📂 비교 대상 (EDC Export)", type=['xlsx', 'xls'], key="edc")

if doc_file_up and edc_file_up:
    st.markdown("---")
    
    try:
        doc_excel = load_excel_file(doc_file_up)
        edc_excel = load_excel_file(edc_file_up)
    except Exception as e:
        st.error(f"엑셀 파일 로드 중 오류: {e}")
        st.stop()

    c1, c2 = st.columns(2)

    # DB Spec 설정
    with c1:
        st.subheader("📄 DB Spec 설정")
        doc_sheet = st.selectbox("시트 선택", doc_excel.sheet_names, key="s1")
        doc_header = st.number_input("헤더 행 (Row Index)", min_value=0, value=1, step=1, key="h1")
        
        doc_df = get_dynamic_preview(doc_excel, doc_sheet, doc_header)
        st.caption(f"▼ '{doc_sheet}' 시트의 {doc_header}번 행을 헤더로 인식한 결과:")
        st.dataframe(doc_df.head(3), use_container_width=True, hide_index=True)
        
        is_ok, msg, missing = check_columns_status(doc_df)
        if is_ok:
            st.markdown(f'<div class="success-box">{msg}</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="error-box">{msg}</div>', unsafe_allow_html=True)
        doc_ready = is_ok

    # EDC Export 설정
    with c2:
        st.subheader("📄 EDC Export 설정")
        edc_sheet = st.selectbox("시트 선택", edc_excel.sheet_names, key="s2")
        edc_header = st.number_input("헤더 행 (Row Index)", min_value=0, value=0, step=1, key="h2")
        
        edc_df = get_dynamic_preview(edc_excel, edc_sheet, edc_header)
        st.caption(f"▼ '{edc_sheet}' 시트의 {edc_header}번 행을 헤더로 인식한 결과:")
        st.dataframe(edc_df.head(3), use_container_width=True, hide_index=True)
        
        is_ok, msg, missing = check_columns_status(edc_df)
        if is_ok:
            st.markdown(f'<div class="success-box">{msg}</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="error-box">{msg}</div>', unsafe_allow_html=True)
        edc_ready = is_ok

    st.markdown("---")
    
    with st.expander("📌 버전 정보 (Optional)", expanded=False):
        v1, v2, v3 = st.columns(3)
        bv = v1.text_input("Blank Ver.", "1.0")
        dv = v2.text_input("DB Spec Ver.", "1.0")
        av = v3.text_input("Annotated Ver.", "1.0")

    if not os.path.exists(TEMPLATE_PATH):
        st.error(f"🚨 중요: 실행 경로에 '{TEMPLATE_PATH}' 파일이 없습니다.")
        btn_disabled = True
    else:
        btn_disabled = not (doc_ready and edc_ready)

    if st.button("🚀 검증 시작 (Start Validation)", type="primary", disabled=btn_disabled):
        with st.status("데이터 분석 중...", expanded=True) as status:
            df_final_doc = process_data_final(doc_excel, doc_sheet, doc_header)
            df_final_edc = process_data_final(edc_excel, edc_sheet, edc_header)
            
            if df_final_doc.empty or df_final_edc.empty:
                status.update(label="❌ 오류 발생", state="error")
                st.error("데이터 로드 실패.")
            else:
                ver_info = {'blank': bv, 'db': dv, 'annotated': av}
                result_file = save_to_template(TEMPLATE_PATH, df_final_doc, df_final_edc, ver_info)
                
                if result_file:
                    status.update(label="🎉 완료!", state="complete")
                    st.success("검증이 성공적으로 완료되었습니다.")
                    
                    today_str = pd.Timestamp.now().strftime('%Y%m%d')
                    file_name = f"EDC Validation List_{today_str}.xlsx"
                    st.download_button(
                        label="📥 결과 리포트 다운로드",
                        data=result_file,
                        file_name=file_name, # 다운로드 일자 반영
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    status.update(label="❌ 템플릿 저장 실패", state="error")
                    st.error("결과 파일 생성 중 오류가 발생했습니다.")
else:
    st.info("👆 먼저 상단에서 두 개의 파일을 업로드해주세요.")