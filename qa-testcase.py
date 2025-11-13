# =====================================================================================

#2025-11-10 : 비밀번호 인증 기능 추가
#2025-11-11 : JSON 다운로드, [수정] 버튼 추가, 테스트 케이스 - 줄글 형식 저장 기능 추가
#2025-11-12 : JSON 파싱 오류 개선 (간헐적), 속도 향상 개선 함수 추가
#2025-11-13 : 속도 향상 개선 함수 제거, 줄글 형식/기획 문서에 링크 url 항목 추가, [샘플 테스트 케이스 로드] 버튼 제거, AI 테스트 케이스 저장 기능 추가

# =====================================================================================

import streamlit as st
import json
from datetime import datetime
import google.generativeai as genai
import os
import pandas as pd
from io import BytesIO, StringIO

# Excel 지원 확인
try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False
    st.warning("⚠️ Excel 다운로드 기능을 사용하려면 터미널에서 다음 명령을 실행하세요: pip install openpyxl")

# Google Gemini API 클라이언트 초기화
@st.cache_resource
def get_gemini_client():
    api_key = os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        st.error("GOOGLE_API_KEY 환경 변수가 설정되지 않았습니다.")
        return None
    genai.configure(api_key=api_key)
    return genai.GenerativeModel('models/gemini-2.5-flash')
    # return genai.GenerativeModel('models/gemini-2.5-pro') # 품질 중요시

# JSON 파일 경로
TEST_CASES_FILE = "test_cases.json"
SPEC_DOCS_FILE = "spec_docs.json"

# JSON 파일에서 테스트 케이스 불러오기
def load_test_cases_from_file():
    try:
        if os.path.exists(TEST_CASES_FILE):
            with open(TEST_CASES_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        st.error(f"파일 불러오기 실패: {str(e)}")
    return []

# JSON 파일로 테스트 케이스 저장
def save_test_cases_to_file(test_cases):
    try:
        with open(TEST_CASES_FILE, 'w', encoding='utf-8') as f:
            json.dump(test_cases, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        st.error(f"파일 저장 실패: {str(e)}")
        return False

# JSON 파일에서 기획 문서 불러오기
def load_spec_docs_from_file():
    try:
        if os.path.exists(SPEC_DOCS_FILE):
            with open(SPEC_DOCS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        st.error(f"기획 문서 불러오기 실패: {str(e)}")
    return []

# JSON 파일로 기획 문서 저장
def save_spec_docs_to_file(spec_docs):
    try:
        with open(SPEC_DOCS_FILE, 'w', encoding='utf-8') as f:
            json.dump(spec_docs, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        st.error(f"기획 문서 저장 실패: {str(e)}")
        return False

# 세션 스테이트 초기화
if 'test_cases' not in st.session_state:
    st.session_state.test_cases = load_test_cases_from_file()

if 'spec_docs' not in st.session_state:
    st.session_state.spec_docs = load_spec_docs_from_file()

if 'search_history' not in st.session_state:
    st.session_state.search_history = []

# 편집 모드 세션 스테이트
if 'editing_test_case_id' not in st.session_state:
    st.session_state.editing_test_case_id = None

if 'editing_spec_doc_id' not in st.session_state:
    st.session_state.editing_spec_doc_id = None

# 페이지 설정
st.set_page_config(
    page_title="테케봇 (QA Test Case Assistant)",
    page_icon="👾",
    layout="wide"
)

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 테케봇 로그인")
    st.markdown("---")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.info("💡 비밀번호를 입력하세요.")
        
        password = st.text_input(
            "비밀번호",
            type="password",
            placeholder="비밀번호를 입력하세요"
        )
        
        col_a, col_b, col_c = st.columns([1, 1, 1])
        with col_b:
            if st.button("🔓 로그인", type="primary", use_container_width=True):
                correct_password = os.environ.get("APP_PASSWORD", "qabot2025")
                
                if password == correct_password:
                    st.session_state.authenticated = True
                    st.success("✅ 로그인 성공!")
                    st.rerun()
                else:
                    st.error("❌ 잘못된 비밀번호입니다.")    
    st.stop()

st.title("👾 테케봇 (QA Test Case Bot)")
st.markdown("---")

# 사이드바
with st.sidebar:
    st.header("👾 WELCOME")
    
    st.markdown("---")
    
    # 탭으로 구분
    tab1, tab2 = st.tabs(["📝 테스트 케이스", "📚 기획 문서"])
    
    # ============================================
    # 📝 탭 1: 테스트 케이스 추가
    # ============================================
    with tab1:
        with st.expander("➕ [QA팀 전용 버튼]\n테스트 케이스 추가", expanded=False):
            st.markdown("### 📝 테스트 케이스 입력")
            st.info("💡 3가지 방법 중 편한 방식으로 테스트 케이스를 추가하세요!")
            
            # 세션 스테이트에 편집용 데이터프레임 초기화
            if 'edit_df' not in st.session_state:
                st.session_state.edit_df = pd.DataFrame({
                    'NO': [''],
                    'CATEGORY': [''],
                    'DEPTH 1': [''],
                    'DEPTH 2': [''],
                    'DEPTH 3': [''],
                    'PRE-CONDITION': [''],
                    'STEP': [''],
                    'EXPECT RESULT': ['']
                })
            
            # ========== 방법 1: 표 형식 입력 ==========
            st.markdown("**방법 1: 표에서 직접 입력/편집**")
            
            # 행 추가/삭제 버튼
            col1, col2 = st.columns([1, 1])
            with col1:
                if st.button("➕ 행 추가", key="add_row_tc"):
                    new_row = pd.DataFrame({
                        'NO': [''],
                        'CATEGORY': [''],
                        'DEPTH 1': [''],
                        'DEPTH 2': [''],
                        'DEPTH 3': [''],
                        'PRE-CONDITION': [''],
                        'STEP': [''],
                        'EXPECT RESULT': ['']
                    })
                    st.session_state.edit_df = pd.concat([st.session_state.edit_df, new_row], ignore_index=True)
                    st.rerun()
            
            with col2:
                if st.button("🗑️ 모두 지우기", key="clear_tc"):
                    st.session_state.edit_df = pd.DataFrame({
                        'NO': [''],
                        'CATEGORY': [''],
                        'DEPTH 1': [''],
                        'DEPTH 2': [''],
                        'DEPTH 3': [''],
                        'PRE-CONDITION': [''],
                        'STEP': [''],
                        'EXPECT RESULT': ['']
                    })
                    st.rerun()

            # 데이터 에디터를 위한 고유 키 생성
            if 'editor_key' not in st.session_state:
                st.session_state.editor_key = 0
            
            # 데이터 에디터 표시
            edited_df = st.data_editor(
                st.session_state.edit_df,
                use_container_width=True,
                num_rows="dynamic",
                hide_index=True,
                column_config={
                    "NO": st.column_config.TextColumn("NO", width="small", help="번호"),
                    "CATEGORY": st.column_config.TextColumn("CATEGORY", width="medium", help="카테고리 (필수)"),
                    "DEPTH 1": st.column_config.TextColumn("DEPTH 1", width="medium", help="대분류 (필수)"),
                    "DEPTH 2": st.column_config.TextColumn("DEPTH 2", width="medium", help="중분류 (선택)"),
                    "DEPTH 3": st.column_config.TextColumn("DEPTH 3", width="medium", help="소분류 (선택)"),
                    "PRE-CONDITION": st.column_config.TextColumn("PRE-CONDITION", width="large", help="사전 조건 (선택)"),
                    "STEP": st.column_config.TextColumn("STEP", width="large", help="수행 단계"),
                    "EXPECT RESULT": st.column_config.TextColumn("EXPECT RESULT", width="large", help="예상 결과"),
                },
                # key="test_case_editor"
                key=f"test_case_editor_{st.session_state.editor_key}"  # 동적 키 사용
            )
            # 변경사항 즉시 반영
            if not edited_df.equals(st.session_state.edit_df):
                st.session_state.edit_df = edited_df.copy()
                st.session_state.editor_key += 1  # 키 업데이트로 강제 리프레시
                st.rerun()
            
            st.session_state.edit_df = edited_df
            
            # 표 형식 저장 버튼
            if st.button("💾 표 형식 저장", type="primary", disabled=(len(edited_df) == 0), key="save_table_tc"):
                if len(edited_df) > 0:
                    added_count = 0
                    
                    for index, row in edited_df.iterrows():
                        if pd.isna(row['CATEGORY']) or row['CATEGORY'] == '' or pd.isna(row['DEPTH 1']) or row['DEPTH 1'] == '':
                            continue
                        
                        no = str(row['NO']) if row['NO'] and str(row['NO']).strip() else str(len(st.session_state.test_cases) + added_count + 1)
                        
                        structured_test = {
                            "id": len(st.session_state.test_cases) + added_count + 1,
                            "category": str(row['CATEGORY']),
                            "name": f"{row['CATEGORY']} - {row['DEPTH 1']}" + (f" - {row['DEPTH 2']}" if row['DEPTH 2'] else ""),
                            "description": f"NO: {no}\nCATEGORY: {row['CATEGORY']}\nDEPTH1: {row['DEPTH 1']}\nDEPTH2: {row.get('DEPTH 2', '')}\nDEPTH3: {row.get('DEPTH 3', '')}\nPRE-CONDITION: {row.get('PRE-CONDITION', '')}\nSTEP: {row.get('STEP', '')}\nEXPECT RESULT: {row.get('EXPECT RESULT', '')}",
                            "steps": [str(row.get('STEP', ''))] if row.get('STEP', '') else [],
                            "related_features": [x for x in [str(row['CATEGORY']), str(row['DEPTH 1']), str(row.get('DEPTH 2', '')), str(row.get('DEPTH 3', ''))] if x],
                            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "structured_data": {
                                "no": no,
                                "category": str(row['CATEGORY']),
                                "depth1": str(row['DEPTH 1']),
                                "depth2": str(row.get('DEPTH 2', '')),
                                "depth3": str(row.get('DEPTH 3', '')),
                                "pre_condition": str(row.get('PRE-CONDITION', '')),
                                "step": str(row.get('STEP', '')),
                                "expect_result": str(row.get('EXPECT RESULT', ''))
                            }
                        }
                        st.session_state.test_cases.append(structured_test)
                        added_count += 1
                    
                    if added_count > 0:
                        save_test_cases_to_file(st.session_state.test_cases)
                        st.session_state.edit_df = pd.DataFrame({
                            'NO': [''],
                            'CATEGORY': [''],
                            'DEPTH 1': [''],
                            'DEPTH 2': [''],
                            'DEPTH 3': [''],
                            'PRE-CONDITION': [''],
                            'STEP': [''],
                            'EXPECT RESULT': ['']
                        })
                        st.success(f"✅ {added_count}개의 테스트 케이스가 추가되었습니다!")
                        st.rerun()
                    else:
                        st.warning("유효한 테스트 케이스가 없습니다. CATEGORY와 DEPTH 1은 필수 항목입니다.")
            
            st.markdown("---")
            
            # ========== 방법 2: CSV/Excel 파일 업로드 ==========
            st.markdown("**방법 2: CSV/Excel 파일 업로드**")
            uploaded_file = st.file_uploader("CSV 또는 Excel 파일 선택", type=['csv', 'xlsx'], key="upload_tc")
            
            if uploaded_file is not None:
                try:
                    if uploaded_file.name.endswith('.csv'):
                        df = pd.read_csv(uploaded_file)
                    else:
                        df = pd.read_excel(uploaded_file)
                    
                    required_columns = ['NO', 'CATEGORY', 'DEPTH 1', 'DEPTH 2', 'DEPTH 3', 'PRE-CONDITION', 'STEP', 'EXPECT RESULT']
                    
                    if not all(col in df.columns for col in required_columns):
                        st.warning("컬럼명이 일치하지 않습니다. 데이터를 확인해주세요.")
                        st.dataframe(df.head())
                    else:
                        st.session_state.edit_df = df[required_columns].fillna('')
                        st.success(f"✅ {len(df)}개 행이 로드되었습니다!")
                        st.info("👆 위의 표를 확인하고 '💾 표 형식 저장' 버튼을 눌러주세요.")
                        
                except Exception as e:
                    st.error(f"파일 읽기 오류: {str(e)}")
            
            st.markdown("---")
            
            # ========== 방법 3: 줄글 형식 (자유 입력) ==========
            st.markdown("**방법 3: 줄글 형식 (자유 입력)**")
            st.info("💡 테스트 케이스를 자유롭게 작성하고 AI가 학습할 수 있도록 저장하세요!")
            
            tc_free_title = st.text_input(
                "제목 *",
                placeholder="예: 쿠폰 지정 발행 테스트 설계",
                key="tab1_tc_free_title"
            )

            tc_free_link = st.text_input(
                "링크 URL *",
                placeholder="https://www.notion.so/imweb/...",
                key="tab1_tc_free_link"
            )
            
            tc_free_content = st.text_area(
                "내용 *",
                placeholder="테스트 설계 내용을 자유롭게 작성하세요.\n\n[예시]\n1. BO에서 쿠폰 생성\n2. 특정 회원에게 쿠폰 지정 발행\n3. FO에서 쿠폰 사용 가능 여부 확인\n...",
                height=300,
                key="tab1_tc_free_content"
            )
            
            tc_free_category = st.text_input(
                "카테고리 *",
                placeholder="쿠폰",
                key="tab1_tc_free_category"
            )
            
            # 저장 버튼 및 로직
            if st.button("💾 줄글 형식 저장", type="primary", key="tab1_save_free_form_tc"):
                if not tc_free_title or not tc_free_link or not tc_free_content or not tc_free_category:
                    st.warning("⚠️ 모든 항목을 입력해주세요!")
                else:
                    # 줄글 형식으로 저장 (structured_data 없이 저장)
                    free_form_test = {
                        "id": len(st.session_state.test_cases) + 1,
                        "category": tc_free_category if tc_free_category else "기타",
                        "name": tc_free_title,
                        "description": tc_free_content,
                        "steps": [],
                        "related_features": [tc_free_category] if tc_free_category else [],
                        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "input_type": "free_form"  # 🆕 구분자 추가
                    }
                    st.session_state.test_cases.append(free_form_test)
                    save_test_cases_to_file(st.session_state.test_cases)
                    st.success(f"✅ '{tc_free_title}' 테스트 케이스가 저장되었습니다!")
                    st.rerun()

        
        st.markdown("---")
        
        # 테스트 케이스 요약
        st.subheader(f"📋 저장된 테스트 케이스")
        st.metric("전체 케이스 수", f"{len(st.session_state.test_cases)}개")
        
        # JSON 다운로드 버튼
        if st.session_state.test_cases:
            json_data = json.dumps(st.session_state.test_cases, ensure_ascii=False, indent=2)
            st.download_button(
                label="📥 JSON 파일 다운로드",
                data=json_data,
                file_name=f"test_cases_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json",
                use_container_width=True
            )
        
        if st.session_state.test_cases:
            categories = {}
            for tc in st.session_state.test_cases:
                cat = tc.get('category', '미분류')
                categories[cat] = categories.get(cat, 0) + 1
            
            with st.expander("📊 카테고리별 통계", expanded=False):
                for cat, count in sorted(categories.items(), key=lambda x: x[1], reverse=True):
                    st.write(f"**{cat}**: {count}개")
            
            with st.expander("📝 전체 테스트 케이스 보기", expanded=False):
                for tc in st.session_state.test_cases:
                    # 헤더 생성
                    if 'structured_data' in tc:
                        data = tc['structured_data']
                        header = f"[{data['category']}] {data['depth1']}"
                        if data.get('depth2'):
                            header += f" > {data['depth2']}"
                        input_type_badge = "🔹 표"
                    else:
                        header = f"[{tc['category']}] {tc['name']}"
                        input_type_badge = "🔸 줄글" if tc.get('input_type') == 'free_form' else "📥"
                    
                    # 입력 방식 표시
                    with st.expander(f"{input_type_badge} {header}", expanded=False):
                        # 편집 모드
                        if st.session_state.editing_test_case_id == tc['id']:
                            st.markdown("### ✏️ 테스트 케이스 수정")
                            
                            if 'structured_data' in tc:
                                data = tc['structured_data']
                                edit_no = st.text_input("NO", value=data.get('no', ''), key=f"edit_no_{tc['id']}")
                                edit_category = st.text_input("CATEGORY *", value=data.get('category', ''), key=f"edit_cat_{tc['id']}")
                                edit_depth1 = st.text_input("DEPTH 1 *", value=data.get('depth1', ''), key=f"edit_d1_{tc['id']}")
                                edit_depth2 = st.text_input("DEPTH 2", value=data.get('depth2', ''), key=f"edit_d2_{tc['id']}")
                                edit_depth3 = st.text_input("DEPTH 3", value=data.get('depth3', ''), key=f"edit_d3_{tc['id']}")
                                edit_pre_condition = st.text_area("PRE-CONDITION", value=data.get('pre_condition', ''), key=f"edit_pre_{tc['id']}")
                                edit_step = st.text_area("STEP", value=data.get('step', ''), height=150, key=f"edit_step_{tc['id']}")
                                edit_expect = st.text_area("EXPECT RESULT", value=data.get('expect_result', ''), key=f"edit_exp_{tc['id']}")
                            else:
                                edit_category = st.text_input("CATEGORY *", value=tc.get('category', ''), key=f"edit_cat_{tc['id']}")
                                edit_name = st.text_input("제목 *", value=tc.get('name', ''), key=f"edit_name_{tc['id']}")
                                edit_description = st.text_area("내용", value=tc.get('description', ''), height=150, key=f"edit_desc_{tc['id']}")
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                if st.button("💾 저장", key=f"save_edit_{tc['id']}", type="primary"):
                                    if 'structured_data' in tc:
                                        tc['category'] = edit_category
                                        tc['name'] = f"{edit_category} - {edit_depth1}" + (f" - {edit_depth2}" if edit_depth2 else "")
                                        tc['structured_data'] = {
                                            "no": edit_no,
                                            "category": edit_category,
                                            "depth1": edit_depth1,
                                            "depth2": edit_depth2,
                                            "depth3": edit_depth3,
                                            "pre_condition": edit_pre_condition,
                                            "step": edit_step,
                                            "expect_result": edit_expect
                                        }
                                        tc['description'] = f"NO: {edit_no}\nCATEGORY: {edit_category}\nDEPTH1: {edit_depth1}\nDEPTH2: {edit_depth2}\nDEPTH3: {edit_depth3}\nPRE-CONDITION: {edit_pre_condition}\nSTEP: {edit_step}\nEXPECT RESULT: {edit_expect}"
                                    else:
                                        tc['category'] = edit_category
                                        tc['name'] = edit_name
                                        tc['description'] = edit_description
                                    
                                    save_test_cases_to_file(st.session_state.test_cases)
                                    st.session_state.editing_test_case_id = None
                                    st.success("✅ 저장되었습니다!")
                                    st.rerun()
                            
                            with col2:
                                if st.button("❌ 취소", key=f"cancel_edit_{tc['id']}"):
                                    st.session_state.editing_test_case_id = None
                                    st.rerun()
                        
                        # 일반 보기 모드
                        else:
                            if 'structured_data' in tc:
                                data = tc['structured_data']
                                st.write(f"**NO:** {data.get('no', '')}")
                                st.write(f"**CATEGORY:** {data.get('category', '')}")
                                st.write(f"**DEPTH 1:** {data.get('depth1', '')}")
                                if data.get('depth2'):
                                    st.write(f"**DEPTH 2:** {data.get('depth2', '')}")
                                if data.get('depth3'):
                                    st.write(f"**DEPTH 3:** {data.get('depth3', '')}")
                                if data.get('pre_condition'):
                                    st.write(f"**PRE-CONDITION:** {data.get('pre_condition', '')}")
                                st.write(f"**STEP:** {data.get('step', '')}")
                                st.write(f"**EXPECT RESULT:** {data.get('expect_result', '')}")
                            else:
                                st.write(f"**제목:** {tc['name']}")
                                st.write(f"**내용:**")
                                st.text(tc['description'])
                                if tc.get('steps'):
                                    st.write(f"**주요 단계:** {', '.join(tc['steps'])}")
                                if tc.get('related_features'):
                                    st.write(f"**관련 기능:** {', '.join(tc['related_features'])}")
                            
                            # 수정/삭제 버튼
                            col1, col2 = st.columns(2)
                            with col1:
                                if st.button("✏️ 수정", key=f"edit_tc_{tc['id']}"):
                                    st.session_state.editing_test_case_id = tc['id']
                                    st.rerun()
                            with col2:
                                if st.button("🗑️ 삭제", key=f"delete_tc_{tc['id']}"):
                                    st.session_state.test_cases = [t for t in st.session_state.test_cases if t['id'] != tc['id']]
                                    save_test_cases_to_file(st.session_state.test_cases)
                                    st.success("✅ 삭제되었습니다!")
                                    st.rerun()

    
    # 개발자 도구
    with tab1:
        st.markdown("---")
        with st.expander("🔧 개발자 도구", expanded=False):
            if st.button("🔍 사용 가능한 Gemini 모델 확인"):
                try:
                    api_key = os.environ.get("GOOGLE_API_KEY")
                    genai.configure(api_key=api_key)
            
                    models = genai.list_models()
                    st.write("### 사용 가능한 모델 목록:")
                    for model in models:
                        if 'generateContent' in model.supported_generation_methods:
                            st.write(f"✅ {model.name}")
                except Exception as e:
                    st.error(f"오류: {str(e)}")
    
    # ============================================
    # 📚 탭 2: 기획 문서 추가
    # ============================================
    with tab2:
        with st.expander("➕ [QA팀 전용 버튼]\n기획 문서 추가", expanded=False):
            st.markdown("### 📄 기획 문서 입력")
            st.info("💡 노션, Jira에서 작성한 문서를 복사해서 붙여넣으세요.\nAI가 이 내용을 학습합니다!")
            
            # 문서 제목
            doc_title = st.text_input(
                "문서 제목 *",
                placeholder="예: 공동구매 기능 스펙 문서",
                key="tab2_spec_title"
            )
            
            # 문서 유형
            doc_type = st.selectbox(
                "문서 유형 *",
                ["Notion", "Jira", "기타"],
                key="tab2_spec_type"
            )

            # 링크 URL
            doc_link = st.text_input(
                "링크 URL *",
                placeholder="https://www.notion.so/imweb/...",
                key="tab2_spec_link"
            )
            
            # 문서 내용
            doc_content = st.text_area(
                "문서 내용 *",
                placeholder="기획 의도, 스펙, 요구사항 등을 자유롭게 붙여넣으세요.\n\n예:\n[기획 배경]\n현재 공동구매 기능은...\n\n[주요 기능]\n1. 브랜드 정보 입력 모달\n2. 캠페인 생성 기능\n...",
                height=300,
                key="tab2_spec_content"
            )
            
            # 저장 버튼
            if st.button("💾 기획 문서 저장", type="primary", key="tab2_save_spec"):
                if not doc_title or not doc_type or not doc_link or not doc_content:
                    st.warning("⚠️ 모든 항목을 입력해주세요!")
                else:
                    new_spec = {
                        "id": len(st.session_state.spec_docs) + 1,
                        "title": doc_title,
                        "doc_type": doc_type,
                        "link": doc_link,
                        "content": doc_content,
                    }
                    st.session_state.spec_docs.append(new_spec)
                    save_spec_docs_to_file(st.session_state.spec_docs)
                    st.success(f"✅ 기획 문서 '{doc_title}'가 저장되었습니다!")
                    st.rerun()
        
        st.markdown("---")
        
        # 기획 문서 요약
        st.subheader(f"📄 저장된 기획 문서")
        st.metric("전체 문서 수", f"{len(st.session_state.spec_docs)}개")
        
        # JSON 다운로드 버튼
        if st.session_state.spec_docs:
            json_data = json.dumps(st.session_state.spec_docs, ensure_ascii=False, indent=2)
            st.download_button(
                label="📥 JSON 파일 다운로드",
                data=json_data,
                file_name=f"spec_docs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json",
                use_container_width=True
            )
        
        if st.session_state.spec_docs:
            with st.expander("📝 전체 기획 문서 보기", expanded=False):
                for doc in st.session_state.spec_docs:
                    with st.expander(f"[{doc['doc_type']}] {doc['title']}", expanded=False):
                        # 편집 모드
                        if st.session_state.editing_spec_doc_id == doc['id']:
                            st.markdown("### ✏️ 기획 문서 수정")
                            
                            edit_title = st.text_input("문서 제목 *", value=doc['title'], key=f"edit_spec_title_{doc['id']}")
                            edit_type = st.selectbox("문서 유형 *", ["Notion", "Jira", "기타"], index=["Notion", "Jira", "기타"].index(doc['doc_type']), key=f"edit_spec_type_{doc['id']}")
                            edit_link = st.text_input("링크 URL *", value=doc['link'], key=f"edit_spec_link_{doc['id']}")
                            edit_content = st.text_area("문서 내용 *", value=doc['content'], height=300, key=f"edit_spec_content_{doc['id']}")
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                if st.button("💾 저장", key=f"save_spec_edit_{doc['id']}", type="primary"):
                                    doc['title'] = edit_title
                                    doc['doc_type'] = edit_type
                                    doc['link'] = edit_link
                                    doc['content'] = edit_content
                                    
                                    save_spec_docs_to_file(st.session_state.spec_docs)
                                    st.session_state.editing_spec_doc_id = None
                                    st.success("✅ 저장되었습니다!")
                                    st.rerun()
                            
                            with col2:
                                if st.button("❌ 취소", key=f"cancel_spec_edit_{doc['id']}"):
                                    st.session_state.editing_spec_doc_id = None
                                    st.rerun()
                        
                        # 일반 보기 모드
                        else:
                            st.write(f"**내용 미리보기:**")
                            preview = doc['content'][:200] + "..." if len(doc['content']) > 200 else doc['content']
                            st.text(preview)
                            
                            # 수정/삭제 버튼
                            col1, col2 = st.columns(2)
                            with col1:
                                if st.button("✏️ 수정", key=f"edit_spec_{doc['id']}"):
                                    st.session_state.editing_spec_doc_id = doc['id']
                                    st.rerun()
                            with col2:
                                if st.button("🗑️ 삭제", key=f"delete_spec_{doc['id']}"):
                                    st.session_state.spec_docs = [d for d in st.session_state.spec_docs if d['id'] != doc['id']]
                                    save_spec_docs_to_file(st.session_state.spec_docs)
                                    st.success("✅ 삭제되었습니다!")
                                    st.rerun()


# ============================================
# 메인 영역 - AI 기반 테스트 케이스 추천
# ============================================

col1, col2 = st.columns([2, 1])

with col1:
    st.header("🔍 AI 기반 테스트 케이스 추천")
    
    if len(st.session_state.test_cases) == 0 and len(st.session_state.spec_docs) == 0:
        st.warning("⚠️ 먼저 테스트 케이스나 기획 문서를 추가해주세요.")
    else:
        search_query = st.text_area(
            "테스트하고 싶은 기능을 입력하세요.\n설명을 상세하게 적을수록 AI는 더 정확한 케이스를 찾아서 추천해줍니다!",
            placeholder="예: 상품별 구매평 연동 기능 QA\nBO 쇼핑 > 구매평 > 구매평 연동에 해당 기능이 추가될 예정\n테스트 케이스 30개 이상 만들어봐",
            height=150,
            key="search_input"
        )
        
        if st.button("AI 추천 받기", type="primary"):
            if search_query:
                with st.spinner("AI가 연관된 테스트 케이스를 찾고 있습니다..."):
                    client = get_gemini_client()
                    
                    if client:
                        test_cases_str = json.dumps(st.session_state.test_cases, ensure_ascii=False, indent=2)
                        
                        spec_docs_str = ""
                        if st.session_state.spec_docs:
                            spec_docs_str = "\n\n=== 기획 문서 ===\n"
                            for doc in st.session_state.spec_docs:
                                spec_docs_str += f"\n[문서 제목: {doc['title']}]\n[문서 유형: {doc['doc_type']}]\n[내용]\n{doc['content']}\n\n---\n"
                        
                        prompt = f"""[역할 부여]
너는 나와 같이 IT SaaS에 다니고 있는 QA 전문가, QA 엔지니어야.
(1) 테스트 설계, 테스트 케이스 작성, 자동화 업무 수행
(3) 서비스 안정성 기여. 리그레이션을 중심 업무 수행

꼼꼼함이 제일 중요해
확실하지 않은 정보는 '추정' 또는 '불확실'하다고 명시하고, 최신 정보가 필요한 경우 그렇게 알려줘.
혹시나 실제 고객, 회원 이름이 들어간 문서가 있다면, 실제 이름 대신 'Customer A, B, C'를 사용해. 또는 '홍길동', '김영희'와 같은 가명을 사용해줘.
개인정보나 기밀 정보는 일반화하여 처리해.

[제품 정보]
확인하는 제품은 노코드 웹 빌더 시스템이야.
1. IO: 서비스 메인 페이지. 사용자는 회원가입, 로그인을 하고 본인 소유 사이트를 관리하는 페이지
2. BO: Back Office. 사이트 관리자가 접속해서 사이트를 관리하는 공간 (쇼핑몰 세팅, 예약 기능 세팅, 컨텐츠 관리 등). 관리자 페이지에서 '디자인 모드'에 접속할 수 있어.
3. DM: 디자인 모드(Design Mode). 사이트 관리자가 접속해서 사이트를 디자인하는 공간 (상품 상세페이지 디자인 설정, 메뉴 추가/삭제, 메뉴 안에 위젯 추가/삭제 등)
4. FO: Front Office. 실제 사이트 방문자(엔드유저)가 상품을 보고 구매하거나, 예약하거나, 게시글을 보는 곳

[요청]
"{search_query}"에 대한 테스트 케이스 작성

[학습 데이터]
다음은 현재 시스템에 등록된 테스트 케이스들입니다:
{test_cases_str}

{spec_docs_str}

[테스트 케이스 표 양식]
반드시 다음 양식을 따라서 테스트 케이스를 작성해줘:
| NO | CATEGORY | DEPTH 1 | DEPTH 2 | DEPTH 3 | PRE-CONDITION | STEP | EXPECT RESULT |

사용자의 요청을 분석하고, 다음을 수행할 것:
1. 사용자가 테스트하려는 기능과 **직접 관련된** 테스트 케이스를 찾을 것
2. 기획 문서를 참고하여 기능의 의도와 맥락을 파악할 것
3. 그 기능이 작동하기 위해 **의존하는 다른 기능**들을 추론할 것
4. 논리적인 순서로 테스트 체크리스트를 만들 것
5. **반드시 위 표 양식으로 신규 테스트 케이스들을 생성할 것. NO 1부터 번호 시작**

응답 형식:
```json
{{
  "reasoning": "왜 이런 테스트 케이스들이 필요한지 단계별 추론 과정 (한국어로 설명)",
  "existing_test_cases": [
    {{
      "id": 테스트케이스ID,
      "reason": "이 기존 테스트가 왜 필요한지 간단한 설명"
    }}
  ],
  "new_test_cases": [
    {{
      "no": 번호,
      "category": "카테고리",
      "depth1": "대분류",
      "depth2": "중분류 또는 빈 문자열",
      "depth3": "소분류 또는 빈 문자열",
      "pre_condition": "사전조건 또는 빈 문자열",
      "step": "수행 단계",
      "expect_result": "예상 결과"
    }}
  ],
  "test_order": "추천하는 테스트 순서 설명",
  "additional_suggestions": "추가로 필요할 수 있는 테스트 제안(edge case)"
}}
```

중요: 
1. 반드시 JSON 형식으로만 응답
2. new_test_cases는 반드시 표 양식에 맞춰 작성
3. 테스트 케이스와 기획 문서의 맥락을 충분히 반영할 것"""

                        try:
                            response = client.generate_content(prompt)
                            response_text = response.text
                            
                            # 1. 마크다운 제거
                            if "```json" in response_text:
                                json_str = response_text.split("```json")[1].split("```")[0].strip()
                            else:
                                json_str = response_text.strip()

                            # 2. 제어 문자 사전 제거 (파싱 전에 처리)
                            import re
                            json_str_cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', json_str)
                            
                            
                            # 3. JSON 파싱 시도
                            try:
                                ai_response = json.loads(json_str_cleaned)
                            except json.JSONDecodeError as e:
                                st.error(f"❌ JSON 파싱 오류: {str(e)}")
        
                                # 디버깅용: 문제가 되는 부분 표시
                                with st.expander("🔧 디버깅 정보 (개발자용)", expanded=False):
                                    st.write(f"**오류 위치:** line {e.lineno}, column {e.colno}")
                                    st.write(f"**오류 메시지:** {e.msg}")
                                    st.code(json_str_cleaned[:1000], language="json")
        
                                # 4. 최종 fallback: 더 공격적인 정리
                                try:
                                    # 줄바꿈을 공백으로 치환
                                    json_str_final = json_str_cleaned.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
                                    # 연속된 공백 제거
                                    json_str_final = re.sub(r'\s+', ' ', json_str_final)
                                    ai_response = json.loads(json_str_final)
                                    st.warning("⚠️ JSON 파싱에 문제가 있어 일부 데이터가 손실되었을 수 있습니다.")
                                except:
                                    st.error("❌ AI 응답을 처리할 수 없습니다. 다시 시도해주세요.")
                                    st.stop()
                            
                            st.session_state.search_history.append({
                                "query": search_query,
                                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "response": ai_response
                            })



                            # ✅ ai_response를 세션에 저장
                            st.session_state.last_ai_response = ai_response
                            
                            
                            st.success("✅ AI 분석이 완료되었습니다!")



                            
                    # ✅ 세션에 저장된 ai_response 사용
                    if 'last_ai_response' in st.session_state:
                        ai_response = st.session_state.last_ai_response


                   
                            
                            st.markdown("### 🧠 AI의 사고 과정")
                            st.info(ai_response.get("reasoning", "추론 과정 없음"))
                            
                            if ai_response.get("new_test_cases"):
                                st.markdown("### AI가 생성한 신규 테스트 케이스")
                                
                                df_data = []
                                for tc in ai_response.get("new_test_cases", []):
                                    df_data.append({
                                        "NO": tc.get("no", ""),
                                        "CATEGORY": tc.get("category", ""),
                                        "DEPTH 1": tc.get("depth1", ""),
                                        "DEPTH 2": tc.get("depth2", ""),
                                        "DEPTH 3": tc.get("depth3", ""),
                                        "PRE-CONDITION": tc.get("pre_condition", ""),
                                        "STEP": tc.get("step", ""),
                                        "EXPECT RESULT": tc.get("expect_result", "")
                                    })
                                
                                df = pd.DataFrame(df_data)
                                
                                st.dataframe(
                                    df,
                                    use_container_width=True,
                                    hide_index=True
                                )

                                col1, col2 = st.columns(2)

                                with col1:
                                    if EXCEL_AVAILABLE:
                                        output = BytesIO()
                                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                            df.to_excel(writer, index=False, sheet_name='테스트케이스')
                                            workbook = writer.book
                                            worksheet = writer.sheets['테스트케이스']
                                        
                                            header_fill = PatternFill(start_color='4A90A4', end_color='4A90A4', fill_type='solid')
                                            header_font = Font(bold=True, color='FFFFFF')
                                            center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                                        
                                            for cell in worksheet[1]:
                                                cell.fill = header_fill
                                                cell.font = header_font
                                                cell.alignment = center_alignment
                                        
                                            column_widths = {'A': 5, 'B': 15, 'C': 15, 'D': 20, 'E': 20, 'F': 30, 'G': 40, 'H': 40}
                                            for column, width in column_widths.items():
                                                worksheet.column_dimensions[column].width = width
                                    
                                        output.seek(0)
                                        st.download_button(
                                            label="📥 테스트 케이스 Excel로 다운로드",
                                            data=output,
                                            file_name=f"test_cases_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            use_container_width=True
                                        )


                                # 학습 데이터로 저장 버튼 추가
                                with col2:
                                    if st.button("💾 학습시키기", type="primary", use_container_width=True):
                                        # 🔧 저장 전 개수 (추가!)
                                        before_count = len(st.session_state.test_cases)
                                        added_count = 0
            
                                        for tc in ai_response.get("new_test_cases", []):
                                            # 표 형식 데이터를 structured_data로 저장
                                            structured_test = {
                                                "id": len(st.session_state.test_cases) + added_count + 1,
                                                "category": tc.get("category", ""),
                                                "name": f"{tc.get('category', '')} - {tc.get('depth1', '')}" + (f" - {tc.get('depth2', '')}" if tc.get('depth2') else ""),
                                                "description": f"NO: {tc.get('no', '')}\nCATEGORY: {tc.get('category', '')}\nDEPTH1: {tc.get('depth1', '')}\nDEPTH2: {tc.get('depth2', '')}\nDEPTH3: {tc.get('depth3', '')}\nPRE-CONDITION: {tc.get('pre_condition', '')}\nSTEP: {tc.get('step', '')}\nEXPECT RESULT: {tc.get('expect_result', '')}",
                                                "steps": [tc.get('step', '')] if tc.get('step') else [],
                                                "related_features": [x for x in [tc.get('category', ''), tc.get('depth1', ''), tc.get('depth2', ''), tc.get('depth3', '')] if x],
                                                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                                "structured_data": {
                                                    "no": tc.get("no", ""),
                                                    "category": tc.get("category", ""),
                                                    "depth1": tc.get("depth1", ""),
                                                    "depth2": tc.get("depth2", ""),
                                                    "depth3": tc.get("depth3", ""),
                                                    "pre_condition": tc.get("pre_condition", ""),
                                                    "step": tc.get("step", ""),
                                                    "expect_result": tc.get("expect_result", "")
                                                },
                                                "source": "AI_generated"  # AI가 생성한 케이스임을 표시
                                            }
                                            st.session_state.test_cases.append(structured_test)
                                            added_count += 1


                                        # 🔧 저장 후 개수 확인
                                        after_count = len(st.session_state.test_cases)

                                        # JSON 파일에 저장
                                        save_result = save_test_cases_to_file(st.session_state.test_cases)

                                        # ✅ 즉시 표시 (st.rerun() 없이)
                                        if save_result:
                                             st.success(f"✅ {added_count}개 저장 완료! (전: {before_count}개 → 후: {after_count}개)")
                                            # ✅ 저장 후 ai_response 제거 (중복 저장 방지)
                                            del st.session_state.last_ai_response
                                            st.rerun()  # 사이드바 업데이트를 위해
                                        else:
                                            st.error("❌ 저장 실패!")



                            
                            if ai_response.get("test_order"):
                                st.markdown("### 🔄 권장 테스트 순서")
                                st.write(ai_response["test_order"])
                            
                            if ai_response.get("additional_suggestions"):
                                st.markdown("### 💡 추가 제안 (Edge Cases)")
                                st.warning(ai_response["additional_suggestions"])

                            if ai_response.get("existing_test_cases"):
                                st.markdown("### 📝 기존 테스트 케이스 활용")
                                
                                for i, rec in enumerate(ai_response.get("existing_test_cases", []), 1):
                                    test_case = next((tc for tc in st.session_state.test_cases if tc["id"] == rec["id"]), None)
                                    
                                    if test_case:
                                        with st.expander(f"✓ {i}. [{test_case['category']}] {test_case['name']}", expanded=False):
                                            st.markdown(f"**왜 필요한가?** {rec.get('reason', '')}")
                                            st.markdown(f"**설명:** {test_case['description']}")

                        except Exception as e:
                            st.error(f"❌ AI 분석 중 오류가 발생했습니다: {str(e)}")
            else:
                st.warning("검색어를 입력해주세요.")

with col2:
    st.header("📊 검색 히스토리")
    
    if st.session_state.search_history:
        for i, history in enumerate(reversed(st.session_state.search_history[-5:]), 1):
            with st.expander(f"{history['timestamp'][:10]} - {history['query'][:20]}...", expanded=(i==1)):
                st.write(f"**검색어:** {history['query']}")
                existing_count = len(history['response'].get('existing_test_cases', []))
                new_count = len(history['response'].get('new_test_cases', []))
                st.write(f"**기존 테스트:** {existing_count}개")
                st.write(f"**신규 생성:** {new_count}개")
    else:
        st.info("아직 검색 히스토리가 없습니다.")

# 하단 정보
st.markdown("---")
st.markdown("""
#### 💡 사용 방법
1. **학습 데이터 추가 (사이드바)**
   - 📝 테스트 케이스: 기존 테스트 케이스를 표/CSV/Excel로 추가
   - 📚 기획 문서: 노션, Jira 등에서 기획 문서를 복사해서 추가
2. **검색창**에 테스트하고 싶은 기능을 입력하세요
3. **AI가 자동으로** 기존 데이터를 학습하여 신규 테스트 케이스를 생성합니다
4. 생성된 테스트 케이스는 표 형식으로 확인하고 Excel로 다운로드할 수 있습니다


#### 💾 데이터 백업
- JSON 다운로드 버튼으로 테스트 케이스와 기획 문서를 정기적으로 백업할 수 있습니다.
- 앱 재배포 시 데이터가 초기화될 수 있습니다.
""")
