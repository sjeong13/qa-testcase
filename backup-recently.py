# 2025-11-10 백업
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

# 페이지 설정
st.set_page_config(
    page_title="테케봇 (QA Test Case Assistant)",
    page_icon="👾",
    layout="wide"
)

st.title("👾 테케봇 (QA Test Case Bot)")
st.markdown("---")

# 사이드바
with st.sidebar:
    st.header("👾 WELCOME")
    
    # 탭으로 구분
    tab1, tab2 = st.tabs(["📝 테스트 케이스", "📚 기획 문서"])
    
    # ============================================
    # 📋 탭 1: 테스트 케이스 추가 (기존)
    # ============================================
    with tab1:
        with st.expander("➕ [QA팀 전용 버튼]\n테스트 케이스 추가", expanded=False):
            st.markdown("### 📝 테스트 케이스 입력")
            st.info("💡 표에서 직접 입력하거나, 엑셀/구글시트에서 데이터를 복사해서 붙여넣으세요.")
            
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
            
            # 데이터 에디터 (표 형식 입력)
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
                key="test_case_editor"
            )
            
            st.session_state.edit_df = edited_df
            
            st.markdown("---")
            
            # CSV 파일 업로드
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
                        st.success(f"✅ {len(df)}개 행이 로드되었습니다.")
                        st.rerun()
                        
                except Exception as e:
                    st.error(f"파일 읽기 오류: {str(e)}")
            
            # 데이터 추가 버튼
            if st.button("💾 테스트 케이스 저장", type="primary", disabled=(len(edited_df) == 0), key="save_tc"):
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

    # 사이드바 맨 아래에 임시로 추가
    with st.sidebar:
        if st.button("🔍 사용 가능한 모델 확인"):
            try:
                import google.generativeai as genai
                api_key = os.environ.get("GOOGLE_API_KEY")
                genai.configure(api_key=api_key)
            
                models = genai.list_models()
                st.write("### 사용 가능한 모델 목록:")
                for model in models:
                    if 'generateContent' in model.supported_generation_methods:
                        st.write(f"✅ {model.name}")
            except Exception as e:
                st.error(f"오류: {str(e)}")


    
                        
    # 샘플 데이터 로드
    if st.button("📋 샘플 테스트 케이스 로드"):
        sample_cases = [
            {
                "id": 1,
                "category": "회원가입",
                "name": "로그인/가입 모달로 사용 ON/OFF 테스트",
                "description": "'로그인/가입 모달로 사용' 기능 활성화 여부에 따라 회원, 주문 관련 동작이 정상인지 확인",
                "steps": [
                    "디자인 모드 > 공통 디자인 설정에서 회원가입 모달 사용 ON 설정",
                    "비회원이 회원가입 시도",
                    "모달이 정상적으로 표시되는지 확인",
                    "로그인/가입 모달 사용 OFF 설정하여 동작 확인",
                ],
                "related_features": ["회원가입", "로그인", "주문", "가입", "구매"],
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            },
            {
                "id": 2,
                "category": "회원가입",
                "name": "가입 유형별 테스트",
                "description": "일반 이메일 가입, 소셜 가입이 모두 정상 작동하는지 확인",
                "steps": [
                    "이메일 가입 버튼 클릭 > 이메일, 비밀번호 입력 후 가입 > 가입 완료 확인",
                    "카카오 로그인 또는 카카오 싱크로 회원가입 완료 시도",
                    "구글 로그인으로 회원가입 완료 테스트"
                    "네이버 로그인으로 회원가입 완료 테스트",
                    "라인 로그인으로 회원가입 완료 테스트",
                    "애플 로그인으로 회원가입 완료 테스트",
                    "페이스북 로그인으로 회원가입 완료 테스트"

                ],
                "related_features": ["회원가입", "이메일", "소셜로그인", "카카오", "구글", "페이스북", "네이버", "애플", "라인" "로그인", "주문", "가입", "구매"],
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            },
            {
                "id": 3,
                "category": "회원가입",
                "name": "회원 유형별 가입 테스트",
                "description": "일반 회원, 사업자 회원, 새 사용자 추가 유형으로 가입이 되는지 확인",
                "steps": [
                    "일반 회원으로 가입 진행 > 필수 정보 입력 및 가입 완료",
                    "사업자 회원 선택 후 사업자 정보 입력 > 가입 완료"
                    "'BO 환경설정 > 회원가입·그룹·등급에서 사용자 추가' 기능 사용 > 엔드유저가 FO에서 사용자가 추가한 회원 유형으로 회원가입 완료"
                ],
                "related_features": ["회원가입", "일반회원", "사업자회원", "가입", "주문", "구매", "가입유형"],
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            },
            {
                "id": 4,
                "category": "주문",
                "name": "회원 주문 프로세스 테스트",
                "description": "로그인한 회원의 주문 전체 프로세스 검증",
                "steps": [
                    "상품 상세페이지에서 [구매하기] 버튼 클릭 > 로그인 상태 확인 > 주문서로 이동 > 주문 완료",
                    "상품 상세페이지에서 장바구니 담기 > 장바구니 페이지에서 [주문하기] 버튼 클릭 > 로그인 상태 확인 > 주문서로 이동 > 주문 완료]",
                ],
                "related_features": ["주문", "회원", "장바구니", "결제", "구매", "상품 상세페이지"],
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            },
            {
                "id": 5,
                "category": "주문",
                "name": "비회원 주문 프로세스 테스트",
                "description": "비로그인 상태에서 주문이 가능한지 확인",
                "steps": [
                    "로그아웃 상태 확인",
                    "상품 선택 및 장바구니 담기",
                    "상품 상세페이지에서 [구매하기] 버튼 클릭 > 로그인 페이지 또는 로그인 모달에서 [비회원 주문] 버튼 클릭 > 주문서로 이동 > 주문 완료",
                    "상품 상세페이지에서 장바구니 담기 > 장바구니 페이지에서 [주문하기] 버튼 클릭 > 로그인 페이지 또는 로그인 모달에서 [비회원 주문] 버튼 클릭 > 주문서로 이동 > 주문 완료]",
                ],
                "related_features": ["주문", "비회원", "장바구니", "결제", "구매", "상품 상세페이지"],
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            },
            {
                "id": 6,
                "category": "결제",
                "name": "국내 전자 결제 수단별 테스트",
                "description": "신용카드, 계좌이체, 간편결제 등 다양한 결제 수단 테스트",
                "steps": [
                    "무통장 입금 결제 테스트"
                    "신용카드 결제 선택 및 완료",
                    "가상계좌 결제 테스트"
                    "실시간 계좌이체 결제 테스트",
                    "카카오페이 간편결제(직연동) 결제 테스트",
                    "네이버페이 주문형 결제 테스트",
                    "네이버페이 결제형 결제 테스트"
                    "카카오페이 간편결제(이니시스) 결제 테스트",
                    "토스페이(직연동) 결제 테스트",
                    "PAYCO 결제 테스트",
                    "삼성페이 결제 테스트",
                    "휴대폰 결제 테스트",
                    "정기구독 결제 테스트",
                    "톡체크아웃 결제 테스트",
                    "결제 실패 시나리오 테스트"
                ],
                "related_features": ["결제", "신용카드", "간편결제", "주문", "PG"],
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
        ]
        st.session_state.test_cases = sample_cases
        save_test_cases_to_file(st.session_state.test_cases)  # 파일로 저장
        st.success("✅ 샘플 테스트 케이스가 로드되었습니다!")
        st.rerun()
        
        st.markdown("---")
        
        # 테스트 케이스 요약
        st.subheader(f"📋 저장된 테스트 케이스")
        st.metric("전체 케이스 수", f"{len(st.session_state.test_cases)}개")
        
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
                    if 'structured_data' in tc:
                        data = tc['structured_data']
                        header = f"[{data['category']}] {data['depth1']}"
                        if data.get('depth2'):
                            header += f" > {data['depth2']}"
                    else:
                        header = f"[{tc['category']}] {tc['name']}"
                        
                    with st.expander(header, expanded=False):
                        if 'structured_data' in tc:
                            st.write(f"**NO:** {data.get('no', '')}")
                            st.write(f"**CATEGORY:** {data.get('category', '')}")
                            st.write(f"**STEP:** {data.get('step', '')}")
                        else:
                            st.write(f"**설명:** {tc['description']}")
                            
                        if st.button(f"삭제", key=f"delete_tc_{tc['id']}"):
                            st.session_state.test_cases = [t for t in st.session_state.test_cases if t['id'] != tc['id']]
                            save_test_cases_to_file(st.session_state.test_cases)
                            st.rerun()
    
    # ============================================
    # 🆕 탭 2: 기획 문서 추가 (신규)
    # ============================================
    with tab2:
        with st.expander("➕ [QA팀 전용 버튼]\n기획 문서 추가", expanded=False):
            st.markdown("### 📄 기획 문서 입력")
            st.info("💡 노션, Jira에서 작성한 문서를 복사해서 붙여넣으세요.\nAI가 이 내용을 학습합니다!")
            
            # 문서 제목
            doc_title = st.text_input(
                "문서 제목 *",
                placeholder="예: 공동구매 기능 스펙 문서",
                key="spec_title"
            )
            
            # 문서 유형
            doc_type = st.selectbox(
                "문서 유형",
                ["Notion", "Jira", "Confluence", "Google Docs", "기타"],
                key="spec_type"
            )
            
            # 문서 내용 (긴 텍스트)
            doc_content = st.text_area(
                "문서 내용 *",
                placeholder="기획 의도, 스펙, 요구사항 등을 자유롭게 붙여넣으세요.\n\n예:\n[기획 배경]\n현재 공동구매 기능은...\n\n[주요 기능]\n1. 브랜드 정보 입력 모달\n2. 캠페인 생성 기능\n...",
                height=300,
                key="spec_content"
            )
            
            # 저장 버튼
            if st.button("💾 기획 문서 저장", type="primary", key="save_spec"):
                if not doc_title or not doc_content:
                    st.warning("⚠️ 문서 제목과 내용은 필수입니다!")
                else:
                    new_spec = {
                        "id": len(st.session_state.spec_docs) + 1,
                        "title": doc_title,
                        "doc_type": doc_type,
                        "content": doc_content,
                        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    st.session_state.spec_docs.append(new_spec)
                    save_spec_docs_to_file(st.session_state.spec_docs)
                    st.success(f"✅ 기획 문서 '{doc_title}'가 저장되었습니다!")
                    st.rerun()
        
        st.markdown("---")
        
        # 기획 문서 요약
        st.subheader(f"📄 저장된 기획 문서")
        st.metric("전체 문서 수", f"{len(st.session_state.spec_docs)}개")
        
        if st.session_state.spec_docs:
            with st.expander("📝 전체 기획 문서 보기", expanded=False):
                for doc in st.session_state.spec_docs:
                    with st.expander(f"[{doc['doc_type']}] {doc['title']}", expanded=False):
                        st.write(f"**작성일:** {doc['created_at']}")
                        st.write(f"**내용 미리보기:**")
                        preview = doc['content'][:200] + "..." if len(doc['content']) > 200 else doc['content']
                        st.text(preview)
                        
                        if st.button(f"삭제", key=f"delete_spec_{doc['id']}"):
                            st.session_state.spec_docs = [d for d in st.session_state.spec_docs if d['id'] != doc['id']]
                            save_spec_docs_to_file(st.session_state.spec_docs)
                            st.rerun()

# 메인 영역
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
                        # 테스트 케이스 데이터를 문자열로 변환
                        test_cases_str = json.dumps(st.session_state.test_cases, ensure_ascii=False, indent=2)
                        
                        # 기획 문서 데이터를 문자열로 변환
                        spec_docs_str = ""
                        if st.session_state.spec_docs:
                            spec_docs_str = "\n\n=== 기획 문서 ===\n"
                            for doc in st.session_state.spec_docs:
                                spec_docs_str += f"\n[문서 제목: {doc['title']}]\n[문서 유형: {doc['doc_type']}]\n[내용]\n{doc['content']}\n\n---\n"
                        
                        # AI 프롬프트 생성
                        prompt = f"""[역할 부여]
너는 나와 같이 IT SaaS에 다니는 QA 전문가, QA 엔지니어로 일하고 있어.
(1) 테스트 설계, 테스트 케이스 작성 
(2) 자동화 구현 (우선 모니터링용으로) 
(3) 서비스 안정성 기여. 리그레이션을 중점으로 일해.

꼼꼼함이 제일 중요해
확실하지 않은 정보는 '추정' 또는 '불확실'하다고 명시하고, 최신 정보가 필요한 경우 그렇게 알려줘.

[제품 정보]
확인하는 제품은 노코드 웹 빌더 시스템이야.
1. IO: 서비스 메인 페이지. 사용자는 회원가입, 로그인을 하고 본인 소유 사이트를 관리하는 페이지
2. BO: Back Office. 사이트 관리자가 접속해서 사이트를 관리하는 공간 (쇼핑몰 세팅, 예약 기능 세팅, 컨텐츠 관리 등)
3. FO: Front Office. 실제 사이트 방문자(엔드유저)가 상품을 보고 구매하거나, 예약하거나, 게시글을 보는 곳

[현재 요청]
사용자가 "{search_query}"에 대한 테스트를 하려고 합니다.

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
5. **반드시 위 표 양식으로 신규 테스트 케이스들을 생성할 것**

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
1. 반드시 JSON 형식으로만 응답하세요.
2. new_test_cases는 반드시 표 양식에 맞춰 작성하세요.
3. 기획 문서의 맥락을 충분히 반영하세요."""

                        try:
                            response = client.generate_content(prompt)
                            response_text = response.text
                            
                            # JSON 추출
                            if "```json" in response_text:
                                json_str = response_text.split("```json")[1].split("```")[0].strip()
                            else:
                                json_str = response_text.strip()
                                
                            # json.loads 대신 더 관대한 파싱 사용
                            import json
                            import ast
    
                            try:
                                ai_response = json.loads(json_str)
                            except json.JSONDecodeError as e:
                                st.error(f"JSON 파싱 오류: {str(e)}")
                                st.code(json_str[:500])  # 디버깅용: 앞부분만 표시
        
                                # 제어 문자 제거 후 재시도
                                import re
                                json_str_cleaned = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', json_str)
                                ai_response = json.loads(json_str_cleaned)
                            
                            # 검색 히스토리에 추가
                            st.session_state.search_history.append({
                                "query": search_query,
                                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "response": ai_response
                            })
                            
                            # 결과 표시
                            st.success("✅ AI 분석이 완료되었습니다!")
                            
                            # AI의 추론 과정
                            st.markdown("### 🧠 AI의 사고 과정")
                            st.info(ai_response.get("reasoning", "추론 과정 없음"))
                            
                            # 기존 테스트 케이스 추천
                            if ai_response.get("existing_test_cases"):
                                st.markdown("### 📝 기존 테스트 케이스 활용")
                                
                                for i, rec in enumerate(ai_response.get("existing_test_cases", []), 1):
                                    test_case = next((tc for tc in st.session_state.test_cases if tc["id"] == rec["id"]), None)
                                    
                                    if test_case:
                                        with st.expander(f"✓ {i}. [{test_case['category']}] {test_case['name']}", expanded=False):
                                            st.markdown(f"**왜 필요한가?** {rec.get('reason', '')}")
                                            st.markdown(f"**설명:** {test_case['description']}")
                            
                            # 새로 생성된 테스트 케이스 (표 형식)
                            if ai_response.get("new_test_cases"):
                                st.markdown("### 🆕 AI가 생성한 신규 테스트 케이스")
                                
                                # 데이터프레임 생성
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
                                
                                # 스타일링된 표로 표시
                                st.dataframe(
                                    df,
                                    use_container_width=True,
                                    hide_index=True
                                )
                                
                                # Excel 다운로드
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
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                            
                            # 테스트 순서 설명
                            if ai_response.get("test_order"):
                                st.markdown("### 🔄 권장 테스트 순서")
                                st.write(ai_response["test_order"])
                            
                            # 추가 제안
                            if ai_response.get("additional_suggestions"):
                                st.markdown("### 💡 추가 제안 (Edge Cases)")
                                st.warning(ai_response["additional_suggestions"])
                            
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
                st.write(f"**시간(UTC):** {history['timestamp']}")
                existing_count = len(history['response'].get('existing_test_cases', []))
                new_count = len(history['response'].get('new_test_cases', []))
                st.write(f"**기존 테스트:** {existing_count}개")
                st.write(f"**신규 생성:** {new_count}개")
    else:
        st.info("아직 검색 히스토리가 없습니다.")

# 하단 정보
st.markdown("---")
st.markdown("""
### 💡 사용 방법
1. **학습 데이터 추가 (사이드바)**
   - 📝 테스트 케이스: 기존 테스트 케이스를 표/CSV/Excel로 추가
   - 📚 기획 문서: 노션, Jira 등에서 기획 문서를 복사해서 추가
2. **검색창**에 테스트하고 싶은 기능을 입력하세요
3. **AI가 자동으로** 기존 데이터를 학습하여 신규 테스트 케이스를 생성합니다
4. 생성된 테스트 케이스는 표 형식으로 확인하고 Excel로 다운로드할 수 있습니다
""")
