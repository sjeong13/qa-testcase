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
    return genai.GenerativeModel('gemini-2.0-flash-exp')

# JSON 파일 경로
TEST_CASES_FILE = "test_cases.json"

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

# 세션 스테이트 초기화
if 'test_cases' not in st.session_state:
    st.session_state.test_cases = load_test_cases_from_file()  # 파일에서 자동 불러오기

if 'search_history' not in st.session_state:
    st.session_state.search_history = []

# 페이지 설정
st.set_page_config(
    page_title="테케봇 (QA Test Case Assistant)",
    page_icon="👾",
    layout="wide"
)

st.title("AI 기반 QA Testcase Assistant (일명 테케봇!)")
st.markdown("---")

# 사이드바 - 테스트 케이스 관리
with st.sidebar:
    st.header("👾 테스트 케이스 관리")
    
# 테스트 케이스 추가
with st.expander("➕ [QA팀 전용] 테스트 케이스 추가", expanded=False):
    st.markdown("## 📝 테스트 케이스 입력")
    st.info("💡  표에서 직접 입력하거나, 엑셀/구글시트에서 데이터를 복사해서 붙여넣으세요.")
    
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
    col1, col2, col3 = st.columns([1, 1, 4])
    with col1:
        if st.button("➕ 행 추가"):
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
        if st.button("🗑️ 모두 지우기"):
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
        num_rows="dynamic",  # 행 추가/삭제 가능
        hide_index=True,
        column_config={
            "NO": st.column_config.TextColumn(
                "NO",
                width="small",
                help="번호"
            ),
            "CATEGORY": st.column_config.TextColumn(
                "CATEGORY",
                width="medium",
                help="카테고리 (필수)"
            ),
            "DEPTH 1": st.column_config.TextColumn(
                "DEPTH 1",
                width="medium",
                help="대분류 (필수)"
            ),
            "DEPTH 2": st.column_config.TextColumn(
                "DEPTH 2",
                width="medium",
                help="중분류 (선택)"
            ),
            "DEPTH 3": st.column_config.TextColumn(
                "DEPTH 3",
                width="medium",
                help="소분류 (선택)"
            ),
            "PRE-CONDITION": st.column_config.TextColumn(
                "PRE-CONDITION",
                width="large",
                help="사전 조건 (선택)"
            ),
            "STEP": st.column_config.TextColumn(
                "STEP",
                width="large",
                help="수행 단계"
            ),
            "EXPECT RESULT": st.column_config.TextColumn(
                "EXPECT RESULT",
                width="large",
                help="예상 결과"
            ),
        },
        key="test_case_editor"
    )
    
    # 편집된 내용을 세션 스테이트에 저장
    st.session_state.edit_df = edited_df
    
    st.markdown("---")
    
    # CSV 파일 업로드 옵션
    st.markdown("**방법 2: CSV/Excel 파일 업로드**")
    uploaded_file = st.file_uploader("CSV 또는 Excel 파일 선택", type=['csv', 'xlsx'])
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            # 컬럼명 확인 및 매핑
            required_columns = ['NO', 'CATEGORY', 'DEPTH 1', 'DEPTH 2', 'DEPTH 3', 'PRE-CONDITION', 'STEP', 'EXPECT RESULT']
            
            # 컬럼명이 다를 경우 매핑 시도
            if not all(col in df.columns for col in required_columns):
                st.warning("컬럼명이 일치하지 않습니다. 데이터를 확인해주세요.")
                st.dataframe(df.head())
            else:
                st.session_state.edit_df = df[required_columns].fillna('')
                st.success(f"✅ {len(df)}개 행이 로드되었습니다.")
                st.rerun()
                
        except Exception as e:
            st.error(f"파일 읽기 오류: {str(e)}")
    
    st.markdown("---")
    
    # 텍스트 영역 입력 (대안)
    with st.expander("📋 텍스트로 붙여넣기 (대안)", expanded=False):
        st.markdown("구글 시트/엑셀에서 복사한 텍스트를 여기 붙여넣으세요.")
        
        # CSV 형식 입력
        csv_input = st.text_area(
            "CSV 형식으로 입력",
            height=200,
            placeholder="""NO,CATEGORY,DEPTH 1,DEPTH 2,DEPTH 3,PRE-CONDITION,STEP,EXPECT RESULT
1,회원가입,공동구매,브랜드 정보 입력,,브랜드 정보 없음,[공동구매 만들기] 버튼 클릭,브랜드 정보 입력 모달 출력
2,주문,주문하기,회원 주문,,로그인 상태,장바구니에서 주문하기 클릭,주문서 페이지로 이동"""
        )
        
        if st.button("CSV 데이터 로드"):
            if csv_input.strip():
                try:
                    # StringIO를 사용하여 CSV로 파싱
                    from io import StringIO
                    csv_data = StringIO(csv_input)
                    df = pd.read_csv(csv_data)
                    
                    # 세션 스테이트에 저장
                    st.session_state.edit_df = df.fillna('')
                    st.success(f"✅ {len(df)}개 행이 로드되었습니다.")
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"CSV 파싱 오류: {str(e)}")
    
    # 데이터 추가 버튼
    col1, col2 = st.columns(2)
    with col1:
        if st.button("💾 테스트 케이스 저장", type="primary", disabled=(len(edited_df) == 0)):
            if len(edited_df) > 0:
                added_count = 0
                
                for index, row in edited_df.iterrows():
                    # 빈 행 스킵
                    if pd.isna(row['CATEGORY']) or row['CATEGORY'] == '' or pd.isna(row['DEPTH 1']) or row['DEPTH 1'] == '':
                        continue
                    
                    # NO가 비어있으면 자동 생성
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
                    # 테이블 초기화
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
    
    with col2:
        # 샘플 다운로드 버튼
        sample_df = pd.DataFrame({
            'NO': ['1', '2', '3'],
            'CATEGORY': ['1. UI 및 진입 경로', '2. 공동구매 메뉴', '3. 주문'],
            'DEPTH 1': ['(BO) 쇼핑 > 상품목록', '공동구매', '주문하기'],
            'DEPTH 2': ['상품 등록', '브랜드 정보 입력', '회원 주문'],
            'DEPTH 3': ['PC', '', ''],
            'PRE-CONDITION': ['', '브랜드 정보 없음', '로그인 상태'],
            'STEP': ['1. 상품 목록 - 상품 등록 버튼을 클릭합니다.\n2. 상품 등록 페이지 내 가로 해상도가 1024px 초과로 조정합니다.', '[공동구매 만들기] 버튼 클릭', '장바구니에서 주문하기 클릭'],
            'EXPECT RESULT': ['1. 신.상품등록 페이지가 PC UI로 출력되는지 확인합니다.', '브랜드 정보 입력 모달 출력', '주문서 페이지로 이동']
        })
        
        if EXCEL_AVAILABLE:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                sample_df.to_excel(writer, index=False, sheet_name='샘플')
            output.seek(0)
            
            st.download_button(
                label="📥 샘플 Excel 다운로드",
                data=output,
                file_name="test_case_sample.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            csv = sample_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 샘플 CSV 다운로드",
                data=csv,
                file_name="test_case_sample.csv",
                mime="text/csv"
            )
    
    # 샘플 데이터 로드
    if st.button("📝 샘플 테스트 케이스 로드"):
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
    
    # 현재 테스트 케이스 목록
    st.markdown("---")
    col_save, col_load = st.columns(2)

    with col_save:
        if st.button("💾 파일로 저장"):
            if save_test_cases_to_file(st.session_state.test_cases):
                st.success("저장 완료!")

    with col_load:
        if st.button("📂 파일에서 불러오기"):
            loaded_cases = load_test_cases_from_file()
            if loaded_cases:
                st.session_state.test_cases = loaded_cases
                st.success(f"{len(loaded_cases)}개 불러오기 완료!")
                st.rerun()

    # 요약만 표시, 상세는 접기
    st.subheader(f"📋 저장된 테스트 케이스")
    st.metric("전체 케이스 수", f"{len(st.session_state.test_cases)}개")

    if st.session_state.test_cases:
        # 카테고리별 통계
        categories = {}
        for tc in st.session_state.test_cases:
            cat = tc.get('category', '미분류')
            categories[cat] = categories.get(cat, 0) + 1
    
        with st.expander("📊 카테고리별 통계", expanded=False):
            for cat, count in sorted(categories.items(), key=lambda x: x[1], reverse=True):
                st.write(f"**{cat}**: {count}개")
    
        # 상세 목록은 expander로 접기
        with st.expander("📝 전체 테스트 케이스 보기", expanded=False):
            for tc in st.session_state.test_cases:
                # 구조화된 데이터가 있는 경우 더 자세히 표시
                if 'structured_data' in tc:
                    data = tc['structured_data']
                    header = f"[{data['category']}] {data['depth1']}"
                    if data.get('depth2'):
                        header += f" > {data['depth2']}"
                else:
                    header = f"[{tc['category']}] {tc['name']}"
                
                with st.expander(header):
                    if 'structured_data' in tc:
                        col1, col2 = st.columns(2)
                        with col1:
                            st.write(f"**NO:** {data.get('no', '')}")
                            st.write(f"**CATEGORY:** {data.get('category', '')}")
                            st.write(f"**DEPTH 1:** {data.get('depth1', '')}")
                            if data.get('depth2'):
                                st.write(f"**DEPTH 2:** {data.get('depth2', '')}")
                            if data.get('depth3'):
                                st.write(f"**DEPTH 3:** {data.get('depth3', '')}")
                        with col2:
                            if data.get('pre_condition'):
                                st.write(f"**PRE-CONDITION:** {data.get('pre_condition', '')}")
                            st.write(f"**STEP:** {data.get('step', '')}")
                            st.write(f"**EXPECT RESULT:** {data.get('expect_result', '')}")
                    else:
                        st.write(f"**설명:** {tc['description']}")
                        if tc.get('related_features'):
                            st.write(f"**연관 기능:** {', '.join(tc['related_features'])}")
                        
                    if st.button(f"삭제", key=f"delete_{tc['id']}"):
                        st.session_state.test_cases = [t for t in st.session_state.test_cases if t['id'] != tc['id']]
                        save_test_cases_to_file(st.session_state.test_cases)
                        st.rerun()

# 메인 영역
col1, col2 = st.columns([2, 1])

with col1:
    st.header("🔍 AI 기반 테스트 케이스 추천")
    
    if len(st.session_state.test_cases) == 0:
        st.warning("⚠️ 먼저 테스트 케이스를 추가하거나 샘플 데이터를 로드해주세요.")
    else:
        search_query = st.text_area(
            "테스트하고 싶은 기능을 입력하세요. 설명을 상세하게 적을수록 AI는 더 정확한 케이스를 찾아서 추천해줍니다!",
            placeholder="예: 상품별 구매평 연동 기능 QA./nBO 쇼핑 > 구매평 > 구매평 연동에 해당 기능이 추가될 예정./n테스트 케이스 30개 이상 만들어봐",
            height=150,
            key="search_input"
        )
        
        if st.button("🤖 AI 추천 받기", type="primary"):
            if search_query:
                with st.spinner("AI가 연관된 테스트 케이스를 찾고 있습니다..."):
                    client = get_gemini_client()
                    
                    if client:
                        # 테스트 케이스 데이터를 문자열로 변환
                        test_cases_str = json.dumps(st.session_state.test_cases, ensure_ascii=False, indent=2)
                        
                        # AI 프롬프트 생성
                        prompt = f"""[역할 부여]
너는 나와 같이 IT SaaS에 다니는 QA 전문가, QA 엔지니어로 일하고 있어.
(1) 테스트 설계, 테스트 케이스 작성 
(2) 자동화 구현 (우선 모니터링용으로) 
(3) 서비스 안정성 기여. 리그레이션을 중점으로 일해.

꼼꼼함이 제일 중요해
확실하지 않은 정보는 '추정' 또는 '불확실'하다고 명시하고, 최신 정보가 필요한 경우 그렇게 알려줘.
개인정보나 기밀 정보는 일반화해서 처리해.

[제품 정보]
확인하는 제품은 노코드 웹 빌더 시스템이야.
1. IO: 서비스 메인 페이지를 의미해. 서비스에 진입하여 사용자는 회원가입, 로그인을 하고 본인 소유 사이트를 관리하는 페이지야.
2. BO: Back Office. 본인 소유 사이트의 백그라운드를 의미해. 사이트 관리자가 접속해서 사이트를 관리하는 공간이야.
(쇼핑몰 세팅, 예약 기능 세팅, 컨텐츠 관리 등등). 관리자 페이지에서 "디자인 모드(Design Mode)(aka. DM)"에 접속할 수 있어.
디자인 모드에서는 사이트 디자인을 해.
3. FO: Front Office. 실제 사이트에 방문하는 곳을 의미해. 여기에 방문하는 사람은 엔드유저(End user)라고 부르곤 해. 엔드유저가 사이트에 진입해서 상품을 보고 구매하거나, 예약하거나, 게시글을 볼 수 있어.

IO, BO, FO는 서로 연관도 많이 되어 있고, 얽혀있어.
즉, QA 엔지니어인 너는 각 영역을 종합해서 설계할 수 있어야 해.

                        
[현재 요청]
사용자가 "{search_query}"에 대한 테스트를 하려고 합니다.

다음은 현재 시스템에 등록된 테스트 케이스들입니다:

{test_cases_str}

[테스트 케이스 표 양식]
반드시 다음 양식을 따라서 테스트 케이스를 작성해줘:
| NO | CATEGORY | DEPTH 1 | DEPTH 2 | DEPTH 3 | PRE-CONDITION | STEP | EXPECT RESULT |
|----|----------|---------|---------|---------|---------------|------|---------------|
| 1  | 카테고리명 | 대분류 | 중분류(선택) | 소분류(선택) | 사전조건(선택) | 수행 단계 | 예상 결과 |

예시:
- CATEGORY: 공동구매 메뉴, LNB 메뉴 등
- DEPTH 1: 공동구매, 상단 헤더 등  
- DEPTH 2: 공동구매 브랜드 정보 모달, 공동구매 만들기 등
- DEPTH 3: 더 세부적인 기능
- PRE-CONDITION: 브랜드 정보 없음 + 캠페인 없음, 파트너 하위 사이트 등
- STEP: 공동구매 메뉴 확인, [공동구매 만들기] 버튼 클릭 등
- EXPECT RESULT: 공동구매 메뉴 노출 확인, 브랜드 정보 저장 완료 모달 출력 등

사용자의 요청을 분석하고, 다음을 수행할 것:

1. 사용자가 테스트하려는 기능과 **직접 관련된** 테스트 케이스를 찾을 것
2. 그 기능이 작동하기 위해 **의존하는 다른 기능**들을 추론할 것
3. 의존하는 기능들의 테스트 케이스도 포함할 것
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
3. 기존 테스트 케이스와 새로 생성한 테스트 케이스를 구분해서 제공하세요."""

                        try:
                            response = client.generate_content(prompt)
                            response_text = response.text
                            
                            # JSON 추출
                            if "```json" in response_text:
                                json_str = response_text.split("```json")[1].split("```")[0].strip()
                            else:
                                json_str = response_text.strip()
                            
                            ai_response = json.loads(json_str)
                            
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
                                            st.markdown("**테스트 단계:**")
                                            for step_num, step in enumerate(test_case['steps'], 1):
                                                st.markdown(f"{step_num}. {step}")
                            
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
                                    hide_index=True,
                                    column_config={
                                        "NO": st.column_config.NumberColumn(width="small"),
                                        "CATEGORY": st.column_config.TextColumn(width="medium"),
                                        "DEPTH 1": st.column_config.TextColumn(width="medium"),
                                        "DEPTH 2": st.column_config.TextColumn(width="medium"),
                                        "DEPTH 3": st.column_config.TextColumn(width="medium"),
                                        "PRE-CONDITION": st.column_config.TextColumn(width="large"),
                                        "STEP": st.column_config.TextColumn(width="large"),
                                        "EXPECT RESULT": st.column_config.TextColumn(width="large")
                                    }
                                )
                                
                                # 다운로드 버튼 (Excel 또는 CSV)
                                if EXCEL_AVAILABLE:
                                    # Excel 다운로드 버튼
                                    # BytesIO 객체 생성
                                    output = BytesIO()
                                    
                                    # Excel Writer 생성 및 DataFrame 쓰기
                                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                        df.to_excel(writer, index=False, sheet_name='테스트케이스')
                                        
                                        # 워크시트 가져오기
                                        workbook = writer.book
                                        worksheet = writer.sheets['테스트케이스']
                                        
                                        # 헤더 스타일 적용
                                        header_fill = PatternFill(start_color='4A90A4', end_color='4A90A4', fill_type='solid')
                                        header_font = Font(bold=True, color='FFFFFF')
                                        center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                                        thin_border = Border(
                                            left=Side(style='thin'),
                                            right=Side(style='thin'),
                                            top=Side(style='thin'),
                                            bottom=Side(style='thin')
                                        )
                                        
                                        # 헤더 행 스타일 적용
                                        for cell in worksheet[1]:
                                            cell.fill = header_fill
                                            cell.font = header_font
                                            cell.alignment = center_alignment
                                            cell.border = thin_border
                                        
                                        # 데이터 행 스타일 적용
                                        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
                                            for cell in row:
                                                cell.alignment = Alignment(vertical='center', wrap_text=True)
                                                cell.border = thin_border
                                        
                                        # 컬럼 너비 조정
                                        column_widths = {
                                            'A': 5,   # NO
                                            'B': 15,  # CATEGORY
                                            'C': 15,  # DEPTH 1
                                            'D': 20,  # DEPTH 2
                                            'E': 20,  # DEPTH 3
                                            'F': 30,  # PRE-CONDITION
                                            'G': 40,  # STEP
                                            'H': 40   # EXPECT RESULT
                                        }
                                        
                                        for column, width in column_widths.items():
                                            worksheet.column_dimensions[column].width = width
                                        
                                        # 행 높이 자동 조정
                                        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
                                            worksheet.row_dimensions[row[0].row].height = 30
                                    
                                    # BytesIO 객체의 포인터를 처음으로 되돌림
                                    output.seek(0)
                                    
                                    st.download_button(
                                        label="📥 테스트 케이스 Excel로 다운로드",
                                        data=output,
                                        file_name=f"test_cases_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                                else:
                                    # CSV 다운로드 버튼 (폴백)
                                    csv = df.to_csv(index=False, encoding='utf-8-sig')
                                    st.download_button(
                                        label="📥 테스트 케이스 CSV로 다운로드",
                                        data=csv,
                                        file_name=f"test_cases_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                        mime="text/csv"
                                    )
                                    st.info("💡 Excel 형식으로 다운로드하려면 'pip install openpyxl' 명령을 실행하세요.")
                            
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
            with st.expander(f"{history['timestamp'][:10]} - {history['query']}", expanded=(i==1)):
                st.write(f"**검색어:** {history['query']}")
                st.write(f"**시간:** {history['timestamp']}")
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
1. 테스트 케이스를 추가하거나 샘플 데이터를 로드하세요
2. **검색창**에 테스트하고 싶은 기능을 입력하세요 (예: "주문 QA", "로그인 테스트", "공동구매 메뉴")
3. **AI가 자동으로** 기존 테스트 케이스를 활용하고 신규 테스트 케이스를 생성합니다
4. 생성된 테스트 케이스는 표 형식으로 확인하고 Excel/CSV로 다운로드할 수 있습니다

### 🎯 주요 기능
- ✅ 테스트 케이스 학습 및 저장
- 🤖 AI 기반 연관 테스트 케이스 추론
- 📋 표 형식의 구조화된 테스트 케이스 생성
- 🔄 의존성 기반 테스트 순서 추천
- 📥 Excel(.xlsx) 또는 CSV 파일로 내보내기 기능
""")
