import streamlit as st
import json
from datetime import datetime
import google.generativeai as genai
import os

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
    page_title="QA 테스트 케이스 어시스턴트",
    page_icon="🧪",
    layout="wide"
)

st.title("🧪 AI 기반 QA 테스트 케이스 어시스턴트")
st.markdown("---")

# 사이드바 - 테스트 케이스 관리
with st.sidebar:
    st.header("📚 테스트 케이스 관리")
    
# 테스트 케이스 추가
with st.expander("➕ 새 테스트 케이스 추가", expanded=False):
    st.markdown("**자유 형식으로 작성하세요!**")
    st.caption("예: # 주문 취소 테스트\n- 로그인 후 주문 내역 확인\n- 취소 가능한 주문 선택...")
    
    test_content = st.text_area(
        "테스트 케이스 내용",
        height=300,
        placeholder="""예시:
# 카테고리: 주문

## 테스트: 회원 주문 취소 프로세스

**목적**: 주문 완료 후 취소가 정상적으로 처리되는지 확인

**테스트 순서**:
1. 로그인 상태 확인
2. 마이페이지 > 주문내역 이동
3. 최근 주문 중 취소 가능한 주문 선택
4. 취소 사유 선택
5. 주문 취소 버튼 클릭
6. 취소 완료 확인

**연관 기능**: 주문, 회원, 마이페이지, 결제환불

**추가 확인사항**:
- 취소 후 포인트/쿠폰 복구 확인
- 결제 취소 알림톡 발송 확인
"""
    )
        
    if st.button("테스트 케이스 추가"):
        if test_content.strip():
            new_test_case = {
                "id": len(st.session_state.test_cases) + 1,
                "category": "자유형식",  # 나중에 AI가 분석할 수 있음
                "name": f"테스트 케이스 {len(st.session_state.test_cases) + 1}",
                "description": test_content.strip(),  # 전체 내용을 여기에
                "steps": [],  # 빈 리스트
                "related_features": [],  # 빈 리스트
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            st.session_state.test_cases.append(new_test_case)
            save_test_cases_to_file(st.session_state.test_cases)
            st.success(f"✅ 테스트 케이스가 추가되었습니다!")
            st.rerun()
        else:
            st.error("테스트 케이스 내용을 입력해주세요.")
    
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
                
    st.subheader(f"📋 저장된 테스트 케이스 ({len(st.session_state.test_cases)}개)")
    
    if st.session_state.test_cases:
        for tc in st.session_state.test_cases:
            with st.expander(f"[{tc['category']}] {tc['name']}"):
                st.write(f"**설명:** {tc['description']}")
                st.write(f"**연관 기능:** {', '.join(tc['related_features'])}")
                if st.button(f"삭제", key=f"delete_{tc['id']}"):
                    st.session_state.test_cases = [t for t in st.session_state.test_cases if t['id'] != tc['id']]
                    save_test_cases_to_file(st.session_state.test_cases)  # 파일로 저장
                    st.rerun()

# 메인 영역
col1, col2 = st.columns([2, 1])

with col1:
    st.header("🔍 AI 기반 테스트 케이스 추천")
    
    if len(st.session_state.test_cases) == 0:
        st.warning("⚠️ 먼저 테스트 케이스를 추가하거나 샘플 데이터를 로드해주세요.")
    else:
        search_query = st.text_input(
            "테스트하고 싶은 기능을 입력하세요",
            placeholder="예: 주문 QA, 로그인 테스트, 결제 검증",
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
                        prompt = f"""당신은 QA 전문가입니다. 사용자가 "{search_query}"에 대한 테스트를 하려고 합니다.

다음은 현재 시스템에 등록된 테스트 케이스들입니다:

{test_cases_str}

사용자의 요청을 분석하고, 다음을 수행하세요:

1. 사용자가 테스트하려는 기능과 **직접 관련된** 테스트 케이스를 찾으세요
2. 그 기능이 작동하기 위해 **의존하는 다른 기능**들을 추론하세요
3. 의존하는 기능들의 테스트 케이스도 포함하세요
4. 논리적인 순서로 테스트 체크리스트를 만드세요

응답 형식:
```json
{{
  "reasoning": "왜 이런 테스트 케이스들이 필요한지 단계별 추론 과정 (한국어로 설명)",
  "recommended_test_cases": [
    {{
      "id": 테스트케이스ID,
      "reason": "이 테스트가 왜 필요한지 간단한 설명"
    }}
  ],
  "test_order": "추천하는 테스트 순서 설명",
  "additional_suggestions": "추가로 필요할 수 있는 테스트 제안"
}}
```

중요: 반드시 JSON 형식으로만 응답하세요."""

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
                            
                            # 추천된 테스트 케이스
                            st.markdown("### 📝 추천 테스트 체크리스트")
                            
                            recommended_ids = [tc["id"] for tc in ai_response.get("recommended_test_cases", [])]
                            recommended_cases = [tc for tc in st.session_state.test_cases if tc["id"] in recommended_ids]
                            
                            if recommended_cases:
                                for i, rec in enumerate(ai_response.get("recommended_test_cases", []), 1):
                                    test_case = next((tc for tc in st.session_state.test_cases if tc["id"] == rec["id"]), None)
                                    
                                    if test_case:
                                        with st.expander(f"✓ {i}. [{test_case['category']}] {test_case['name']}", expanded=True):
                                            st.markdown(f"**왜 필요한가?** {rec.get('reason', '')}")
                                            st.markdown(f"**설명:** {test_case['description']}")
                                            st.markdown("**테스트 단계:**")
                                            for step_num, step in enumerate(test_case['steps'], 1):
                                                st.markdown(f"{step_num}. {step}")
                            
                            # 테스트 순서 설명
                            if ai_response.get("test_order"):
                                st.markdown("### 🔄 권장 테스트 순서")
                                st.write(ai_response["test_order"])
                            
                            # 추가 제안
                            if ai_response.get("additional_suggestions"):
                                st.markdown("### 💡 추가 제안")
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
                rec_count = len(history['response'].get('recommended_test_cases', []))
                st.write(f"**추천된 테스트:** {rec_count}개")
    else:
        st.info("아직 검색 히스토리가 없습니다.")

# 하단 정보
st.markdown("---")
st.markdown("""
### 💡 사용 방법
1. 테스트 케이스를 추가하거나 샘플 데이터를 로드하세요
2. **검색창**에 테스트하고 싶은 기능을 입력하세요 (예: "주문 QA", "로그인 테스트")
3. **AI가 자동으로** 필요한 테스트 케이스를 추론하고 추천합니다
4. 연관된 기능의 테스트도 자동으로 포함됩니다

### 🎯 주요 기능
- ✅ 테스트 케이스 학습 및 저장
- 🤖 AI 기반 연관 테스트 케이스 추론
- 📋 자동 체크리스트 생성
- 🔄 의존성 기반 테스트 순서 추천
""")