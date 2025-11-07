# QA 테스트 케이스 어시스턴트 설치 가이드

## 📋 필수 요구사항
- Python 3.8 이상
- Google Gemini API Key

## 🚀 설치 방법

### 1. 기본 패키지 설치 (필수)
```bash
pip install streamlit google-generativeai pandas
```

### 2. Excel 지원 패키지 설치 (선택)
Excel(.xlsx) 파일로 다운로드하려면 추가로 설치:
```bash
pip install openpyxl
```

또는 한 번에 모두 설치:
```bash
pip install -r requirements.txt
```

## 🔑 Google Gemini API 설정

### 방법 1: 환경 변수 설정 (추천)
```bash
# Windows
set GOOGLE_API_KEY=your-api-key-here

# Mac/Linux
export GOOGLE_API_KEY=your-api-key-here
```

### 방법 2: 코드에 직접 입력
코드의 `get_gemini_client()` 함수에서:
```python
api_key = "your-api-key-here"  # os.environ.get 대신 직접 입력
```

## ▶️ 실행 방법
```bash
streamlit run qa_test_assistant_improved.py
```

## 🎯 주요 기능
1. **테스트 케이스 관리**: 자유 형식으로 테스트 케이스 추가/삭제
2. **AI 분석**: Google Gemini를 활용한 지능형 테스트 케이스 추천
3. **구조화된 출력**: 표 형식(NO, CATEGORY, DEPTH 1-3, PRE-CONDITION, STEP, EXPECT RESULT)
4. **다운로드 지원**: 
   - Excel(.xlsx) - openpyxl 설치 시
   - CSV - 기본 지원

## 🔧 문제 해결

### "No module named 'openpyxl'" 오류
```bash
pip install openpyxl
```

### "GOOGLE_API_KEY 환경 변수가 설정되지 않았습니다" 오류
1. Google AI Studio에서 API 키 발급: https://makersuite.google.com/app/apikey
2. 위의 "Google Gemini API 설정" 참조

### Streamlit 실행 오류
```bash
pip install --upgrade streamlit
```

## 📝 사용 팁
- **샘플 데이터 로드**: 처음 사용 시 샘플 데이터를 로드하여 테스트
- **검색어 예시**: "주문 QA", "로그인 테스트", "공동구매 메뉴", "결제 프로세스"
- **Excel 다운로드**: openpyxl 설치 시 스타일이 적용된 Excel 파일 생성
- **CSV 폴백**: openpyxl 미설치 시 자동으로 CSV 다운로드 제공
