# 프로젝트 가이드 (yoon - Billing System)

## 🛠 빌드 및 실행 명령
- 환경 설치: `pip install -r requirements.txt`
- 메인 로직 실행: `python main.py`
- 웹 인터페이스 실행: `python webapp.py`
- 테스트 실행: `pytest`

## 📊 프로젝트 구조 및 핵심 파일
- `billing/`: 청구 관련 핵심 비즈니스 로직 위치
- `schema/`: 데이터 모델 및 스키마 정의
- `excel_formatter.py`: 엑셀 파일 서식 및 스타일 지정 담당
- `main.py`: 전체 프로세스 엔트리 포인트
- `webapp.py`: 웹 기반 UI 제공

## 🤖 코딩 및 협업 규칙 (토큰 절약)
- **기존 스타일 준수**: 엑셀 서식 수정 시 `excel_formatter.py`에 구현된 기존 스타일 함수를 최대한 활용할 것.
- **간결한 응답**: 코드 수정 요청 시 파일 전체를 다시 쓰지 말고, 변경된 함수나 코드 블록만 출력할 것. 생략된 부분은 `# ... existing code` 주석으로 표시.
- **데이터 처리**: Pandas 데이터프레임 조작 시 기존 프로젝트의 네이밍 컨벤션을 따를 것.
- **테스트 코드**: 새로운 기능 추가 시 `tests/` 폴더에 관련 테스트 케이스를 함께 제안할 것.