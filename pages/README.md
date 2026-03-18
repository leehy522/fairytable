# 요정비닐 스마트 시스템

## 📁 프로젝트 구조

```
yojung_system/
│
├── app.py                  ← 진입점 (메인 라우터)
├── auth.py                 ← 로그인 인증 모듈
├── requirements.txt        ← 패키지 목록
│
└── pages/                  ← 메뉴별 페이지 모듈
    ├── __init__.py
    ├── product_status.py   ← 🏷️ 요정비닐 상품 현황
    ├── milkrun_ppt.py      ← 🚚 밀크런 PPT 변환
    ├── invoice.py          ← 📦 택배 송장 변환
    ├── cost_simulator.py   ← 🏭 원가 시뮬레이터
    ├── market_index.py     ← 📈 시장 지표 분석
    └── narajangte.py       ← 🏛️ 나라장터 입찰
```

## ▶️ 실행 방법

```bash
# 패키지 설치
pip install -r requirements.txt

# 앱 실행
streamlit run app.py
```

## 🔑 설계 원칙

| 파일 | 역할 |
|------|------|
| `app.py` | 페이지 설정, 로그인 체크, 사이드바 메뉴, 라우팅만 담당 |
| `auth.py` | 로그인 UI + 세션 관리만 담당 |
| `pages/*.py` | 각 메뉴는 `render()` 함수 하나만 외부에 노출 |

## 🛠️ 메뉴 추가 방법

1. `pages/` 아래 새 파일 생성 (예: `pages/new_menu.py`)
2. `render()` 함수 구현
3. `pages/__init__.py`에 import 추가
4. `app.py`의 `MENU_MAP`에 항목 추가

```python
# app.py MENU_MAP 예시
MENU_MAP = {
    ...
    "🆕 새 메뉴": new_menu,
}
```

## ⚙️ 주요 변경 사항 (v기존 → 분리 버전)

- **단일 파일 → 모듈 분리**: 900줄 단일 파일을 7개 파일로 분리
- **중복 함수 제거**: `load_google_sheet_data` 중복 정의 제거
- **auth 분리**: 로그인 로직을 `auth.py`로 독립
- **내부 헬퍼 은닉**: `_`prefix로 모듈 내부 함수와 공개 API 구분
- **밀크런 PPT 리팩토링**: `_extract_pdf_data`, `_build_pptx` 함수로 분리하여 테스트 가능하게 구성
