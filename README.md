# WebSquare GridView & TabControl Components Sample

Inswave WebSquare5 호환 그리드(GridView) 및 탭 컨트롤(TabControl) 가상 시뮬레이션 및 커스텀 기능 구현 샘플 프로젝트입니다.

## 📌 주요 제공 기능 및 샘플 페이지

### 1. [sample.html](sample.html) - 메인 그리드 샘플
- **CheckComboBox 실시간 검색 (Filtering)**: 드롭다운 상단 검색 입력 폼을 통해 대상 항목 검색 지원
- **그리드 확대 / 축소 / 100% 리셋**: 그리드 화면 비율 조절 컨트롤 (70% ~ 160%)
- **작업유형 선택 연동**: 1열(작업유형) 변경 시 숨김 열(작업구분) 자동 갱신 및 2열(작업대상) 초기화
- **다중 선택 항목 세로 줄바꿈(`<br/>`) 표기**: 비편집 모드 시 세로형 배지로 다중 선택 아이템 표시

### 2. [header_sample.html](header_sample.html) - 다중 헤더 & 사원 조회 팝업 샘플
- **2행 다중 헤더 (Multi-Header)**: 1행(근무 조: 통상조, 1조, 2조, 3조) / 2행(조원, 사번)
- **조원 셀 호버 삭제 기능**: 마우스 호버 시 삭제 버튼(✕) 동적 표시 및 제거
- **클릭 편집 & 사원 조회 연동**: [조원] 셀 클릭 후 이름 입력 후 `Enter` 키 조절
  - **1명 단일 매칭 시**: 사번(`EMP-XXXX`) 자동 채움
  - **0명 또는 2명 이상 매칭 시**: 사원 조회 검색 팝업창(Modal) 띄움

### 3. [tab_main.html](tab_main.html) - 3-Tab 서브페이지 & 통합 엑셀 리포트 다운로드
- **TabControl 서브페이지 로더**: 3개 독립 서브페이지([sub_tab1.html](sub_tab1.html), [sub_tab2.html](sub_tab2.html), [sub_tab3.html](sub_tab3.html)) 호출
- **멀티시트(Multi-Sheet) 엑셀 통합 다운로드**: 3개 그리드 데이터를 1개의 엑셀 워크북(시트1, 시트2, 시트3)으로 일괄 다운로드
- **엑셀 보고서 서식 양식 적용**: 타이틀 밴드, 헤더 배경색(`#2563EB`), 테두리선, 금액/수량 정렬, 하단 통계 합계 행 포함

---

## 🚀 실행 방법

### 로컬 웹 서버 실행
```bash
python3 server.py
```
서버 실행 후 브라우저에서 `http://localhost:8080/tab_main.html` 접속
