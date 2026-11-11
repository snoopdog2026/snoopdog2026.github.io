# v142 — Storage/RESET 구조 안정화 (UI Freeze)

UI, 캐릭터, 이미지, CSS, 레이아웃, 비율은 변경하지 않았습니다.

## 변경 사항
- 프로젝트 저장 키를 버전별 키에서 `proposalAgentProject` 단일 키로 통합
- 프로젝트/제안서/리뷰/AI 대화 히스토리 초기화 로직을 `ProposalAgentStore`로 통합
- 새 빌드 최초 실행 시 이전 프로젝트 등록 및 히스토리 제거
- RESET 성공 후에만 브라우저 프로젝트/히스토리 제거
- RESET Google Sheet 범위 고정
  - `Project_Input!B2:B7`
  - `Document` 2행 이하
  - `Page` 2행 이하
- `Research Board`, `Storyline`, `Partner PPT`, `Review` 및 기타 시트는 수정하지 않음

## Apps Script
`APPS_SCRIPT_CODE.gs`를 기존 Apps Script에 전체 교체한 뒤 새 버전으로 배포해야 합니다.
