# v143 — NPC 진행상태 동기화 수정

- UI/CSS/이미지/레이아웃 변경 없음
- 고객사 정보 저장 직후 실제 런타임 `state.data.projectRegistered` 동기화
- 최백업 매니저 진입 조건을 단일 저장 키 `proposalAgentProject` 기준으로 통합
- 새로고침 후에도 저장된 고객사 정보를 런타임 상태에 복원
- RESET 범위는 v142와 동일하게 유지
