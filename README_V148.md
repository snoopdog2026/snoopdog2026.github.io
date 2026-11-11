# v148

Changes limited to RESET behavior and storyline loading popup.

- Clears previous browser-side company/project registration once on first launch of this build.
- RESET always clears browser data even if Google Sheet JSONP is slow or fails.
- Restores a custom cute reset dialog: "당신은 오늘 막 입사한 인턴사원으로 돌아갑니다..."
- Google Sheet reset is attempted in parallel and no longer blocks local reset.
- Enlarges the "진짜로 AI가 스토리 생성 중..." popup.
- Storyline section heights and all other UI remain unchanged.

Apps Script replacement: not required for this build if v141+ reset scope code is already deployed.
