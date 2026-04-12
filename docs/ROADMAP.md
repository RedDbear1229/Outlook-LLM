# MailPrioritizer — 향후 구현 예정 아이템

> 최종 업데이트: 2026-04-12

---

## 완료된 기능

### ~~R-01. 발신자/도메인 기반 규칙 엔진~~ (구현 완료)

`SenderRule` / `RulesConfig` 모델, `MailProcessor.ApplySenderRules()`, SettingsForm 4번째 탭 "발신자 규칙" DataGridView로 구현됨. 단건/배치 분석 모두 LLM 호출 전에 규칙 우선 적용.

---

### ~~R-02. 오프라인 재시도 큐~~ (구현 완료)

`Services/RetryQueue.cs`로 구현됨. `%AppData%/MailPrioritizer/retry_queue.json`에 실패 EntryID 저장, 최대 3회 재시도. 큐 최대 크기 1000건 제한 추가.

---

### ~~R-03. Task Pane 상태 기억~~ (구현 완료)

`DisplayConfig.TaskPaneVisible/TaskPaneWidth` 필드로 저장/복원. 열림·닫힘은 `VisibleChanged` 이벤트로 즉시 저장, 너비는 1초 간격 폴링 타이머로 변화 감지 후 저장.

---

### ~~R-04. 첨부파일 메타데이터 활용~~ (구현 완료)

`MailProcessor.GetAttachmentNames()`로 구현됨. `Processing.IncludeAttachmentNames` 설정으로 제어.

---

### ~~R-05. 분석 결과 로컬 인덱스 (SQLite)~~ (구현 완료)

`Services/IndexDatabase.cs`로 구현됨. `%AppData%/MailPrioritizer/index.db`에 분석 결과 저장. 통계 조회와 CSV 내보내기를 COM 순회 없이 DB 쿼리로 처리.

---

### ~~R-06. 사용자 피드백 수집 및 정확도 통계~~ (구현 완료)

`Services/FeedbackStore.cs`로 구현됨. `feedback.json`에 수동 변경 이력 저장, `GetStats()` 메서드로 정확도/오분류 패턴 통계 제공. 90일 보관 정책 추가.

---

### ~~R-07. 분석 결과 버전 관리~~ (구현 완료)

UserProperty `LLM_ModelName`, `LLM_AnalyzedAt` 추가됨. `MailPropertyNames.cs`에 상수 정의.

---

### ~~R-08. 키보드 단축키~~ (구현 완료)

`Utils/KeyboardShortcutManager.cs`로 구현됨. Win32 `RegisterHotKey` + 메시지 전용 `NativeWindow` 방식.

| 단축키 | 기능 |
|--------|------|
| `Ctrl+Shift+M` | 선택 메일 분석 |
| `Ctrl+Shift+1~4` | 우선순위 즉시 변경 (긴급/높음/보통/낮음) |

---

### ~~R-12. 증분 분석 최적화~~ (구현 완료)

`ProcessingConfig.LastBatchAnalyzedAt` 타임스탬프로 구현됨. 마지막 배치 이후 수신 메일만 스캔. 설정 저장 시 타임스탬프 보존 버그도 수정됨(BUGFIXES #21).

---

### ~~R-13. Health Check 상태 표시~~ (구현 완료)

`LlmService`에 마지막 API 호출 성공/실패 시각을 기록. Task Pane 하단에 상태 아이콘 표시 (초록/빨강/회색). 메일 선택 시 자동 갱신.

---

### ~~Task Pane 색상 테마~~ (구현 완료)

`TaskPane/ThemePalette.cs`로 구현됨. Light / Grey / Dark 3가지 테마, 설정에서 `DisplayConfig.ThemeName`으로 제어. 설정 저장 즉시 반영.

---

### ~~C# GUI 설치 관리자~~ (구현 완료)

`MailPrioritizer.Installer/` 프로젝트. PowerShell 스크립트 방식을 대체. 설치/복구/제거 UI 제공, `deploy-clickonce.ps1` 빌드 시 자동 패키징.

---

## 미구현 항목

### R-10. 단위 테스트 도입

핵심 로직에 대한 테스트 커버리지를 확보하여 리팩토링 안전성을 높인다.

**즉시 테스트 가능 (순수 함수):**
- `PriorityExtensions.FromString` — 문자열→enum 매핑
- `LlmService.ParseLlmResponse` (private → internal로 변경 필요) — JSON 파싱
- `MailProcessor.IsAnalyzed` — bool/int 양쪽 처리
- `MailProcessor.CsvEscape` — 특수문자 이스케이프
- `ConfigManager.EncryptToken/DecryptToken` — DPAPI 라운드트립

**인터페이스 추출 필요:**
- `ILlmService` — LLM 호출 모킹
- `IFolderManager` — 폴더 작업 모킹
- COM 접근 추상화 — Outlook 없이 MailProcessor 테스트

**프로젝트 구조:** `MailPrioritizer.Tests` 프로젝트 추가 (NUnit 또는 MSTest).

---

### R-11. LLM 응답 스트리밍

Claude API의 SSE 스트리밍을 활용하여 Task Pane에서 요약이 실시간으로 표시되도록 한다.

**현재 상태:** 전체 응답이 완료될 때까지 "분석 중..." 표시 — 긴 메일은 10초+ 대기.
**구현 방식:** `HttpClient` + `HttpCompletionOption.ResponseHeadersRead` + SSE 파싱. `SummaryControl`에 부분 텍스트 업데이트 콜백.
**주의:** OpenAI API도 스트리밍 지원하므로 양쪽 구현 필요. 배치 분석에는 적용하지 않음 (단건 분석만).

---

## 우선순위 매트릭스

| 아이템 | 난이도 | 효과 | 상태 |
|--------|--------|------|------|
| ~~R-01 발신자/도메인 규칙~~ | 낮음 | 높음 | **완료** |
| ~~R-02 오프라인 재시도 큐~~ | 낮음 | 중간 | **완료** |
| ~~R-03 Task Pane 상태 기억~~ | 낮음 | 중간 | **완료** |
| ~~R-04 첨부파일 메타데이터~~ | 낮음 | 중간 | **완료** |
| ~~R-05 SQLite 로컬 인덱스~~ | 중간 | 높음 | **완료** |
| ~~R-06 피드백 수집/정확도~~ | 중간 | 높음 | **완료** |
| ~~R-07 분석 결과 버전 관리~~ | 중간 | 중간 | **완료** |
| ~~R-08 키보드 단축키~~ | 중간 | 중간 | **완료** |
| R-10 단위 테스트 | 높음 | 높음 | **미구현** |
| R-11 LLM 스트리밍 | 높음 | 중간 | **미구현** |
| ~~R-12 증분 분석 최적화~~ | 중간 | 중간 | **완료** |
| ~~R-13 Health Check~~ | 중간 | 낮음 | **완료** |
| ~~Task Pane 색상 테마~~ | 낮음 | 중간 | **완료** |
| ~~C# GUI 설치 관리자~~ | 중간 | 중간 | **완료** |
