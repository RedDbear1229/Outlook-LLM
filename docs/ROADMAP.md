# MailPrioritizer — 향후 구현 예정 아이템

> 최종 업데이트: 2026-04-05

---

## Phase 1: 즉시 효과 (구현 난이도 낮음, 체감 효과 높음)

### ~~R-01. 발신자/도메인 기반 규칙 엔진~~ (구현 완료)

`SenderRule` / `RulesConfig` 모델, `MailProcessor.ApplySenderRules()`, SettingsForm 4번째 탭 "발신자 규칙" DataGridView로 구현됨. 단건/배치 분석 모두 LLM 호출 전에 규칙 우선 적용.

---

### ~~R-02. 오프라인 재시도 큐~~ (구현 완료)

`Services/RetryQueue.cs`로 구현됨. `%AppData%/MailPrioritizer/retry_queue.json`에 실패 EntryID 저장, 최대 3회 재시도.

---

### ~~R-03. Task Pane 상태 기억~~ (구현 완료)

`DisplayConfig.TaskPaneVisible/TaskPaneWidth` 필드로 저장/복원. 열림·닫힘은 `VisibleChanged` 이벤트로 즉시 저장, 너비는 1초 간격 폴링 타이머로 변화 감지 후 저장 (`CustomTaskPane`에 `WidthChanged` 이벤트 없음).

---

### ~~R-04. 첨부파일 메타데이터 활용~~ (구현 완료)

`MailProcessor.GetAttachmentNames()`로 구현됨. `Processing.IncludeAttachmentNames` 설정으로 제어.

---

## Phase 2: 핵심 개선 (중간 난이도)

### R-05. 분석 결과 로컬 인덱스 (SQLite)

분석 결과를 SQLite DB에 별도 저장하여 검색/통계/필터를 고속화한다.

**현재 문제:** 통계 조회(`CollectInboxStats`)와 CSV 내보내기 시 받은편지함 전체 COM 순회 필요 — 메일이 수천 건이면 수십 초 소요.
**구현 방식:** `%AppData%/MailPrioritizer/index.db`에 `{EntryID, Subject, Sender, Priority, Summary, AnalyzedAt}` 테이블. 분석 완료 시 INSERT/UPDATE. 통계는 SQL `COUNT/GROUP BY`.
**추가 이점:** Outlook 재설치/프로필 변경 시에도 분석 이력 보존, "긴급 메일 검색" UI 제공 가능.
**주의:** NuGet `System.Data.SQLite` 의존성 추가 필요.
**예상 변경 파일:** 신규 `Services/IndexDatabase.cs`, `Services/MailProcessor.cs`, `MailPrioritizer.csproj`

---

### ~~R-06. 사용자 피드백 수집 및 정확도 통계~~ (구현 완료)

`Services/FeedbackStore.cs`로 구현됨. `feedback.json`에 수동 변경 이력 저장, `GetStats()` 메서드로 정확도/오분류 패턴 통계 제공.

---

### ~~R-07. 분석 결과 버전 관리~~ (구현 완료)

UserProperty `LLM_ModelName`, `LLM_AnalyzedAt` 추가됨. `MailPropertyNames.cs`에 상수 정의.

---

### R-08. 키보드 단축키

자주 사용하는 기능에 전역 키보드 단축키를 바인딩한다.

| 단축키 | 기능 |
|--------|------|
| `Ctrl+Shift+M` | 선택 메일 분석 |
| `Ctrl+Shift+1~4` | 우선순위 즉시 변경 (긴급/높음/보통/낮음) |

**구현 방식:** `Application.ActiveExplorer().CommandBars` 또는 Ribbon `getKeytip` 콜백.
**예상 변경 파일:** `Ribbon/MailRibbon.xml`, `Ribbon/MailRibbon.cs`

---

## Phase 3: 장기 과제 (높은 난이도)

### R-09. 다국어 지원 (i18n)

UI 문자열을 리소스 파일로 분리하여 한국어/영어/일본어 등 다국어를 지원한다.

**현재 상태:** 모든 UI 텍스트가 코드에 하드코딩 (~80개 문자열). 프롬프트 응답 언어도 한국어 고정.
**구현 방식:** `Resources.resx` (한국어 기본) + `Resources.en.resx` (영어) 등. 프롬프트 템플릿도 언어별 분리.
**예상 변경 파일:** 모든 UI 파일, 신규 `Properties/Resources.resx`

---

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

### R-12. 증분 분석 최적화

마지막 분석 시점 이후 수신된 메일만 대상으로 하여 수집 단계를 최소화한다.

**현재 상태:** "전체 분류" 실행 시 매번 받은편지함 전체를 순회하여 미분석 메일 식별.
**구현 방식:** `config.json`에 `lastBatchAnalyzedAt` 타임스탬프 저장. DASL 필터에 `ReceivedTime >= lastBatchAnalyzedAt` 조건 추가.
**예상 변경 파일:** `Models/AppConfig.cs`, `Services/MailProcessor.cs`, `Config/ConfigManager.cs`

---

### R-13. Health Check 상태 표시

API 연결 상태를 Ribbon 또는 Task Pane에 상시 표시하여 문제를 즉시 인지할 수 있게 한다.

**구현 방식:**
- 마지막 API 호출 성공/실패 시각과 상태를 `LlmService`에 기록
- Task Pane 하단에 상태 아이콘 표시 (초록/빨강/회색)
- 선택적: 주기적 핑 (5분 간격, 설정으로 비활성화 가능)

**예상 변경 파일:** `Services/LlmService.cs`, `TaskPane/SummaryControl.cs`

---

## 우선순위 매트릭스

| 아이템 | 난이도 | 효과 | 상태 |
|--------|--------|------|------|
| ~~R-01 발신자/도메인 규칙~~ | 낮음 | 높음 | **완료** |
| ~~R-02 오프라인 재시도 큐~~ | 낮음 | 중간 | **완료** |
| ~~R-03 Task Pane 상태 기억~~ | 낮음 | 중간 | **완료** |
| ~~R-04 첨부파일 메타데이터~~ | 낮음 | 중간 | **완료** |
| R-05 SQLite 로컬 인덱스 | 중간 | 높음 | 미구현 |
| ~~R-06 피드백 수집/정확도~~ | 중간 | 높음 | **완료** |
| ~~R-07 분석 결과 버전 관리~~ | 중간 | 중간 | **완료** |
| R-08 키보드 단축키 | 중간 | 중간 | 미구현 |
| R-09 다국어 지원 | 높음 | 중간 | 미구현 |
| R-10 단위 테스트 | 높음 | 높음 | 미구현 |
| R-11 LLM 스트리밍 | 높음 | 중간 | 미구현 |
| R-12 증분 분석 최적화 | 중간 | 중간 | 미구현 |
| R-13 Health Check | 중간 | 낮음 | 미구현 |
