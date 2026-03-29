# MailPrioritizer — 구현 완료 기능 목록

> 최종 업데이트: 2026-03-29

---

## 1. LLM 기반 메일 분석 (핵심 기능)

**파일:** `Services/LlmService.cs`, `Services/MailProcessor.cs`

메일의 제목, 본문(plaintext), 발신자를 LLM에 전송하여 3가지 결과를 받는다:
- **요약** (2-3문장 한국어)
- **우선순위** (Urgent / High / Normal / Low)
- **판단 근거** (1문장)

분석 결과는 Outlook MailItem의 UserProperty에 저장되어 Outlook 재시작 후에도 유지된다:
- `LLM_Priority` (Text) — 우선순위 문자열
- `LLM_Summary` (Text) — 요약
- `LLM_PriorityReason` (Text) — 판단 근거
- `LLM_Analyzed` (YesNo) — 분석 완료 플래그

---

## 2. 듀얼 LLM API 지원

**파일:** `Services/LlmService.cs`

Endpoint URL로 API 형식을 자동 감지하여 두 가지 API를 지원한다:

| 조건 | API 형식 | 인증 | 엔드포인트 |
|------|---------|------|-----------|
| URL에 "anthropic" 포함 | Claude Messages API | `x-api-key` 헤더 | `{endpoint}/messages` |
| 그 외 | OpenAI Chat Completions | `Bearer` 토큰 | `{endpoint}/chat/completions` |

- 지수 백오프 재시도: 429/500/502/503 응답 시 2초→4초→8초, 최대 3회
- `SemaphoreSlim` 기반 동시 요청 제한 (기본 3개)
- JSON 응답 파싱: `` ```json ``` `` 펜스 및 `{ }` 블록 자동 추출
- 파싱 실패 시 Normal 우선순위 폴백

---

## 3. 선택 메일 단건 분석

**파일:** `Ribbon/MailRibbon.cs` → `OnAnalyzeSelectedClick`

Ribbon의 "선택 메일 요약" 버튼으로 현재 선택된 메일 1건을 즉시 분석한다.
- Task Pane에 "분석 중..." 표시 → 완료 후 결과 표시
- 이미 분석된 메일은 저장된 결과를 즉시 반환 (API 호출 없음)
- API 미설정 시 안내 메시지 → 설정 창 자동 오픈

---

## 4. 받은편지함 일괄 분석

**파일:** `Ribbon/MailRibbon.cs` → `OnAnalyzeAllClick`, `Services/MailProcessor.cs` → `AnalyzeInboxAsync`

받은편지함의 미분석 메일을 모두 찾아 배치 분석한다.

**3-Phase 파이프라인:**
1. **수집:** DASL 필터로 분석 완료 메일을 빠르게 식별 → 미분석 EntryID 목록 구성
2. **Phase 1 (STA):** COM에서 메일 데이터 추출 → `MailDataItem` POCO로 변환
3. **Phase 2 (ThreadPool):** `Task.WhenAll` + `SemaphoreSlim`으로 LLM 병렬 호출
4. **Phase 3 (STA):** UserProperty 저장 + 제목 태그 + 폴더 이동 (순차)

**배치 특성:**
- 배치 크기: `Max(ConcurrentRequests * 5, 15)` — 파이프라인 효율 극대화
- 진행률 표시: `ProgressForm`에 실시간 카운트 (긴급/높음/보통/낮음/실패)
- 취소 지원: `CancellationToken`으로 즉시 중단, 처리 완료분은 유지
- 연속 실패 감지: 5건 연속 실패 시 사용자에게 계속 여부 확인
- 완료 보고: 우선순위별 건수 + 실패 메일 제목 (최대 10건)

---

## 5. 전체 재분석

**파일:** `Ribbon/MailRibbon.cs` → `OnReanalyzeAllClick`, `Services/MailProcessor.cs` → `ResetAllAnalysisFlags`

모든 메일의 `LLM_Analyzed` 플래그를 초기화한 뒤 일괄 분석을 재실행한다.
- 실행 전 확인 대화상자 표시
- 리셋 건수 로깅 후 `OnAnalyzeAllClick`과 동일 흐름 실행

---

## 6. 신규 메일 자동 분석

**파일:** `ThisAddIn.cs` → `Application_NewMailEx`, `AnalyzeNewMailAsync`

설정에서 "새 메일 수신 시 자동 분석" 활성화 시 동작한다.
- `Application.NewMailEx` 이벤트로 수신 즉시 감지
- 복수 메일 동시 수신 시 `Task.WhenAll`로 병렬 분석
- 현재 Task Pane에서 보고 있는 메일이면 자동 갱신 (`RefreshTaskPaneIfSelected`)
- API 미설정 시 자동으로 비활성화

---

## 7. Custom Task Pane (우측 분석 패널)

**파일:** `TaskPane/SummaryControl.cs`

Explorer 우측에 고정되는 320px 패널. 메일 선택 시 자동 갱신된다.

**상태별 표시:**
| 상태 | 표시 내용 |
|------|----------|
| 메일 미선택 | "메일을 선택하면 분석 결과가 표시됩니다." |
| 미분석 메일 | "아직 분석되지 않은 메일입니다." + "지금 분석하기" 버튼 |
| 분석 중 | "분석 중..." (파란색) |
| 분석 완료 | 우선순위 배지(색상) + 요약 + 판단 근거 |
| 오류 | "오류: {메시지}" |

**인터랙션:**
- 우선순위 수동 변경: ComboBox로 즉시 변경 → UserProperty 업데이트 + 폴더 자동 이동
- 재분석 버튼: 플래그 초기화 후 LLM 재호출
- 폴더 이동 버튼: 현재 우선순위에 해당하는 폴더로 수동 이동
- 통계 새로고침: 받은편지함 전체/분석됨/우선순위별 건수 표시

**SelectionChange 최적화:**
- 200ms 디바운스 타이머로 빠른 키보드 탐색 시 마지막 선택만 처리
- 동일 메일 재선택 시 `_lastSelectedEntryId` 비교로 COM 작업 생략

---

## 8. 우선순위 폴더 자동 분류

**파일:** `Services/FolderManager.cs`

분석 완료 후 메일을 우선순위별 하위 폴더로 자동 이동한다.

**폴더 구조:**
```
받은편지함/
  └── 우선순위/          ← FolderPrefix (설정 변경 가능)
      ├── 긴급/          ← Urgent
      ├── 높음/          ← High
      ├── 보통/          ← Normal
      └── 낮음/          ← Low
```

- 폴더 미존재 시 자동 생성 (`FindOrCreateFolder`)
- 접두사 및 하위 폴더명 설정 가능 (기본: 한글)
- 이모지 폴더명 미사용 (Exchange/IMAP 호환성)
- 설정에서 자동 이동 비활성화 가능

---

## 9. 다중 저장소(계정) 지원

**파일:** `Services/FolderManager.cs` → `GetTargetInbox`, `Forms/SettingsForm.cs`

Outlook에 등록된 여러 메일 계정 중 분석 대상을 선택할 수 있다.
- 설정 > 분류 설정 탭의 "대상 계정" ComboBox에서 선택
- 기본값: `(기본 저장소)` — Outlook 기본 계정
- 선택한 저장소 미발견 시 기본 저장소로 자동 폴백 + 경고 로그

---

## 10. 분석 결과 CSV 내보내기

**파일:** `Ribbon/MailRibbon.cs` → `OnExportResultsClick`, `Services/MailProcessor.cs` → `ExportAnalyzedMails`

분석 완료된 메일의 결과를 CSV 파일로 내보낸다.
- 컬럼: Subject, Sender, Priority, Summary, PriorityReason
- 파일명 기본값: `MailPrioritizer_Export_yyyyMMdd.csv`
- UTF-8 인코딩, CSV 특수문자(쉼표/따옴표/줄바꿈) 이스케이프 처리
- 내보내기 완료 시 건수 + 파일 경로 표시

---

## 11. 환경설정 다이얼로그

**파일:** `Forms/SettingsForm.cs`, `Config/ConfigManager.cs`

3탭 구성의 설정 창:

**탭 1 — API 설정:**
- Endpoint URL, API Token, 모델 이름
- "연결 테스트" 버튼 (실제 LLM 호출로 검증, 결과 표시)

**탭 2 — 프롬프트:**
- 시스템 프롬프트 편집기 (Consolas 폰트, 스크롤)
- "기본값으로 복원" 버튼
- JSON 형식 유지 경고 표시

**탭 3 — 분류 설정:**
- 폴더 접두사, 자동 폴더 이동, 제목 태그 삽입
- 본문 최대 길이 (500~20000)
- 대상 계정 선택
- 새 메일 자동 분석 토글

설정은 `%AppData%/MailPrioritizer/config.json`에 저장. API 토큰은 Windows DPAPI로 암호화.

---

## 12. 파일 기반 로깅

**파일:** `Utils/Logger.cs`

운영 환경에서의 문제 진단을 위한 구조화된 로깅 시스템.
- 경로: `%AppData%/MailPrioritizer/Logs/MailPrioritizer-yyyy-MM-dd.log`
- 포맷: `[2026-03-29 14:30:45.123] [INFO] 메시지`
- 4단계 레벨: Debug, Info, Warn, Error
- Error 레벨: 예외 타입, 메시지, 스택 트레이스, InnerException 포함
- `lock` + `File.AppendAllText`로 스레드 안전
- 7일 초과 로그 자동 삭제 (`CleanupOldLogs`)
- 로거 실패 시 Add-in 정상 동작 보장 (모든 내부 예외 묵살)

**로깅 범위:** ThisAddIn 생명주기, 설정 로드/저장, LLM API 호출/응답, 배치 처리 진행, 폴더 생성/이동, UI 오류
**민감 정보 보호:** API 토큰, 메일 본문, DPAPI 바이트는 절대 로깅하지 않음

---

## 13. 제목 우선순위 태그

**파일:** `Services/MailProcessor.cs` → `TagSubject`

설정 활성화 시 메일 제목 앞에 우선순위 태그를 삽입한다.
- 예: `[긴급] 서버 장애 보고`, `[높음] 계약서 검토 요청`
- 이미 `[` 로 시작하는 제목은 중복 삽입 방지
- `mail.Save()` 호출 없이 태그만 설정 — 호출자가 일괄 Save
