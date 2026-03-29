# MailPrioritizer — 오류 수정 및 코드 개선 이력

> 최종 업데이트: 2026-03-29

---

## 치명적 오류 수정 (Compile Error / Runtime Crash)

### 1. `ShowIdle()` 미정의 — 컴파일 오류

**파일:** `TaskPane/SummaryControl.cs:45`

`SummaryControl` 생성자에서 `ShowIdle()`을 호출했으나 해당 메서드가 정의되지 않았다.

**수정:** `ShowNoSelection()`으로 교체. 동일한 초기 상태(메일 미선택 안내)를 표시.

---

### 2. `OnAnalyzeNowClick` null 인자 전달 — NullReferenceException

**파일:** `TaskPane/SummaryControl.cs`

Task Pane의 "지금 분석하기" 버튼이 `ribbon.OnAnalyzeSelectedClick(sender, null)`을 호출하면서 `RibbonControlEventArgs`에 null을 전달해 런타임 오류 발생.

**수정:** Ribbon 의존을 제거하고, `MailProcessor.AnalyzeSingleAsync`를 직접 호출하는 독립적인 `async void OnAnalyzeNowClick` 메서드로 재작성.

---

## COM 객체 누수 수정 (5건)

COM 객체를 해제하지 않으면 Outlook 종료 시 프로세스가 잔류하거나, 장시간 사용 시 메모리 누수가 발생한다.

### 3. Explorer.SelectionChange — Selection, MailItem 누수

**파일:** `ThisAddIn.cs` → `Explorer_SelectionChange_Debounced`

`_explorer.Selection`과 `selection[1] as MailItem`이 finally 블록 없이 사용됨.

**수정:** `try/finally` + `ComHelper.ReleaseAll(mail, selection)` 추가.

---

### 4. OnAnalyzeSelectedClick — Selection, MailItem 누수

**파일:** `Ribbon/MailRibbon.cs` → `OnAnalyzeSelectedClick`

Ribbon 버튼 클릭 시 Selection과 MailItem COM 객체가 예외 발생 시 해제되지 않음.

**수정:** `try/finally` + `ComHelper.ReleaseAll(mail, selection)` 추가.

---

### 5. Application_NewMailEx — MailItem, NameSpace 누수

**파일:** `ThisAddIn.cs` → `AnalyzeNewMailAsync`

자동 분석 시 `Session.GetItemFromID()`로 획득한 MailItem과 NameSpace가 해제되지 않음.

**수정:** `try/finally` + `ComHelper.ReleaseAll(mail, session)` 추가.

---

### 6. LoadStoreList — Stores, Store, NameSpace 누수

**파일:** `Forms/SettingsForm.cs` → `LoadStoreList`

설정 창에서 Outlook 저장소 목록을 열거할 때 `Stores`, 개별 `Store`, `NameSpace` 객체가 해제되지 않음.

**수정:** 각 Store를 개별 `try/finally`로 해제, Stores와 NameSpace도 외부 `try/finally`로 해제.

---

### 7. GetTargetInbox — NameSpace, Store 누수

**파일:** `Services/FolderManager.cs` → `GetTargetInbox`

대상 저장소의 Inbox를 가져올 때 `Session`, `Store` 객체가 해제되지 않음.

**수정:** `try/finally` + `ComHelper.ReleaseAll(store, session)` 추가. 오류 시 기본 저장소로 폴백.

---

### 8. LlmService 연결 테스트 — HttpClient 누수

**파일:** `Forms/SettingsForm.cs` → `OnTestConnectionClick`

테스트용 `LlmService` 인스턴스가 `Dispose()` 없이 사용됨. 내부 `HttpClient`와 `SemaphoreSlim` 누수.

**수정:** `using (var svc = new LlmService(testConfig))` 블록으로 감싸 자동 해제.

---

## 코드 중복 제거 (4건)

### 9. `IsAnalyzed` 패턴 통합 (5곳 → 1곳)

**파일:** `Services/MailProcessor.cs` → `IsAnalyzed(object)`

`is bool b && b` 패턴이 5곳에 산재해 있었고, 일부 Outlook 버전에서 `olYesNo` 값이 `int`로 저장되어 `bool` 캐스팅 실패 가능성이 있었다.

**수정:** `internal static bool IsAnalyzed(object value)` 메서드로 통합. `bool`과 `int` 양쪽 모두 처리.

---

### 10. `GetSender` 패턴 통합 (3곳 → 1곳)

**파일:** `Services/MailProcessor.cs` → `GetSender(MailItem)`

`mail.SenderEmailAddress ?? mail.SenderName ?? ""` 인라인 코드가 3곳에 중복.

**수정:** `private static string GetSender(MailItem)` 메서드로 통합.

---

### 11. `SetViewState` 통합 (4곳 → 1곳)

**파일:** `TaskPane/SummaryControl.cs` → `SetViewState`

Task Pane의 상태 전환 시 10~15개 컨트롤의 `.Visible` 속성을 개별 설정하는 코드가 4곳에 걸쳐 ~50줄 반복.

**수정:** `SetViewState(showResult, showNotAnalyzed, showAnalyzing)` 단일 메서드로 통합. 각 호출부는 1줄로 단순화.

---

### 12. `ClearAnalysisFlag` 중앙화 (2곳 → 1곳)

**파일:** `Services/MailProcessor.cs` → `ClearAnalysisFlag(MailItem)`

SummaryControl과 MailProcessor에 각각 동일한 플래그 초기화 로직이 존재.

**수정:** `MailProcessor.ClearAnalysisFlag`를 `public static`으로 통합. SummaryControl에서 호출.

---

## 성능 개선 (5건)

### 13. DASL 필터 기반 미분석 메일 수집

**파일:** `Services/MailProcessor.cs` → `CollectUnanalyzedEntryIds`

**이전:** 받은편지함의 모든 메일을 순회하며 각각 `UserProperties.Find()`로 분석 여부 확인 — O(N) COM 호출.

**이후:** DASL 필터(`@SQL="...LLM_Analyzed" = 1`)로 분석된 메일 ID를 `HashSet`으로 수집 → 전체 메일에서 제외. DASL 실패 시 기존 방식으로 폴백.

---

### 14. 배치 크기와 동시성 제한 분리

**파일:** `Services/MailProcessor.cs` → `AnalyzeInboxAsync`

**이전:** 배치 크기 = `ConcurrentRequests` (기본 3). 3건 처리 후 다음 배치 시작 — LLM 파이프라인 유휴 시간 발생.

**이후:** 배치 크기 = `Max(ConcurrentRequests * 5, 15)`. SemaphoreSlim이 동시 요청을 제한하므로, 큰 배치에서 하나의 호출 완료 즉시 다음 호출 시작.

---

### 15. NewMailEx 병렬 처리

**파일:** `ThisAddIn.cs` → `Application_NewMailEx`

**이전:** 복수 메일 수신 시 순차 처리 (`foreach` + `await`).

**이후:** `Task.WhenAll(tasks)`로 병렬 처리. `SemaphoreSlim`이 동시성 제한.

---

### 16. SelectionChange 디바운스

**파일:** `ThisAddIn.cs`

**이전:** 키보드 화살표로 빠르게 메일 탐색 시 매번 COM 작업 실행 — UI 지연 및 CPU 낭비.

**이후:** 200ms `System.Windows.Forms.Timer` 디바운스 + 동일 EntryID 재선택 방지 (`_lastSelectedEntryId`).

---

### 17. 이중 Save 제거

**파일:** `Services/MailProcessor.cs`

**이전:** `TagSubject()` 내부에서 `mail.Save()` 호출 → `SaveAnalysisToMail()`에서 다시 `mail.Save()` 호출 — 불필요한 이중 디스크 I/O.

**이후:** `TagSubject()`에서 Save 제거. 호출 순서를 TagSubject → SaveAnalysisToMail로 변경하여 단일 Save로 통합.

---

## 설정 반영 오류 수정 (1건)

### 18. SemaphoreSlim 설정 변경 미반영

**파일:** `Services/LlmService.cs` → `ReloadConfig`

**이전:** 설정에서 동시 요청 수(`ConcurrentRequests`)를 변경해도 기존 `SemaphoreSlim` 인스턴스가 유지되어 효과 없음.

**이후:** `ReloadConfig`에서 현재 카운트 비교 후 새 `SemaphoreSlim` 생성, 기존 인스턴스 Dispose.

---

## 기타 품질 개선 (3건)

### 19. BatchProgress.Processed 계산 속성화

**파일:** `Services/MailProcessor.cs`

**이전:** `Processed` 필드를 수동으로 증가 — 다른 카운터와 불일치 가능.

**이후:** `get { return Urgent + High + Normal + Low + Failed; }` 계산 속성으로 변경. 항상 정확.

---

### 20. 중복 XML 주석 제거

**파일:** `Services/MailProcessor.cs:367`

`IsAnalyzed`의 `/// <summary>` 주석이 `ClearAnalysisFlag` 메서드 위에 오배치되어 두 개의 `<summary>` 태그가 연속.

**수정:** 오배치된 주석 라인 삭제.

---

### 21. 미사용 using 제거

**파일:** `Ribbon/MailRibbon.cs:2`

`using System.IO;`가 선언되어 있으나 파일 내에서 사용되지 않음 (`SaveFileDialog`는 `System.Windows.Forms`).

**수정:** 해당 줄 삭제.
