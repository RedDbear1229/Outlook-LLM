# MailPrioritizer — 오류 수정 및 코드 개선 이력

> 최종 업데이트: 2026-04-12

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

### 3. `CsvEscape` null 입력 — NullReferenceException

**파일:** `Services/MailProcessor.cs` → `CsvEscape`

메일 필드(요약, 판단 근거 등)가 null인 경우 `.Contains()` 호출에서 런타임 오류 발생.

**수정:** 메서드 최상단에 `if (value == null) return "";` null 가드 추가.

---

## COM 객체 누수 수정 (13건)

COM 객체를 해제하지 않으면 Outlook 종료 시 프로세스가 잔류하거나, 장시간 사용 시 메모리 누수가 발생한다.

### 4. Explorer.SelectionChange — Selection, MailItem 누수

**파일:** `ThisAddIn.cs` → `Explorer_SelectionChange_Debounced`

`_explorer.Selection`과 `selection[1] as MailItem`이 finally 블록 없이 사용됨.

**수정:** `try/finally` + `ComHelper.ReleaseAll(mail, selection)` 추가.

---

### 5. OnAnalyzeSelectedClick — Selection, MailItem 누수

**파일:** `Ribbon/MailRibbon.cs` → `OnAnalyzeSelectedClick`

Ribbon 버튼 클릭 시 Selection과 MailItem COM 객체가 예외 발생 시 해제되지 않음.

**수정:** `try/finally` + `ComHelper.ReleaseAll(mail, selection)` 추가.

---

### 6. Application_NewMailEx — MailItem, NameSpace 누수

**파일:** `ThisAddIn.cs` → `AnalyzeNewMailAsync`

자동 분석 시 `Session.GetItemFromID()`로 획득한 MailItem과 NameSpace가 해제되지 않음.

**수정:** `try/finally` + `ComHelper.ReleaseAll(mail, session)` 추가.

---

### 7. LoadStoreList — Stores, Store, NameSpace 누수

**파일:** `Forms/SettingsForm.cs` → `LoadStoreList`

설정 창에서 Outlook 저장소 목록을 열거할 때 `Stores`, 개별 `Store`, `NameSpace` 객체가 해제되지 않음.

**수정:** 각 Store를 개별 `try/finally`로 해제, Stores와 NameSpace도 외부 `try/finally`로 해제.

---

### 8. GetTargetInbox — NameSpace, Store 누수

**파일:** `Services/FolderManager.cs` → `GetTargetInbox`

대상 저장소의 Inbox를 가져올 때 `Session`, `Store` 객체가 해제되지 않음.

**수정:** `try/finally` + `ComHelper.ReleaseAll(store, session)` 추가. 오류 시 기본 저장소로 폴백.

---

### 9. LlmService 연결 테스트 — HttpClient 누수

**파일:** `Forms/SettingsForm.cs` → `OnTestConnectionClick`

테스트용 `LlmService` 인스턴스가 `Dispose()` 없이 사용됨. 내부 `HttpClient`와 `SemaphoreSlim` 누수.

**수정:** `using (var svc = new LlmService(testConfig))` 블록으로 감싸 자동 해제.

---

### 10. AnalyzeInboxAsync Phase 1 — NameSpace 루프 누수

**파일:** `Services/MailProcessor.cs` → `AnalyzeInboxAsync` Phase 1

`outlookApp.Session.GetItemFromID()`를 메일 아이템 루프 내부에서 매번 호출하여 Session COM 객체가 아이템마다 새로 획득되고 해제되지 않음. 수백 건 처리 시 수백 개의 누수 발생.

**수정:** Session을 루프 외부에서 1회 획득하고 Phase 완료 후 finally에서 1회 해제.

```csharp
Outlook.NameSpace p1Session = outlookApp.Session;
try
{
    foreach (var id in batch)
    {
        // p1Session.GetItemFromID(id) 사용
    }
}
finally
{
    ComHelper.ReleaseAll(p1Session);
}
```

---

### 11. AnalyzeInboxAsync Phase 3 — NameSpace 루프 누수

**파일:** `Services/MailProcessor.cs` → `AnalyzeInboxAsync` Phase 3

Phase 1과 동일한 패턴으로 Phase 3에서도 Session을 루프마다 획득.

**수정:** Phase 1과 동일 패턴으로 `p3Session`을 루프 외부로 호이스트.

---

### 12. SummaryControl.OnMoveFolderClick — NameSpace 누수

**파일:** `TaskPane/SummaryControl.cs` → `OnMoveFolderClick`

`Application.Session.GetItemFromID()`로 메일 획득 후 Session을 해제하지 않음.

**수정:** `Outlook.NameSpace session = ...; try { ... } finally { ComHelper.ReleaseAll(mail, session); }`

---

### 13. SummaryControl.OnPriorityChanged — NameSpace 누수

**파일:** `TaskPane/SummaryControl.cs` → `OnPriorityChanged`

우선순위 수동 변경 시 Session 누수. 12번과 동일 패턴.

**수정:** 동일 패턴으로 Session을 finally에서 해제.

---

### 14. SummaryControl.RunAnalysisAsync — NameSpace 누수

**파일:** `TaskPane/SummaryControl.cs` → `RunAnalysisAsync`

`await` 경계 전후로 Session을 보유하여 COM 스레드 문제 및 누수 발생.

**수정:** Session과 mail 참조를 `await` 이전에 해제하도록 구조 변경.

---

### 15. ProcessRetryQueueAsync — NameSpace 루프 누수

**파일:** `ThisAddIn.cs` → `ProcessRetryQueueAsync`

재시도 큐 처리 루프에서 Session을 루프마다 획득.

**수정:** Session을 루프 외부에서 1회 획득하고 finally에서 해제.

---

## HttpClient 리소스 누수 수정 (2건)

### 16. CallClaudeApiAsync — HttpResponseMessage 미해제

**파일:** `Services/LlmService.cs` → `CallClaudeApiAsync`

`HttpResponseMessage response = await _httpClient.SendAsync(...)` 이후 `response.Dispose()`를 호출하지 않음. API 호출마다 응답 스트림 리소스가 누적.

**수정:** `using (HttpResponseMessage response = await _httpClient.SendAsync(...))` 블록으로 감싸 자동 해제.

---

### 17. CallOpenAiApiAsync — HttpResponseMessage 미해제

**파일:** `Services/LlmService.cs` → `CallOpenAiApiAsync`

16번과 동일한 패턴.

**수정:** 동일하게 `using` 블록으로 감싸 자동 해제.

---

## 무한 성장 방지 수정 (2건)

### 18. RetryQueue 무한 성장

**파일:** `Services/RetryQueue.cs` → `Enqueue`

실패 메일이 지속적으로 누적될 경우 큐가 무한정 커질 수 있었음.

**수정:** `MaxQueueSize = 1000` 상수 추가. Enqueue 시 크기 초과 여부를 확인하고, 초과 시 신규 항목을 추가하지 않으며 경고 로그 기록.

```csharp
if (_queue.Count >= MaxQueueSize)
{
    Logger.Warn("RetryQueue: queue full (1000), dropping entry " + entryId);
    return;
}
```

---

### 19. FeedbackStore 무한 성장

**파일:** `Services/FeedbackStore.cs`

수동 우선순위 변경 이력이 영구 누적되어 장기 운영 시 파일 크기 증가.

**수정:** `RetentionDays = 90` 상수 추가. 생성자에서 `PurgeOldItems()` 호출하여 90일 초과 레코드 자동 삭제.

```csharp
private void PurgeOldItems()
{
    DateTime cutoff = DateTime.UtcNow.AddDays(-RetentionDays);
    _items.RemoveAll(i => i.ChangedAt < cutoff);
}
```

---

## 설정 데이터 유실 수정 (2건)

### 20. SemaphoreSlim 설정 변경 미반영

**파일:** `Services/LlmService.cs` → `ReloadConfig`

설정에서 동시 요청 수(`ConcurrentRequests`)를 변경해도 기존 `SemaphoreSlim` 인스턴스가 유지되어 효과 없음.

**수정:** `ReloadConfig`에서 현재 카운트 비교 후 새 `SemaphoreSlim` 생성, 기존 인스턴스 Dispose.

---

### 21. LastBatchAnalyzedAt 설정 저장 시 초기화

**파일:** `Forms/SettingsForm.cs` → `BuildConfig`

설정 창에서 저장 버튼 클릭 시 새 `AppConfig` 객체를 생성하여 반환하는 과정에서 `ProcessingConfig.LastBatchAnalyzedAt`이 null로 초기화됨. 결과적으로 설정을 저장할 때마다 증분 배치 분석 베이스라인이 리셋되어 매번 전체 스캔이 실행됨.

**수정:** `BuildConfig` 내에서 원본 설정의 값을 신규 설정 객체에 복사.

```csharp
config.Processing.LastBatchAnalyzedAt = _original.Processing.LastBatchAnalyzedAt;
```

---

## 보안/정확성 수정 (1건)

### 22. SelfHealRegistration Office 버전 하드코딩

**파일:** `ThisAddIn.cs` → `SelfHealRegistration`

레지스트리 경로에 Office 버전이 `"16.0"`으로 하드코딩되어 있어 향후 Office 버전 업그레이드 시 자가 복구 기능이 동작하지 않을 수 있음.

**수정:** `Application.Version`에서 Major.Minor를 동적으로 추출하고, 파싱 실패 시 `"16.0"`으로 폴백.

```csharp
string officeVer = "16.0";
try
{
    string appVer = Application.Version ?? "";
    int dot2 = appVer.IndexOf('.', appVer.IndexOf('.') + 1);
    if (dot2 > 0) officeVer = appVer.Substring(0, dot2);
}
catch { /* fall back to 16.0 */ }
```

---

## 코드 중복 제거 (4건)

### 23. `IsAnalyzed` 패턴 통합 (5곳 → 1곳)

**파일:** `Services/MailProcessor.cs` → `IsAnalyzed(object)`

`is bool b && b` 패턴이 5곳에 산재해 있었고, 일부 Outlook 버전에서 `olYesNo` 값이 `int`로 저장되어 `bool` 캐스팅 실패 가능성이 있었다.

**수정:** `internal static bool IsAnalyzed(object value)` 메서드로 통합. `bool`과 `int` 양쪽 모두 처리.

---

### 24. `GetSender` 패턴 통합 (3곳 → 1곳)

**파일:** `Services/MailProcessor.cs` → `GetSender(MailItem)`

`mail.SenderEmailAddress ?? mail.SenderName ?? ""` 인라인 코드가 3곳에 중복.

**수정:** `private static string GetSender(MailItem)` 메서드로 통합.

---

### 25. `SetViewState` 통합 (4곳 → 1곳)

**파일:** `TaskPane/SummaryControl.cs` → `SetViewState`

Task Pane의 상태 전환 시 10~15개 컨트롤의 `.Visible` 속성을 개별 설정하는 코드가 4곳에 걸쳐 ~50줄 반복.

**수정:** `SetViewState(showResult, showNotAnalyzed, showAnalyzing)` 단일 메서드로 통합. 각 호출부는 1줄로 단순화.

---

### 26. `ClearAnalysisFlag` 중앙화 (2곳 → 1곳)

**파일:** `Services/MailProcessor.cs` → `ClearAnalysisFlag(MailItem)`

SummaryControl과 MailProcessor에 각각 동일한 플래그 초기화 로직이 존재.

**수정:** `MailProcessor.ClearAnalysisFlag`를 `public static`으로 통합. SummaryControl에서 호출.

---

## 성능 개선 (5건)

### 27. DASL 필터 기반 미분석 메일 수집

**파일:** `Services/MailProcessor.cs` → `CollectUnanalyzedEntryIds`

**이전:** 받은편지함의 모든 메일을 순회하며 각각 `UserProperties.Find()`로 분석 여부 확인 — O(N) COM 호출.

**이후:** DASL 필터(`@SQL="...LLM_Analyzed" = 1`)로 분석된 메일 ID를 `HashSet`으로 수집 → 전체 메일에서 제외. DASL 실패 시 기존 방식으로 폴백.

---

### 28. 배치 크기와 동시성 제한 분리

**파일:** `Services/MailProcessor.cs` → `AnalyzeInboxAsync`

**이전:** 배치 크기 = `ConcurrentRequests` (기본 3). 3건 처리 후 다음 배치 시작 — LLM 파이프라인 유휴 시간 발생.

**이후:** 배치 크기 = `Max(ConcurrentRequests * 5, 15)`. SemaphoreSlim이 동시 요청을 제한하므로, 큰 배치에서 하나의 호출 완료 즉시 다음 호출 시작.

---

### 29. NewMailEx 병렬 처리

**파일:** `ThisAddIn.cs` → `Application_NewMailEx`

**이전:** 복수 메일 수신 시 순차 처리 (`foreach` + `await`).

**이후:** `Task.WhenAll(tasks)`로 병렬 처리. `SemaphoreSlim`이 동시성 제한.

---

### 30. SelectionChange 디바운스

**파일:** `ThisAddIn.cs`

**이전:** 키보드 화살표로 빠르게 메일 탐색 시 매번 COM 작업 실행 — UI 지연 및 CPU 낭비.

**이후:** 200ms `System.Windows.Forms.Timer` 디바운스 + 동일 EntryID 재선택 방지 (`_lastSelectedEntryId`).

---

### 31. 이중 Save 제거

**파일:** `Services/MailProcessor.cs`

**이전:** `TagSubject()` 내부에서 `mail.Save()` 호출 → `SaveAnalysisToMail()`에서 다시 `mail.Save()` 호출 — 불필요한 이중 디스크 I/O.

**이후:** `TagSubject()`에서 Save 제거. 호출 순서를 TagSubject → SaveAnalysisToMail로 변경하여 단일 Save로 통합.

---

## 기타 품질 개선 (3건)

### 32. BatchProgress.Processed 계산 속성화

**파일:** `Services/MailProcessor.cs`

**이전:** `Processed` 필드를 수동으로 증가 — 다른 카운터와 불일치 가능.

**이후:** `get { return Urgent + High + Normal + Low + Failed; }` 계산 속성으로 변경. 항상 정확.

---

### 33. 중복 XML 주석 제거

**파일:** `Services/MailProcessor.cs:367`

`IsAnalyzed`의 `/// <summary>` 주석이 `ClearAnalysisFlag` 메서드 위에 오배치되어 두 개의 `<summary>` 태그가 연속.

**수정:** 오배치된 주석 라인 삭제.

---

### 34. 미사용 using 제거

**파일:** `Ribbon/MailRibbon.cs:2`

`using System.IO;`가 선언되어 있으나 파일 내에서 사용되지 않음 (`SaveFileDialog`는 `System.Windows.Forms`).

**수정:** 해당 줄 삭제.
