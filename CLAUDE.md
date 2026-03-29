# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**MailPrioritizer** — Outlook 2016 VSTO Add-in that summarizes emails via LLM and auto-classifies them into priority folders (Urgent/High/Normal/Low).

Target: Windows 10 + Outlook 2016, C# 7.3 / .NET Framework 4.7.2.

## Build & Run

This is a Visual Studio VSTO project. It cannot be built in this Termux environment — source code generation only.

```
# On Windows with VS 2019/2022 + Office Developer Tools:
# Open MailPrioritizer.sln → Build → F5 (launches Outlook with add-in attached)
```

NuGet dependencies: `Newtonsoft.Json 13.0.3`

## Architecture

```
ThisAddIn (VSTO entry point, service owner, event wiring)
  ├── Ribbon/MailRibbon          — 5 buttons: Settings, Analyze Selected, Analyze All, Reanalyze All, Export
  ├── Forms/SettingsForm         — 3-tab dialog: API Settings / Prompt / Classification
  ├── Forms/ProgressForm         — Batch analysis progress bar with CancellationToken
  ├── TaskPane/SummaryControl    — Right-docked Task Pane, updates on SelectionChange (200ms debounce)
  ├── Services/
  │   ├── LlmService            — Dual API client (Claude Messages + OpenAI Chat Completions)
  │   ├── MailProcessor          — Core engine: extract → LLM call → save UserProperty → move folder
  │   └── FolderManager          — Create/find priority subfolders, multi-store support
  ├── Models/                    — Priority enum, MailAnalysis, AppConfig (4 nested config classes)
  ├── Config/ConfigManager       — JSON config at %AppData%/MailPrioritizer/, DPAPI token encryption
  └── Utils/
      ├── ComHelper              — Marshal.ReleaseComObject wrapper
      ├── Logger                 — File-based logger, 7-day rotation, thread-safe
      └── MailPropertyNames      — UserProperty name constants (LLM_Priority, LLM_Summary, etc.)
```

**Data flow:** Ribbon click / NewMailEx / Task Pane button → MailProcessor reads `MailItem.Body` → LlmService calls API → JSON parsed to `MailAnalysis` → saved as `UserProperties` on mail → optionally tag subject → mail moved to priority folder → Task Pane updated.

**Service lifecycle:** All services instantiated in `ThisAddIn_Startup`, owned by `ThisAddIn`. Settings changes propagate via `ApplyConfigChange()` which calls `ReloadConfig()` on each service (including HttpClient and SemaphoreSlim recreation in LlmService).

## Features

| Feature | Entry Point | Notes |
|---------|------------|-------|
| Single mail analysis | Ribbon button / Task Pane "지금 분석하기" | Shows result in Task Pane |
| Batch inbox analysis | Ribbon "받은편지함 전체 분류" | ProgressForm with cancel, consecutive-failure guard (5+) |
| Reanalyze all | Ribbon "전체 재분석" | Resets LLM_Analyzed flags, then runs batch |
| Auto-analyze new mail | NewMailEx event | Enabled via Settings, parallel via Task.WhenAll |
| Manual priority change | Task Pane ComboBox | Updates UserProperty + optional auto-move |
| Export to CSV | Ribbon "결과 내보내기" | Subject, Sender, Priority, Summary, Reason |
| Inbox statistics | Task Pane "통계 새로고침" | Total/Analyzed/per-priority counts |
| Multi-store support | Settings "대상 계정" | Falls back to default store on error |

## Batch Processing Pipeline

Batch analysis decouples COM reads from LLM calls for efficiency:

1. **Collect phase:** DASL filter (`@SQL="...LLM_Analyzed" = 1`) finds analyzed mails → HashSet of IDs → iterate all items excluding those IDs. Falls back to full UserProperty scan if DASL fails.
2. **Phase 1 (STA):** Extract mail data (subject, body, sender) into `MailDataItem` POCOs. Batch size = `Max(ConcurrentRequests * 5, 15)`.
3. **Phase 2 (ThreadPool):** Parallel LLM calls via `Task.WhenAll()`, throttled by `SemaphoreSlim`.
4. **Phase 3 (STA):** Sequential COM writes — TagSubject (no Save) → SaveAnalysisToMail (single Save) → MoveToFolder.

## Critical Constraints

### C# 7.3 Only
.NET Framework 4.7.2 defaults to C# 7.3. Do NOT use: switch expressions, nullable reference types, default interface members, `using` declarations, ranges/indices, or other C# 8.0+ features.

### COM Object Cleanup
Every Outlook COM object (MailItem, MAPIFolder, UserProperties, UserProperty, Folders, Items, Selection, NameSpace, Store, Stores) must be released via `ComHelper.ReleaseAll()` in a `finally` block. Failure causes memory leaks and Outlook hanging on exit. Pattern:

```csharp
Outlook.UserProperties props = null;
Outlook.UserProperty prop = null;
try
{
    props = mail.UserProperties;
    prop = props.Find(MailPropertyNames.Priority);
    // use prop...
}
finally
{
    ComHelper.ReleaseAll(prop, props);
}
```

### STA Thread Marshaling
Outlook runs on STA thread. All COM access must happen on this thread. After `await` (HttpClient calls), continuation must return to STA thread — use `ConfigureAwait(true)` (the default). `WindowsFormsSynchronizationContext` is installed in `ThisAddIn_Startup` to guarantee this. Never use `ConfigureAwait(false)` before COM access.

### UserProperty Find-then-Add
Always `props.Find("name")` first, only `props.Add("name", type)` if Find returns null. Calling Add on existing property creates duplicates or throws. Use constants from `MailPropertyNames`.

### olYesNo Type Ambiguity
Outlook stores `olYesNo` UserProperty values as `bool` on some versions and `int` on others. Always use `MailProcessor.IsAnalyzed(object value)` which handles both types.

## LLM API Dual Support

API format auto-detected by endpoint URL:
- URL contains "anthropic" → Claude Messages API (`/messages`, `x-api-key` header, `anthropic-version: 2023-06-01`)
- Otherwise → OpenAI Chat Completions (`/chat/completions`, `Bearer` auth)

Retry: exponential backoff (2s/4s/8s) on 429/500/502/503, max 3 retries.

System prompt is user-configurable (Settings > Prompt tab). Single source of truth: `LlmService.DefaultSystemPrompt`. LLM must return JSON: `{"summary":"...","priority":"urgent|high|normal|low","priority_reason":"..."}`. Invalid JSON or ```` ```json ``` ```` fences are handled by `ExtractJson()`. Parse failure → fallback to Normal priority.

## Design Decisions

- **Custom Task Pane over Form Region:** Adjoining Form Regions do NOT render in Outlook 2016's Reading Pane, only in Inspector windows. Task Pane works everywhere via `Explorer.SelectionChange`.
- **UserProperty over Rules:** Outlook Rules Wizard cannot use UserProperty fields as conditions. Instead, UserProperties work with View Filters, Search Folders, and custom columns.
- **MailItem.Body over HTMLBody:** `.Body` returns Outlook's built-in plaintext conversion. No HTML parsing library needed.
- **No emoji in default folder names:** Some Exchange/IMAP servers reject emoji in folder names.
- **EntryID-based batch processing:** MailItem COM references become invalid after `Move()`. Collect EntryIDs first, then `GetItemFromID()` for each phase.
- **SelectionChange debounce:** 200ms `System.Windows.Forms.Timer` + same-EntryID guard prevents excessive COM operations during rapid keyboard navigation.
- **TagSubject does not Save:** Caller performs single `mail.Save()` after all property writes to avoid double-save overhead.

## Config

Stored at `%AppData%/MailPrioritizer/config.json`. API token encrypted with Windows DPAPI (`ProtectedData.Protect/Unprotect`). Empty `systemPrompt` field means use built-in default. ConfigManager auto-creates directory and default config on first run.

Key config fields:
- `Llm`: Endpoint, ApiToken (DPAPI), ModelName, MaxTokens, TimeoutSeconds, SystemPrompt
- `Classification`: AutoMoveToFolder, FolderPrefix ("우선순위"), FolderNames (dictionary)
- `Display`: TagSubjectWithPriority
- `Processing`: MaxBodyLength (4000), ConcurrentRequests (3), TargetStoreId, AutoAnalyzeNewMail

## Logging

`Utils/Logger.cs` — static class, thread-safe via `lock` + `File.AppendAllText`.
- Path: `%AppData%/MailPrioritizer/Logs/MailPrioritizer-yyyy-MM-dd.log`
- Levels: Debug, Info, Warn, Error (default min: Info)
- Error level includes exception type, message, stack trace, inner exception
- 7-day auto-cleanup on `Initialize()`
- All Logger failures are silently caught — never crashes the Add-in
- **Never log:** API tokens, mail body content, DPAPI bytes
