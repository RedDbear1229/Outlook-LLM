# MailPrioritizer

Outlook 2016 VSTO Add-in으로, LLM(Claude / OpenAI)을 활용하여 수신 메일을 자동 요약하고 우선순위(긴급/높음/보통/낮음)별 폴더로 분류합니다.

---

## 목차

- [요구사항](#요구사항)
- [빌드](#빌드)
- [설치](#설치)
- [초기 설정](#초기-설정)
- [사용 방법](#사용-방법)
- [키보드 단축키](#키보드-단축키)
- [Task Pane 색상 테마](#task-pane-색상-테마)
- [설정 상세](#설정-상세)
- [폴더 구조](#폴더-구조)
- [데이터 저장 위치](#데이터-저장-위치)
- [로그 확인](#로그-확인)
- [문제 해결](#문제-해결)
- [제거](#제거)
- [프로젝트 구조](#프로젝트-구조)

---

## 요구사항

### 실행 환경

| 항목 | 요구사항 |
|------|---------|
| OS | Windows 10 이상 |
| Outlook | Microsoft Outlook 2016 데스크톱 버전 |
| .NET | .NET Framework 4.7.2 런타임 |
| VSTO Runtime | Microsoft Visual Studio Tools for Office Runtime |

> **VSTO Runtime** 이 없으면 Add-in 설치 자체가 불가능합니다. [Microsoft 공식 다운로드](https://www.microsoft.com/ko-kr/download/details.aspx?id=56961)에서 받을 수 있습니다.

### 빌드 환경

| 항목 | 요구사항 |
|------|---------|
| IDE | Visual Studio 2022 (Community 이상) |
| 워크로드 | **Office/SharePoint 개발** (Visual Studio Installer → 워크로드 탭) |
| SDK | .NET Framework 4.7.2 SDK |

### LLM API

다음 중 하나의 API 키가 필요합니다:

| API | 엔드포인트 | 발급처 |
|-----|-----------|--------|
| **Anthropic Claude** | `https://api.anthropic.com/v1` | console.anthropic.com |
| **OpenAI** | `https://api.openai.com/v1` | platform.openai.com |
| **호환 API** (Ollama 등) | 직접 입력 | — |

---

## 빌드

### 1. 저장소 클론

```bash
git clone <repository-url>
cd Outlook-LLM
```

### 2. NuGet 패키지 복원

Visual Studio에서 솔루션을 열면 자동 복원됩니다. 수동 복원:

```powershell
.\nuget.exe restore MailPrioritizer\MailPrioritizer.sln
```

의존 패키지:
- `Newtonsoft.Json 13.0.3`
- `System.Data.SQLite 1.x` (R-05 로컬 인덱스)

### 3. Visual Studio에서 빌드

```
MailPrioritizer\MailPrioritizer.sln 열기 → 빌드(B) → 솔루션 빌드(B)
```

또는 **F5** 로 Outlook과 함께 디버그 실행 (개발용).

### 4. CLI 빌드 (PowerShell)

```powershell
# Debug 빌드
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "C:\project\Outlook-LLM\MailPrioritizer\MailPrioritizer\MailPrioritizer.csproj" `
  /p:Configuration=Debug /v:minimal

# Release 빌드
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "C:\project\Outlook-LLM\MailPrioritizer\MailPrioritizer\MailPrioritizer.csproj" `
  /p:Configuration=Release /v:minimal

# Clean
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "C:\project\Outlook-LLM\MailPrioritizer\MailPrioritizer\MailPrioritizer.csproj" /t:Clean
```

> CLI 빌드는 VSTO 매니페스트/등록 파이프라인을 건너뜁니다(`IsVstoPublish=false`). 생성된 DLL은 기능상 완전하며 VS IDE의 F5 실행에서는 전체 파이프라인이 정상 동작합니다.

빌드 결과물: `MailPrioritizer\MailPrioritizer\bin\Debug\MailPrioritizer.dll`

---

## 설치

### 방법 1: 디버그 실행 (개발용)

Visual Studio에서 **F5**를 누르면 Outlook이 Add-in이 연결된 상태로 실행됩니다. Outlook을 종료하면 Add-in도 자동 해제됩니다.

### 방법 2: 원클릭 배포 (운영용)

프로젝트 루트에서 PowerShell 스크립트를 실행합니다:

```powershell
.\deploy-clickonce.ps1                             # 기본 빌드 (버전 1.0.0.0)
.\deploy-clickonce.ps1 -Open                       # 빌드 후 publish 폴더 열기
.\deploy-clickonce.ps1 -PublishVersion 1.0.0.5 -Open  # 버전 지정
```

`publish\` 폴더가 생성됩니다. **폴더 전체**를 대상 PC에 복사한 후:

#### 설치 방법 (권장: GUI 설치 관리자)

```
publish\MailPrioritizer.Installer.exe 실행 → [설치] 클릭
```

#### 설치 방법 (레거시: 배치 파일)

```
publish\install.bat 실행  (또는 MailPrioritizer.vsto 더블클릭)
→ 보안 경고 대화상자에서 [설치] 클릭
```

> **주의:** 서명 없는 내부 배포입니다. VSTO Runtime과 .NET Framework 4.7.2가 대상 PC에 설치되어 있어야 합니다.

### 설치 확인

Outlook 실행 후 상단 리본에 **"메일 분석"** 탭이 표시되면 설치 성공입니다.

### 자가 복구 (Self-Healing)

Add-in이 Outlook 시작 시 비활성화된 경우(VSTO 오류, 레지스트리 손상 등), 설치 관리자의 **복구** 기능으로 재등록할 수 있습니다:

```
MailPrioritizer.Installer.exe → [복구] 클릭
```

또는 Add-in 자체가 자동으로 레지스트리 등록을 감지하고 복구를 시도합니다 (Outlook 시작 시).

---

## 초기 설정

Add-in 설치 후 최초 1회 API 설정이 필요합니다.

### 1. 설정 창 열기

리본 → **메일 분석** 탭 → **환경설정** 클릭

### 2. API 설정 (탭 1)

| 항목 | 설명 | 예시 |
|------|------|------|
| **Endpoint URL** | LLM API 엔드포인트 | `https://api.anthropic.com/v1` |
| **API Token** | 발급받은 API 키 | `sk-ant-...` 또는 `sk-...` |
| **모델 이름** | 사용할 모델 ID | `claude-sonnet-4-20250514` |

> API 형식은 Endpoint URL로 **자동 감지**됩니다:
> - URL에 `anthropic`이 포함되면 → Claude Messages API (`x-api-key` 헤더)
> - 그 외 → OpenAI Chat Completions API (`Authorization: Bearer`)

### 3. 연결 테스트

**연결 테스트** 버튼을 클릭하여 API 연결을 확인합니다.

| 결과 | 표시 |
|------|------|
| 성공 | "연결 성공! 모델: claude-sonnet-4-20250514" |
| 실패 | 오류 코드 및 메시지 |

### 4. 저장

**저장** 버튼을 클릭합니다. API 토큰은 Windows DPAPI로 암호화되어 로컬 `config.json`에 저장됩니다.

---

## 사용 방법

### 선택 메일 분석

1. 받은편지함에서 메일을 선택합니다
2. 리본 → **선택 메일 요약** 클릭 (또는 우측 패널의 **지금 분석하기**)
3. 우측 Task Pane에 분석 결과가 표시됩니다:
   - **우선순위 배지** (색상 + 한글 라벨)
   - **2~3문장 한국어 요약**
   - **판단 근거** (1문장)
   - **분석 시각 및 모델 이름**

### 받은편지함 전체 분류

1. 리본 → **받은편지함 전체 분류** 클릭
2. 진행률 창이 표시됩니다 (실시간 카운트, 취소 가능)
3. 분석 완료 후 결과 요약 대화상자 표시:
   ```
   분석 완료!
   긴급: 2건 / 높음: 5건 / 보통: 12건 / 낮음: 8건 / 실패: 0건
   ```
4. 메일이 우선순위 폴더로 자동 이동됩니다

> **증분 분석:** 마지막 분석 시점 이후 수신된 메일만 대상으로 하므로, 두 번째 실행부터는 훨씬 빠릅니다.

### 우측 패널 (Task Pane)

메일을 선택하면 우측 패널에 분석 결과가 자동으로 표시됩니다 (200ms 디바운스 적용).

| 컨트롤 | 기능 |
|--------|------|
| 우선순위 드롭다운 | 우선순위를 수동 변경 (변경 즉시 폴더 이동) |
| **지금 분석하기** | 미분석 메일을 LLM으로 분석 |
| **재분석** | 이미 분석된 메일을 LLM으로 다시 분석 |
| **폴더이동** | 현재 우선순위에 해당하는 폴더로 이동 |
| **통계 새로고침** | 받은편지함 전체/분석됨/우선순위별 건수 표시 |
| API 상태 표시 | Task Pane 하단 — 마지막 API 호출 성공/실패 상태 |

### 전체 재분석

리본 → **전체 재분석**: 모든 기존 분석 결과를 초기화하고 처음부터 다시 분석합니다.

> 모델이나 프롬프트를 변경한 후, 또는 발신자 규칙을 대폭 수정한 후 사용하세요.

### 결과 내보내기

리본 → **결과 내보내기**: 분석된 메일의 결과를 CSV 파일로 저장합니다.

- **컬럼:** Subject, Sender, Priority, Summary, PriorityReason
- **인코딩:** UTF-8
- **파일명 기본값:** `MailPrioritizer_Export_yyyyMMdd.csv`
- **특수문자 처리:** 쉼표, 큰따옴표, 줄바꿈 자동 이스케이프

---

## 키보드 단축키

> Outlook이 포커스를 가진 상태에서 전역으로 동작합니다.

| 단축키 | 기능 |
|--------|------|
| `Ctrl + Shift + M` | 현재 선택된 메일 분석 |
| `Ctrl + Shift + 1` | 선택된 메일 우선순위 → **긴급** |
| `Ctrl + Shift + 2` | 선택된 메일 우선순위 → **높음** |
| `Ctrl + Shift + 3` | 선택된 메일 우선순위 → **보통** |
| `Ctrl + Shift + 4` | 선택된 메일 우선순위 → **낮음** |

> 다른 프로그램이 동일 단축키를 이미 등록한 경우 로그에 경고가 기록되며 해당 단축키는 비활성화됩니다.

---

## Task Pane 색상 테마

환경설정 → 분류 설정 탭 → **테마** 드롭다운에서 3가지 테마를 선택할 수 있습니다.

| 테마 | 설명 | 적합한 Outlook 모드 |
|------|------|-------------------|
| **밝은 테마** (기본) | 흰 배경, 밝은 배지 색상 | Outlook White 모드 |
| **회색 테마** | 중간 회색 배경, 진한 배지 | Outlook Dark Grey 모드 |
| **어두운 테마** | 어두운 배경, 밝은 텍스트 | Outlook Black 모드 |

설정 저장 즉시 Task Pane에 반영됩니다.

---

## 설정 상세

리본 → **환경설정**에서 변경할 수 있습니다.

### 탭 1 — API 설정

| 항목 | 기본값 | 설명 |
|------|--------|------|
| Endpoint URL | `https://api.anthropic.com/v1` | LLM API 주소 |
| API Token | (비어있음) | API 인증 키 — DPAPI로 암호화 저장 |
| 모델 이름 | `claude-sonnet-4-20250514` | 호출할 모델 ID |
| 최대 토큰 | `1024` | 응답 최대 토큰 수 |
| 타임아웃 | `30`초 | API 응답 대기 시간 |

### 탭 2 — 프롬프트

시스템 프롬프트를 직접 편집할 수 있습니다.

> **반드시 유지해야 하는 JSON 형식:**
> ```json
> {"summary":"...","priority":"urgent|high|normal|low","priority_reason":"..."}
> ```

**기본값으로 복원** 버튼으로 언제든 내장 프롬프트로 초기화할 수 있습니다.

### 탭 3 — 분류 설정

| 항목 | 기본값 | 설명 |
|------|--------|------|
| 폴더 접두사 | `우선순위` | 받은편지함 하위 폴더 이름 |
| 자동 폴더 이동 | 켜짐 | 분석 후 해당 우선순위 폴더로 자동 이동 |
| 제목 태그 삽입 | 꺼짐 | 제목 앞에 `[긴급]`, `[높음]` 등의 태그 추가 |
| 본문 최대 길이 | `4000`자 | LLM에 전송하는 본문 최대 글자 수 (500~20000) |
| 첨부파일명 포함 | 켜짐 | 첨부파일명을 LLM 프롬프트에 추가 |
| 대상 계정 | (기본 저장소) | 분석 대상 Outlook 계정 선택 |
| 새 메일 자동 분석 | 꺼짐 | 메일 수신 즉시 자동 분석 |
| 동시 요청 수 | `3` | 병렬 LLM 호출 최대 수 |
| 색상 테마 | `밝은 테마` | Task Pane 색상 테마 |

### 탭 4 — 발신자 규칙

발신자 이메일 주소 또는 도메인에 따라 LLM 호출 없이 즉시 우선순위를 결정하는 규칙입니다.

| 컬럼 | 설명 | 예시 |
|------|------|------|
| 유형 | `email` (정확 일치) 또는 `domain` (도메인 일치) | `domain` |
| 패턴 | 매칭할 이메일 또는 도메인 | `newsletter.com` |
| 우선순위 | 적용할 우선순위 | `낮음` |
| 활성 | 규칙 활성화 여부 | ✓ |
| 메모 | 관리용 설명 (LLM에 전달 안 됨) | `뉴스레터 자동 분류` |

규칙은 LLM 호출 전에 순서대로 평가되며, 일치하는 규칙이 있으면 LLM을 호출하지 않습니다.

---

## 폴더 구조

분석이 완료되면 받은편지함 아래에 다음 폴더가 자동 생성됩니다:

```
받은편지함/
  └── 우선순위/
      ├── 긴급/     ← Urgent (빨간색 배지)
      ├── 높음/     ← High   (주황색 배지)
      ├── 보통/     ← Normal (초록색 배지)
      └── 낮음/     ← Low    (회색 배지)
```

- 폴더 접두사와 하위 폴더명은 **탭 3 — 분류 설정**에서 변경 가능
- 이모지 폴더명 미사용 (Exchange/IMAP 서버 호환성)
- 다중 계정 환경에서는 **탭 3 — 대상 계정**에서 원하는 계정 선택

---

## 데이터 저장 위치

| 파일 | 경로 | 설명 |
|------|------|------|
| 설정 | `%AppData%\MailPrioritizer\config.json` | API 키(DPAPI 암호화), 전체 설정 |
| 로그 | `%AppData%\MailPrioritizer\Logs\MailPrioritizer-yyyy-MM-dd.log` | 진단 로그 (7일 보관) |
| 분석 인덱스 | `%AppData%\MailPrioritizer\index.db` | SQLite — 분석 결과 로컬 캐시 |
| 재시도 큐 | `%AppData%\MailPrioritizer\retry_queue.json` | LLM 호출 실패 메일 EntryID |
| 피드백 | `%AppData%\MailPrioritizer\feedback.json` | 수동 우선순위 변경 이력 (90일 보관) |

---

## 로그 확인

문제 발생 시 로그 파일을 확인하세요:

```
%AppData%\MailPrioritizer\Logs\MailPrioritizer-yyyy-MM-dd.log
```

Windows 탐색기 주소창에 `%AppData%\MailPrioritizer\Logs`를 입력하면 바로 이동합니다.

**로그 포맷:**

```
[2026-04-12 14:30:45.123] [INFO]  === MailPrioritizer started ===
[2026-04-12 14:30:52.456] [INFO]  AnalyzeMailAsync: start, subject="서버 장애 보고..."
[2026-04-12 14:30:55.789] [ERROR] AnalyzeMailAsync: API call failed
  Exception: System.Net.Http.HttpRequestException: ...
  StackTrace: at ...
  InnerException: System.Net.WebException: ...
```

**로그 레벨:**

| 레벨 | 용도 |
|------|------|
| `DEBUG` | 상세 내부 동작 (기본 비활성) |
| `INFO` | 주요 이벤트 (시작, 분석 완료, 설정 저장 등) |
| `WARN` | 비치명적 이상 상황 (단축키 충돌, 폴더 미발견 등) |
| `ERROR` | 오류 발생 (API 실패, COM 예외 등) — 스택 트레이스 포함 |

- 7일이 지난 로그 파일은 자동 삭제됩니다
- API 토큰과 메일 본문은 로그에 절대 기록되지 않습니다

---

## 문제 해결

### "API 설정이 필요합니다" 메시지

Endpoint URL, API Token, 모델 이름이 모두 입력되어야 합니다. 환경설정 탭 1에서 확인하세요.

### 연결 테스트 실패

| 증상 | 원인 | 해결 |
|------|------|------|
| HTTP 401 Unauthorized | API 토큰 오류 | 토큰을 다시 확인·입력 |
| HTTP 404 Not Found | Endpoint URL 오류 | URL 끝에 `/v1`이 포함되어 있는지 확인 |
| HTTP 429 Too Many Requests | API 한도 초과 | 잠시 후 재시도, 또는 API 플랜 확인 |
| 연결 시간 초과 | 네트워크 차단 | 방화벽/프록시 설정 확인 |
| Unexpected token | 잘못된 엔드포인트 | URL 형식 확인 (예: `/v1/messages` → `/v1`만 입력) |

### Task Pane API 상태 표시가 빨간색

- 마지막 API 호출이 실패한 상태입니다
- 로그에서 `[ERROR]` 항목을 확인하세요
- 환경설정 → 연결 테스트로 API 상태를 재확인하세요

### 분석 결과가 "보통"으로만 나옴

- LLM 응답이 JSON 형식이 아닌 경우 기본값(Normal)으로 폴백됩니다
- 환경설정 → 탭 2(프롬프트)에서 JSON 응답 형식 지시가 유지되어 있는지 확인하세요
- 로그에서 `ParseLlmResponse: JSON parse failed` 메시지를 확인하세요

### 폴더가 생성되지 않음

- 환경설정 → 탭 3에서 "자동 폴더 이동"이 켜져 있는지 확인
- 일부 Exchange 서버는 한글 폴더명을 거부합니다. 폴더 접두사를 영문으로 변경해 보세요

### 키보드 단축키가 동작하지 않음

- 다른 프로그램이 동일한 단축키를 이미 등록했을 수 있습니다
- 로그에서 `RegisterHotKey failed` 메시지를 확인하세요
- Outlook에 포커스가 있는지 확인하세요

### Outlook이 느려짐

- "새 메일 자동 분석"이 켜져 있으면 대량 수신 시 부하가 발생할 수 있습니다. 끄고 수동 분석을 사용하세요
- 환경설정 → 탭 3에서 **동시 요청 수**를 줄이면 부하가 감소합니다 (권장: 1~2)
- 배치 분석 중에는 다른 Outlook 작업 응답이 느릴 수 있습니다 (정상 동작)

### Task Pane이 보이지 않음

Outlook을 재시작하면 대부분 해결됩니다. 그래도 안 보이면:

1. 리본의 **보기** 탭 → **작업창** 클릭
2. Add-in이 비활성화된 경우: 리본 클릭 시 "Add-in이 비활성화 되었습니다" 메시지 → **설치 관리자** 실행 → **복구**

### "메일 분석" 탭이 리본에 없음

Add-in이 비활성화된 것입니다. 다음 중 하나를 시도하세요:

1. **설치 관리자 복구:** `publish\MailPrioritizer.Installer.exe` → [복구]
2. **Outlook 옵션 확인:** 파일 → 옵션 → 추가 기능 → COM 추가 기능 → MailPrioritizer 활성화
3. **재설치:** 제거 후 재설치

---

## 제거

### GUI 설치 관리자로 제거

```
MailPrioritizer.Installer.exe → [제거] 클릭
```

### Windows 설정에서 제거

1. Windows 설정 → 앱 → 앱 및 기능
2. "MailPrioritizer" 검색 → 제거

### 설정 파일 수동 삭제 (선택)

제거 후에도 설정, 로그, 분석 인덱스 파일은 남아 있습니다. 완전 삭제:

```cmd
rmdir /s /q "%AppData%\MailPrioritizer"
```

> 메일에 저장된 UserProperty(`LLM_Priority`, `LLM_Summary`, `LLM_PriorityReason`, `LLM_Analyzed`, `LLM_ModelName`, `LLM_AnalyzedAt`)는 각 MailItem에 남아 있으며 Outlook 동작에 영향을 주지 않습니다.

---

## 프로젝트 구조

```
Outlook-LLM/
├── README.md
├── CLAUDE.md                          # 개발 가이드 (아키텍처, 제약사항, 설계 결정)
├── deploy-clickonce.ps1               # Release 빌드 + 배포 패키징 스크립트
├── docs/
│   ├── FEATURES.md                    # 구현 완료 기능 목록
│   ├── BUGFIXES.md                    # 오류 수정 및 코드 개선 이력
│   └── ROADMAP.md                     # 향후 구현 예정 아이템
└── MailPrioritizer/
    ├── MailPrioritizer.sln
    ├── MailPrioritizer.CLI.targets    # CLI 빌드용 no-op 타겟
    ├── MailPrioritizer.Publish.targets # 서명 없는 매니페스트 파이프라인
    ├── packages/                      # NuGet 패키지 캐시
    ├── MailPrioritizer.Installer/     # C# GUI 설치 관리자 (설치/복구/제거)
    └── MailPrioritizer/
        ├── MailPrioritizer.csproj     # VSTO 프로젝트 (C# 7.3, .NET 4.7.2)
        ├── packages.config            # NuGet: Newtonsoft.Json, System.Data.SQLite
        ├── ThisAddIn.cs               # VSTO 진입점, 서비스 소유, 이벤트 배선
        ├── Config/
        │   └── ConfigManager.cs       # JSON 설정 로드/저장, DPAPI 토큰 암호화
        ├── Models/
        │   ├── Priority.cs            # 우선순위 enum + 확장 메서드 (ToKorean 등)
        │   ├── MailAnalysis.cs        # LLM 분석 결과 DTO
        │   └── AppConfig.cs           # 설정 모델 (5개 중첩 클래스)
        ├── Services/
        │   ├── LlmService.cs          # Claude/OpenAI 듀얼 API 클라이언트, 재시도
        │   ├── MailProcessor.cs       # 메일 분석 엔진 (단건/배치/통계/내보내기)
        │   ├── FolderManager.cs       # 우선순위 폴더 생성/이동, 다중 계정 지원
        │   ├── RetryQueue.cs          # LLM 실패 메일 자동 재시도 큐 (최대 1000건)
        │   ├── FeedbackStore.cs       # 사용자 수동 변경 이력/정확도 통계 (90일 보관)
        │   └── IndexDatabase.cs       # SQLite 분석 결과 로컬 인덱스
        ├── Ribbon/
        │   ├── MailRibbon.cs          # 리본 버튼 이벤트 핸들러
        │   └── MailRibbon.xml         # 리본 UI 정의 (5개 버튼)
        ├── Forms/
        │   ├── SettingsForm.cs        # 4탭 환경설정 다이얼로그
        │   └── ProgressForm.cs        # 배치 분석 진행률 폼
        ├── TaskPane/
        │   ├── SummaryControl.cs      # 우측 분석 결과 패널
        │   └── ThemePalette.cs        # Task Pane 색상 팔레트 (Light / Grey / Dark)
        └── Utils/
            ├── ComHelper.cs           # COM 객체 안전 해제
            ├── Logger.cs              # 파일 기반 로거 (7일 보관, 스레드 안전)
            ├── KeyboardShortcutManager.cs  # 전역 키보드 단축키 (Win32 RegisterHotKey)
            └── MailPropertyNames.cs   # UserProperty 이름 상수
```

### 핵심 데이터 흐름

```
리본 클릭 / NewMailEx / Task Pane 버튼
  ↓
[STA] ThisAddIn / MailRibbon
  ↓
[STA] MailProcessor.Phase1 — COM에서 메일 데이터 추출 → MailDataItem POCO
  ↓
[ThreadPool] MailProcessor.Phase2 — LlmService 병렬 API 호출 (SemaphoreSlim 제한)
  ↓
[STA] MailProcessor.Phase3 — UserProperty 저장 + 제목 태그 + 폴더 이동 + IndexDatabase 기록
  ↓
Task Pane 갱신
```

### 서비스 생명주기

모든 서비스는 `ThisAddIn_Startup`에서 생성되고 `ThisAddIn`이 소유합니다. 설정 변경은 `ApplyConfigChange()` → `ReloadConfig()` 체인으로 전파됩니다 (HttpClient, SemaphoreSlim 재생성 포함).

---

## 라이선스

이 프로젝트의 라이선스는 별도로 지정되지 않았습니다.
