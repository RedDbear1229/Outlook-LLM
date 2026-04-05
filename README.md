# MailPrioritizer

Outlook 2016 VSTO Add-in으로, LLM(Claude / OpenAI)을 활용하여 수신 메일을 자동 요약하고 우선순위(긴급/높음/보통/낮음)별 폴더로 분류합니다.

---

## 목차

- [요구사항](#요구사항)
- [빌드](#빌드)
- [설치](#설치)
- [초기 설정](#초기-설정)
- [사용 방법](#사용-방법)
- [설정 상세](#설정-상세)
- [폴더 구조](#폴더-구조)
- [로그 확인](#로그-확인)
- [문제 해결](#문제-해결)
- [제거](#제거)
- [프로젝트 구조](#프로젝트-구조)

---

## 요구사항

### 실행 환경
- Windows 10 이상
- Microsoft Outlook 2016 (데스크톱 버전)
- .NET Framework 4.7.2 런타임

### 빌드 환경
- Visual Studio 2019 또는 2022
- **Office/SharePoint 개발** 워크로드 (Visual Studio Installer에서 설치)
- .NET Framework 4.7.2 SDK

### LLM API
다음 중 하나의 API 키가 필요합니다:
- **Anthropic Claude API** — [console.anthropic.com](https://console.anthropic.com)에서 발급
- **OpenAI API** (또는 호환 API) — [platform.openai.com](https://platform.openai.com)에서 발급

---

## 빌드

### 1. 저장소 클론

```bash
git clone <repository-url>
cd MailPrioritizer
```

### 2. NuGet 패키지 복원

Visual Studio에서 솔루션을 열면 자동 복원됩니다. 수동 복원이 필요한 경우:

```
nuget restore MailPrioritizer.sln
```

의존 패키지:
- `Newtonsoft.Json 13.0.3`

### 3. 솔루션 빌드

```
Visual Studio → MailPrioritizer.sln 열기 → 빌드(B) → 솔루션 빌드(B)
```

또는 명령줄:

```cmd
msbuild MailPrioritizer.sln /p:Configuration=Release
```

빌드 결과물은 `MailPrioritizer/bin/Release/` 에 생성됩니다.

---

## 설치

### 방법 1: 디버그 실행 (개발용)

Visual Studio에서 `F5`를 누르면 Outlook이 Add-in이 연결된 상태로 실행됩니다. Outlook을 종료하면 Add-in도 자동 해제됩니다.

### 방법 2: 원클릭 배포 (운영용)

프로젝트 루트에서 PowerShell 스크립트를 실행합니다:

```powershell
.\deploy-clickonce.ps1           # 기본 빌드
.\deploy-clickonce.ps1 -Open     # 빌드 후 폴더 열기
.\deploy-clickonce.ps1 -PublishVersion 1.0.0.5 -Open  # 버전 지정
```

`publish\` 폴더가 생성됩니다. 대상 PC에 폴더 전체를 복사한 뒤:
- `install.bat` 실행 (또는 `MailPrioritizer.vsto` 더블클릭)
- 보안 경고 대화상자에서 [설치] 클릭

> 서명 없는 내부 배포입니다. VSTO Runtime이 대상 PC에 설치되어 있어야 합니다.

### 설치 확인

Outlook 실행 후 상단 리본에 **"메일 분석"** 탭이 표시되면 설치 성공입니다.

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

> API 형식은 Endpoint URL로 자동 감지됩니다:
> - URL에 `anthropic`이 포함되면 → Claude Messages API
> - 그 외 → OpenAI Chat Completions API

### 3. 연결 테스트

**연결 테스트** 버튼을 클릭하여 API 연결을 확인합니다.
- 성공: "연결 성공! 모델: claude-sonnet-4-20250514"
- 실패: 오류 메시지 확인 후 설정 수정

### 4. 저장

**저장** 버튼을 클릭합니다. API 토큰은 Windows DPAPI로 암호화되어 로컬에 저장됩니다.

---

## 사용 방법

### 선택 메일 분석

1. 받은편지함에서 메일을 선택합니다
2. 리본 → **선택 메일 요약** 클릭 (또는 우측 패널의 **지금 분석하기**)
3. 우측 Task Pane에 분석 결과가 표시됩니다:
   - 우선순위 배지 (색상 + 한글 라벨)
   - 2~3문장 한국어 요약
   - 판단 근거

### 받은편지함 전체 분류

1. 리본 → **받은편지함 전체 분류** 클릭
2. 진행률 창이 표시됩니다 (실시간 카운트, 취소 가능)
3. 분석 완료 후 결과 요약 대화상자 표시:
   ```
   분석 완료!
   긴급: 2건 / 높음: 5건 / 보통: 12건 / 낮음: 8건 / 실패: 0건
   ```
4. 메일이 우선순위 폴더로 자동 이동됩니다

### 우측 패널 (Task Pane)

메일을 선택하면 우측 패널에 분석 결과가 자동으로 표시됩니다.

| 버튼 | 기능 |
|------|------|
| 우선순위 드롭다운 | 우선순위를 수동 변경 (변경 즉시 폴더 이동) |
| **재분석** | 현재 메일을 LLM으로 다시 분석 |
| **폴더이동** | 현재 우선순위에 해당하는 폴더로 이동 |
| **통계 새로고침** | 받은편지함 전체/분석됨/우선순위별 건수 표시 |

### 전체 재분석

리본 → **전체 재분석**: 모든 기존 분석 결과를 초기화하고 처음부터 다시 분석합니다. 모델이나 프롬프트를 변경한 후 사용하세요.

### 결과 내보내기

리본 → **결과 내보내기**: 분석된 메일의 결과를 CSV 파일로 저장합니다.

출력 컬럼: `Subject, Sender, Priority, Summary, PriorityReason`

---

## 설정 상세

리본 → **환경설정**에서 변경할 수 있습니다.

### API 설정 (탭 1)

| 항목 | 기본값 | 설명 |
|------|--------|------|
| Endpoint URL | `https://api.anthropic.com/v1` | LLM API 주소 |
| API Token | (비어있음) | API 인증 키 (DPAPI 암호화 저장) |
| 모델 이름 | `claude-sonnet-4-20250514` | 호출할 모델 ID |

### 프롬프트 (탭 2)

시스템 프롬프트를 직접 편집할 수 있습니다. JSON 응답 형식 부분은 반드시 유지해야 합니다:

```json
{"summary":"...","priority":"urgent|high|normal|low","priority_reason":"..."}
```

**기본값으로 복원** 버튼으로 언제든 초기화할 수 있습니다.

### 분류 설정 (탭 3)

| 항목 | 기본값 | 설명 |
|------|--------|------|
| 폴더 접두사 | `우선순위` | 받은편지함 하위 폴더 이름 |
| 자동 폴더 이동 | 켜짐 | 분석 후 해당 우선순위 폴더로 자동 이동 |
| 제목 태그 삽입 | 꺼짐 | 제목 앞에 `[긴급]` 등의 태그 추가 |
| 본문 최대 길이 | 4000자 | LLM에 전송하는 본문 최대 글자 수 |
| 대상 계정 | (기본 저장소) | 분석 대상 Outlook 계정 선택 |
| 새 메일 자동 분석 | 꺼짐 | 메일 수신 즉시 자동 분석 |

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

폴더 접두사와 하위 폴더명은 설정에서 변경할 수 있습니다.

---

## 로그 확인

문제 발생 시 로그 파일을 확인하세요:

```
%AppData%\MailPrioritizer\Logs\MailPrioritizer-yyyy-MM-dd.log
```

Windows 탐색기 주소창에 `%AppData%\MailPrioritizer\Logs`를 입력하면 바로 이동합니다.

로그 형식:
```
[2026-03-29 14:30:45.123] [INFO] Startup complete
[2026-03-29 14:30:52.456] [INFO] AnalyzeMailAsync: start, subject="서버 장애 보고..."
[2026-03-29 14:30:55.789] [ERROR] AnalyzeMailAsync: API call failed
  Exception: System.Net.Http.HttpRequestException: ...
  StackTrace: ...
```

- 7일이 지난 로그 파일은 자동 삭제됩니다
- API 토큰과 메일 본문은 로그에 기록되지 않습니다

---

## 문제 해결

### "API 설정이 필요합니다" 메시지

Endpoint URL, API Token, 모델 이름이 모두 입력되어야 합니다. 환경설정에서 확인하세요.

### 연결 테스트 실패

| 증상 | 원인 | 해결 |
|------|------|------|
| HTTP 401 Unauthorized | API 토큰 오류 | 토큰을 다시 확인·입력 |
| HTTP 404 Not Found | Endpoint URL 오류 | URL 끝에 `/v1`이 포함되어 있는지 확인 |
| 연결 시간 초과 | 네트워크 차단 | 방화벽/프록시 설정 확인 |
| HTTP 429 Too Many Requests | API 한도 초과 | 잠시 후 재시도, 또는 API 플랜 확인 |

### 분석 결과가 "Normal"로만 나옴

- LLM 응답이 JSON 형식이 아닌 경우 기본값(Normal)으로 폴백됩니다
- 프롬프트 설정을 확인하고, JSON 응답 형식 지시가 포함되어 있는지 확인하세요
- 로그에서 `ParseLlmResponse: JSON parse failed` 메시지를 확인하세요

### 폴더가 생성되지 않음

- 설정 > 분류 설정에서 "자동 폴더 이동"이 켜져 있는지 확인
- 일부 Exchange 서버는 특수문자 폴더명을 거부합니다. 폴더 접두사를 영문으로 변경해 보세요

### Outlook이 느려짐

- "새 메일 자동 분석"이 켜져 있으면 대량 수신 시 부하가 발생할 수 있습니다. 끄고 수동 분석을 사용하세요
- 동시 요청 수(기본 3)를 줄이면 API 호출 부하가 감소합니다

### Task Pane이 보이지 않음

Outlook 메뉴에서 수동으로 표시할 수 있습니다:
1. 리본의 **보기** 탭 → **작업창** 클릭
2. 또는 Outlook을 재시작

---

## 제거

### Windows 설정에서 제거

1. Windows 설정 → 앱 → 앱 및 기능
2. "MailPrioritizer" 검색 → 제거

### 설정 파일 수동 삭제 (선택)

제거 후에도 설정과 로그 파일은 남아 있습니다. 완전 삭제:

```cmd
rmdir /s /q "%AppData%\MailPrioritizer"
```

> 메일에 저장된 UserProperty(LLM_Priority, LLM_Summary 등)는 각 메일에 남아 있으며 Outlook 동작에 영향을 주지 않습니다.

---

## 프로젝트 구조

```
MailPrioritizer/
├── MailPrioritizer.sln
└── MailPrioritizer/
    ├── MailPrioritizer.csproj       # VSTO 프로젝트 (C# 7.3, .NET 4.7.2)
    ├── packages.config              # NuGet: Newtonsoft.Json 13.0.3
    ├── ThisAddIn.cs                 # VSTO 진입점, 서비스 소유, 이벤트 배선
    ├── Config/
    │   └── ConfigManager.cs         # JSON 설정 로드/저장, DPAPI 토큰 암호화
    ├── Models/
    │   ├── Priority.cs              # 우선순위 enum + 확장 메서드
    │   ├── MailAnalysis.cs          # LLM 분석 결과 DTO
    │   └── AppConfig.cs             # 설정 모델 (4개 중첩 클래스)
    ├── Services/
    │   ├── LlmService.cs            # Claude/OpenAI 듀얼 API 클라이언트
    │   ├── MailProcessor.cs         # 메일 분석 엔진 (단건/배치/통계/내보내기)
    │   ├── FolderManager.cs         # 우선순위 폴더 생성/이동
    │   ├── RetryQueue.cs            # LLM 실패 메일 자동 재시도 큐
    │   └── FeedbackStore.cs         # 사용자 수동 변경 이력/정확도 통계
    ├── Ribbon/
    │   ├── MailRibbon.cs            # 리본 버튼 이벤트 핸들러
    │   └── MailRibbon.xml           # 리본 UI 정의 (5개 버튼)
    ├── Forms/
    │   ├── SettingsForm.cs          # 3탭 환경설정 다이얼로그
    │   └── ProgressForm.cs          # 배치 분석 진행률 폼
    ├── TaskPane/
    │   └── SummaryControl.cs        # 우측 분석 결과 패널
    └── Utils/
        ├── ComHelper.cs             # COM 객체 안전 해제
        ├── Logger.cs                # 파일 기반 로거 (7일 보관)
        └── MailPropertyNames.cs     # UserProperty 이름 상수
```

### 배포 스크립트

| 파일 | 설명 |
|------|------|
| `deploy-clickonce.ps1` | Release 빌드 + VSTO 매니페스트 생성 + publish 패키징 |
| `MailPrioritizer.CLI.targets` | CLI 빌드용 no-op 타겟 (매니페스트 생략) |
| `MailPrioritizer.Publish.targets` | 서명 없는 매니페스트 생성 파이프라인 |

### 관련 문서

- [`CLAUDE.md`](CLAUDE.md) — 개발 가이드 (아키텍처, 제약사항, 설계 결정)
- [`docs/FEATURES.md`](docs/FEATURES.md) — 구현 완료 기능 목록
- [`docs/BUGFIXES.md`](docs/BUGFIXES.md) — 오류 수정 및 코드 개선 이력
- [`docs/ROADMAP.md`](docs/ROADMAP.md) — 향후 구현 예정 아이템

---

## 라이선스

이 프로젝트의 라이선스는 별도로 지정되지 않았습니다.
