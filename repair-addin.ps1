#Requires -Version 5.0
<#
.SYNOPSIS
    MailPrioritizer Add-in 등록 상태를 진단하고 자동으로 복구합니다.

.DESCRIPTION
    다음 항목을 점검하여 문제를 찾고 수정합니다:
      1. 레지스트리 키 존재 여부
      2. LoadBehavior 값 (0 = 비활성, 3 = 자동 로드)
      3. Manifest 파일 존재 여부 및 경로 유효성
      4. DLL 파일 존재 여부
      5. Outlook Resiliency 비활성 목록 (충돌로 인해 Outlook이 비활성화한 경우)

    -AutoFix 없이 실행하면 진단 결과만 표시합니다.

.PARAMETER AutoFix
    확인 없이 발견된 모든 문제를 자동으로 수정합니다.

.PARAMETER Build
    수정 전 Debug 빌드를 실행합니다 (DLL이 오래된 경우).

.EXAMPLE
    .\repair-addin.ps1            # 진단만
    .\repair-addin.ps1 -AutoFix  # 진단 + 자동 수정
#>

param(
    [switch]$AutoFix,
    [switch]$Build
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$RepoRoot     = Split-Path -Parent $MyInvocation.MyCommand.Path
$DebugDir     = Join-Path $RepoRoot "MailPrioritizer\MailPrioritizer\bin\Debug"
$ReleaseDir   = Join-Path $RepoRoot "MailPrioritizer\MailPrioritizer\bin\Release"
$ProjFile     = Join-Path $RepoRoot "MailPrioritizer\MailPrioritizer\MailPrioritizer.csproj"
$MSBuild      = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
$AddinKey     = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Addins\MailPrioritizer"
$ResiliencyKey = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems"

$issues   = @()
$warnings = @()

# ── 헬퍼 함수 ─────────────────────────────────────────────────────────────────
function Write-OK   ($msg) { Write-Host "  [OK]   $msg" -ForegroundColor Green }
function Write-WARN ($msg) { Write-Host "  [경고] $msg" -ForegroundColor Yellow; $script:warnings += $msg }
function Write-FAIL ($msg) { Write-Host "  [오류] $msg" -ForegroundColor Red;    $script:issues   += $msg }
function Write-Info ($msg) { Write-Host "         $msg" -ForegroundColor Gray }

function Confirm-Fix ($question) {
    if ($AutoFix) { return $true }
    $ans = Read-Host "  → $question [Y/N]"
    return $ans -match '^[Yy]'
}

# ── 시작 ──────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  MailPrioritizer — Add-in 진단 및 복구 도구" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""

# ── [1] Outlook 실행 확인 ─────────────────────────────────────────────────────
Write-Host "[1] Outlook 실행 상태" -ForegroundColor Cyan
$outlookProc = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
if ($outlookProc) {
    Write-WARN "Outlook이 실행 중입니다. 레지스트리 변경 사항은 재시작 후 반영됩니다."
} else {
    Write-OK "Outlook 미실행 (변경 즉시 반영 가능)"
}

# ── [2] 레지스트리 키 확인 ────────────────────────────────────────────────────
Write-Host ""
Write-Host "[2] 레지스트리 키" -ForegroundColor Cyan

$regExists = Test-Path $AddinKey
if (-not $regExists) {
    Write-FAIL "레지스트리 키 없음: $AddinKey"
    Write-Info "Add-in이 Outlook에 등록되지 않아 로드되지 않습니다."

    if (Confirm-Fix "register-dev.ps1를 실행하여 등록할까요?") {
        & "$RepoRoot\register-dev.ps1" -NoRestart $(if ($Build) { } else { '-NoBuild' })
        $regExists = Test-Path $AddinKey
    }
} else {
    Write-OK "레지스트리 키 존재"
}

# ── [3] LoadBehavior 확인 ─────────────────────────────────────────────────────
Write-Host ""
Write-Host "[3] LoadBehavior" -ForegroundColor Cyan

if ($regExists) {
    $loadBehavior = (Get-ItemProperty -Path $AddinKey -ErrorAction SilentlyContinue).LoadBehavior
    switch ($loadBehavior) {
        3 { Write-OK "LoadBehavior = 3 (자동 로드 — 정상)" }
        2 { Write-WARN "LoadBehavior = 2 (첫 Outlook 시작 시 로드)" }
        0 {
            Write-FAIL "LoadBehavior = 0 — Add-in이 비활성화되어 있습니다."
            Write-Info "Outlook이 오류로 인해 Add-in을 비활성화했을 수 있습니다."
            if (Confirm-Fix "LoadBehavior를 3으로 복원할까요?") {
                Set-ItemProperty -Path $AddinKey -Name "LoadBehavior" -Value 3 -Type DWord
                Write-OK "LoadBehavior = 3 으로 복원됨"
            }
        }
        default { Write-WARN "LoadBehavior = $loadBehavior (알 수 없는 값)" }
    }

    # Manifest 값 확인
    $manifestValue = (Get-ItemProperty -Path $AddinKey -ErrorAction SilentlyContinue).Manifest
    if ($manifestValue) {
        Write-Info "Manifest = $manifestValue"
    } else {
        Write-FAIL "Manifest 레지스트리 값이 없습니다."
        if (Confirm-Fix "register-dev.ps1를 다시 실행하여 Manifest 값을 작성할까요?") {
            & "$RepoRoot\register-dev.ps1" -NoRestart -NoBuild
        }
    }
} else {
    Write-Info "레지스트리 키가 없으므로 건너뜁니다."
}

# ── [4] Manifest 파일 존재 확인 ───────────────────────────────────────────────
Write-Host ""
Write-Host "[4] Manifest 파일" -ForegroundColor Cyan

$manifestFilePath = $null
if ($regExists) {
    $manifestValue = (Get-ItemProperty -Path $AddinKey -ErrorAction SilentlyContinue).Manifest
    if ($manifestValue) {
        # "file:///C:/path/to/file.dll.manifest|vstolocal" 에서 경로 추출
        $manifestFilePath = $manifestValue `
            -replace '^\s*file:///', '' `
            -replace '\|vstolocal\s*$', '' `
            -replace '/', '\'
    }
}

# 등록된 manifest 경로가 없으면 Debug 경로를 기본값으로
if (-not $manifestFilePath) {
    $manifestFilePath = Join-Path $DebugDir "MailPrioritizer.dll.manifest"
}

if (Test-Path $manifestFilePath) {
    Write-OK "Manifest 파일 존재: $manifestFilePath"
} else {
    Write-FAIL "Manifest 파일 없음: $manifestFilePath"
    Write-Info "Debug 빌드 후 register-dev.ps1을 실행하면 생성됩니다."

    if (Confirm-Fix "지금 register-dev.ps1을 실행하여 Manifest를 생성할까요?") {
        & "$RepoRoot\register-dev.ps1" -NoRestart $(if ($Build) { } else { '-NoBuild' })
    }
}

# ── [5] DLL 파일 존재 확인 ────────────────────────────────────────────────────
Write-Host ""
Write-Host "[5] DLL 파일" -ForegroundColor Cyan

$debugDll   = Join-Path $DebugDir   "MailPrioritizer.dll"
$releaseDll = Join-Path $ReleaseDir "MailPrioritizer.dll"

if (Test-Path $debugDll) {
    $dllDate = (Get-Item $debugDll).LastWriteTime
    Write-OK "Debug DLL 존재 (수정: $($dllDate.ToString('yyyy-MM-dd HH:mm')))"
} else {
    Write-FAIL "Debug DLL 없음: $debugDll"
    if ($Build -or (Confirm-Fix "지금 Debug 빌드를 실행할까요?")) {
        if (-not (Test-Path $MSBuild)) {
            Write-Host "  [오류] MSBuild를 찾을 수 없습니다: $MSBuild" -ForegroundColor Red
        } else {
            & $MSBuild $ProjFile /p:Configuration=Debug /v:minimal /nologo
        }
    }
}

if (Test-Path $releaseDll) {
    $dllDate = (Get-Item $releaseDll).LastWriteTime
    Write-OK "Release DLL 존재 (수정: $($dllDate.ToString('yyyy-MM-dd HH:mm')))"
} else {
    Write-WARN "Release DLL 없음 (Release 빌드 미실행)"
}

# ── [6] Resiliency 비활성 목록 확인 ──────────────────────────────────────────
Write-Host ""
Write-Host "[6] Outlook 비활성 Add-in 목록 (Resiliency)" -ForegroundColor Cyan

$disabledFound = $false
if (Test-Path $ResiliencyKey) {
    $resiliencyProps = Get-ItemProperty -Path $ResiliencyKey -ErrorAction SilentlyContinue
    if ($resiliencyProps) {
        foreach ($prop in $resiliencyProps.PSObject.Properties) {
            if ($prop.Name -match '^PS') { continue }  # PS 내부 속성 제외
            try {
                $rawBytes = $prop.Value
                if ($rawBytes -is [byte[]]) {
                    $text = [System.Text.Encoding]::Unicode.GetString($rawBytes)
                    if ($text -imatch 'MailPrioritizer') {
                        $disabledFound = $true
                        Write-FAIL "Outlook이 Add-in을 비활성화했습니다: $($prop.Name)"
                        Write-Info "값: $text"
                        if (Confirm-Fix "비활성 항목을 삭제하여 복구할까요?") {
                            Remove-ItemProperty -Path $ResiliencyKey -Name $prop.Name -Force
                            Write-OK "비활성 항목 삭제됨"
                        }
                    }
                }
            } catch { }
        }
    }
}

if (-not $disabledFound) {
    Write-OK "비활성 목록에 MailPrioritizer 없음"
}

# ── [7] 요약 ─────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  진단 결과 요약" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan

if ($issues.Count -eq 0 -and $warnings.Count -eq 0) {
    Write-Host ""
    Write-Host "  모든 항목 정상입니다." -ForegroundColor Green
    Write-Host "  리본이 표시되지 않으면 Outlook을 재시작해 보세요." -ForegroundColor Green
} else {
    if ($issues.Count -gt 0) {
        Write-Host ""
        Write-Host "  오류 ($($issues.Count)건):" -ForegroundColor Red
        $issues | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
    }
    if ($warnings.Count -gt 0) {
        Write-Host ""
        Write-Host "  경고 ($($warnings.Count)건):" -ForegroundColor Yellow
        $warnings | ForEach-Object { Write-Host "    - $_" -ForegroundColor Yellow }
    }
    if (-not $AutoFix) {
        Write-Host ""
        Write-Host "  -AutoFix 옵션으로 자동 수정하거나," -ForegroundColor Gray
        Write-Host "  register-dev.ps1을 직접 실행하세요." -ForegroundColor Gray
    }
}

Write-Host ""

# ── [8] Outlook 재시작 안내 ───────────────────────────────────────────────────
if (($issues.Count -gt 0 -or $warnings.Count -gt 0) -and $outlookProc) {
    $ans = Read-Host "변경 사항 반영을 위해 Outlook을 재시작할까요? [Y/N]"
    if ($ans -match '^[Yy]') {
        $outlookProc | Stop-Process -Force
        Start-Sleep -Seconds 2
        Start-Process "OUTLOOK.EXE"
        Write-Host "Outlook을 재시작했습니다." -ForegroundColor Cyan
    }
}
