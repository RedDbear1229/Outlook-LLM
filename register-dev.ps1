#Requires -Version 5.0
<#
.SYNOPSIS
    MailPrioritizer Debug 빌드를 Outlook VSTO Add-in으로 등록합니다.

.DESCRIPTION
    1. MSBuild로 Debug 빌드 실행
    2. bin\Debug\ 에 MailPrioritizer.dll.manifest 생성
    3. HKCU 레지스트리 키 작성 → Outlook이 Add-in 로드

    이 스크립트는 처음 체크아웃 후 또는 레지스트리가 초기화된 경우 한 번 실행합니다.
    등록 후 Outlook을 재시작하면 "메일 분석" 리본 탭이 나타납니다.

.PARAMETER NoBuild
    MSBuild 단계를 건너뜁니다 (DLL이 이미 최신인 경우).

.PARAMETER NoRestart
    Outlook 재시작 프롬프트를 표시하지 않습니다.

.EXAMPLE
    .\register-dev.ps1
    .\register-dev.ps1 -NoBuild
#>

param(
    [switch]$NoBuild,
    [switch]$NoRestart
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$RepoRoot     = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjFile     = Join-Path $RepoRoot "MailPrioritizer\MailPrioritizer\MailPrioritizer.csproj"
$DebugDir     = Join-Path $RepoRoot "MailPrioritizer\MailPrioritizer\bin\Debug"
$DllPath      = Join-Path $DebugDir "MailPrioritizer.dll"
$ManifestPath = Join-Path $DebugDir "MailPrioritizer.dll.manifest"
$MSBuild      = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
$AddinKey     = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Addins\MailPrioritizer"

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  MailPrioritizer — 개발 환경 등록 스크립트" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""

# ── Step 1: Build ─────────────────────────────────────────────────────────────
if (-not $NoBuild) {
    Write-Host "[1/3] Debug 빌드 중..." -ForegroundColor Cyan
    if (-not (Test-Path $MSBuild)) {
        Write-Error "MSBuild를 찾을 수 없습니다: $MSBuild`nVisual Studio 2022 Community가 설치되어 있는지 확인하세요."
    }
    & $MSBuild $ProjFile /p:Configuration=Debug /v:minimal /nologo
    if ($LASTEXITCODE -ne 0) {
        Write-Error "빌드 실패 (exit code $LASTEXITCODE)"
    }
    Write-Host "      빌드 성공" -ForegroundColor Green
} else {
    Write-Host "[1/3] 빌드 건너뜀 (-NoBuild)" -ForegroundColor Yellow
}

if (-not (Test-Path $DllPath)) {
    Write-Error "DLL을 찾을 수 없습니다: $DllPath`n-NoBuild 없이 실행하거나 먼저 빌드하세요."
}

# ── Step 2: Generate Debug manifest ───────────────────────────────────────────
Write-Host "[2/3] Debug manifest 생성 중..." -ForegroundColor Cyan

$ManifestXml = @"
<?xml version="1.0" encoding="utf-8"?>
<asmv1:assembly manifestVersion="1.0"
  xmlns:asmv1="urn:schemas-microsoft-com:asm.v1"
  xmlns="urn:schemas-microsoft-com:asm.v2"
  xmlns:asmv2="urn:schemas-microsoft-com:asm.v2"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:co.v1="urn:schemas-microsoft-com:clickonce.v1"
  xsi:schemaLocation="urn:schemas-microsoft-com:asm.v1 assembly.adaptive.xsd">
  <asmv1:assemblyIdentity name="MailPrioritizer.dll" version="1.0.0.0"
    publicKeyToken="0000000000000000" language="neutral"
    processorArchitecture="msil" type="win32" />
  <description xmlns="urn:schemas-microsoft-com:asm.v1">MailPrioritizer (Debug)</description>
  <application />
  <entryPoint>
    <co.v1:customHostSpecified />
  </entryPoint>
  <trustInfo>
    <security>
      <applicationRequestMinimum>
        <PermissionSet Unrestricted="true" ID="Custom" SameSite="site" />
        <defaultAssemblyRequest permissionSetReference="Custom" />
      </applicationRequestMinimum>
      <requestedPrivileges xmlns="urn:schemas-microsoft-com:asm.v3">
        <requestedExecutionLevel level="asInvoker" uiAccess="false" />
      </requestedPrivileges>
    </security>
  </trustInfo>
  <vstav3:addIn xmlns:vstav3="urn:schemas-microsoft-com:vsta.v3">
    <vstav3:entryPointsCollection>
      <vstav3:entryPoints>
        <vstav3:entryPoint class="MailPrioritizer.ThisAddIn">
          <assemblyIdentity name="MailPrioritizer" version="1.0.0.0"
            language="neutral" processorArchitecture="msil" />
        </vstav3:entryPoint>
      </vstav3:entryPoints>
    </vstav3:entryPointsCollection>
    <vstav3:update enabled="false" />
    <vstav3:application>
      <vstov4:customizations xmlns:vstov4="urn:schemas-microsoft-com:vsto.v4">
        <vstov4:customization>
          <vstov4:appAddIn application="Outlook" loadBehavior="3" keyName="MailPrioritizer">
            <vstov4:friendlyName>MailPrioritizer</vstov4:friendlyName>
            <vstov4:description>메일 분석 및 우선순위 분류 Add-in (Debug)</vstov4:description>
            <vstov4.1:ribbonTypes xmlns:vstov4.1="urn:schemas-microsoft-com:vsto.v4.1">
              <vstov4.1:ribbonType>MailPrioritizer.Ribbon.MailRibbon</vstov4.1:ribbonType>
            </vstov4.1:ribbonTypes>
          </vstov4:appAddIn>
        </vstov4:customization>
      </vstov4:customizations>
    </vstav3:application>
  </vstav3:addIn>
</asmv1:assembly>
"@

[System.IO.File]::WriteAllText($ManifestPath, $ManifestXml, [System.Text.Encoding]::UTF8)
Write-Host "      Manifest: $ManifestPath" -ForegroundColor Green

# ── Step 3: Write registry ─────────────────────────────────────────────────────
Write-Host "[3/3] 레지스트리 키 작성 중..." -ForegroundColor Cyan

$ManifestUri = "file:///" + $ManifestPath.Replace('\', '/') + "|vstolocal"

if (-not (Test-Path $AddinKey)) {
    New-Item -Path $AddinKey -Force | Out-Null
}

Set-ItemProperty -Path $AddinKey -Name "Description"  -Value "메일 분석 및 우선순위 분류 Add-in" -Type String
Set-ItemProperty -Path $AddinKey -Name "FriendlyName" -Value "MailPrioritizer"                    -Type String
Set-ItemProperty -Path $AddinKey -Name "LoadBehavior" -Value 3                                    -Type DWord
Set-ItemProperty -Path $AddinKey -Name "Manifest"     -Value $ManifestUri                         -Type String

Write-Host "      키:  $AddinKey" -ForegroundColor Green
Write-Host "      URI: $ManifestUri" -ForegroundColor Gray

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Green
Write-Host "  등록 완료! Outlook을 재시작하면 리본 탭이 나타납니다." -ForegroundColor Green
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Green
Write-Host ""

# ── Outlook 재시작 프롬프트 ────────────────────────────────────────────────────
if (-not $NoRestart) {
    $outlookProc = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
    if ($outlookProc) {
        Write-Host "Outlook이 현재 실행 중입니다." -ForegroundColor Yellow
        $ans = Read-Host "지금 Outlook을 종료하고 재시작할까요? [Y/N]"
        if ($ans -match '^[Yy]') {
            $outlookProc | Stop-Process -Force
            Start-Sleep -Seconds 2
            Start-Process "OUTLOOK.EXE"
            Write-Host "Outlook을 재시작했습니다." -ForegroundColor Cyan
        } else {
            Write-Host "Outlook을 수동으로 재시작하세요." -ForegroundColor Yellow
        }
    }
}
