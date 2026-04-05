#Requires -Version 5.1
<#
.SYNOPSIS
    MailPrioritizer VSTO one-click build & deploy script.

.DESCRIPTION
    Builds Release, generates VSTO manifests (unsigned), and packages
    the publish\ folder with install/uninstall .bat wrappers.

    Output files in publish\:
      MailPrioritizer.vsto          -- deployment manifest (share this)
      MailPrioritizer.dll.manifest  -- application manifest
      MailPrioritizer.dll           -- add-in DLL
      Newtonsoft.Json.dll           -- dependency
      install.bat                   -- VSTOInstaller wrapper
      uninstall.bat                 -- uninstall wrapper

    Installation on target PC:
      1. Copy publish\ folder to shared drive or USB
      2. Run install.bat (or double-click .vsto)
      3. Click [Install] on the security dialog

.PARAMETER PublishVersion
    Version string for the deployment manifest (default: 1.0.0.0)

.PARAMETER PublishDir
    Output folder (default: <repo-root>\publish\)

.PARAMETER Open
    Open publish folder in Explorer after build

.EXAMPLE
    .\deploy-clickonce.ps1
    .\deploy-clickonce.ps1 -Open
    .\deploy-clickonce.ps1 -PublishVersion 1.0.0.5 -Open
#>
param(
    [string]$PublishVersion = "1.0.0.0",
    [string]$PublishDir     = "",
    [switch]$Open
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Paths
$ScriptDir  = $PSScriptRoot
$ProjectDir = Join-Path $ScriptDir "MailPrioritizer\MailPrioritizer"
$ProjFile   = Join-Path $ProjectDir "MailPrioritizer.csproj"
$BinRelease = Join-Path $ProjectDir "bin\Release"

if ($PublishDir -eq "") {
    $PublishDir = Join-Path $ScriptDir "publish"
}
$PublishDir = $PublishDir.TrimEnd('\').TrimEnd('/')

$MSBuild = @(
    "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $MSBuild) {
    Write-Error "MSBuild.exe not found. Ensure VS 2022 is installed."
    exit 1
}

Write-Host ""
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "  MailPrioritizer VSTO 배포 빌드" -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "  버전: $PublishVersion"
Write-Host "  출력: $PublishDir"
Write-Host ""

# [1/3] Release build
Write-Host "[1/3] Release 빌드 중..." -ForegroundColor Yellow

& $MSBuild $ProjFile `
    /t:Build `
    /p:Configuration=Release `
    /p:IsVstoPublish=false `
    /v:minimal /nologo

if ($LASTEXITCODE -ne 0) {
    Write-Error "빌드 실패 (exit code $LASTEXITCODE)"
    exit $LASTEXITCODE
}
Write-Host "  완료." -ForegroundColor Green

# [2/3] Generate VSTO manifests (unsigned)
Write-Host "[2/3] VSTO 매니페스트 생성 중 (v$PublishVersion)..." -ForegroundColor Yellow

& $MSBuild $ProjFile `
    /t:VisualStudioForApplicationsBuild `
    /p:Configuration=Release `
    /p:IsVstoPublish=true `
    /p:PublishVersion=$PublishVersion `
    /v:minimal /nologo

if ($LASTEXITCODE -ne 0) {
    Write-Error "매니페스트 생성 실패 (exit code $LASTEXITCODE)"
    exit $LASTEXITCODE
}

foreach ($f in @("MailPrioritizer.dll.manifest", "MailPrioritizer.vsto")) {
    if (-not (Test-Path (Join-Path $BinRelease $f))) {
        Write-Error "Expected manifest not found: $f"
        exit 1
    }
}
Write-Host "  완료." -ForegroundColor Green

# [3/3] Package publish folder
Write-Host "[3/3] 배포 패키지 생성 중..." -ForegroundColor Yellow

if (Test-Path $PublishDir) {
    Remove-Item $PublishDir -Recurse -Force
}
New-Item $PublishDir -ItemType Directory -Force | Out-Null

$copyFiles = @(
    "MailPrioritizer.vsto",
    "MailPrioritizer.dll.manifest",
    "MailPrioritizer.dll",
    "MailPrioritizer.dll.config",
    "Microsoft.Office.Tools.Common.v4.0.Utilities.dll",
    "Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll",
    "Newtonsoft.Json.dll"
)
foreach ($f in $copyFiles) {
    $src = Join-Path $BinRelease $f
    if (Test-Path $src) {
        Copy-Item $src $PublishDir -Force
        Write-Host "    + $f"
    }
}

# Write install.bat
$installBat = Join-Path $PublishDir "install.bat"
$installLines = @(
    "@echo off",
    "echo MailPrioritizer VSTO Add-in 설치 중...",
    "echo.",
    'if not exist "%CommonProgramFiles%\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" (',
    "    echo [오류] VSTO Runtime이 설치되어 있지 않습니다.",
    "    pause",
    "    exit /b 1",
    ")",
    '"%CommonProgramFiles%\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" /install "%~dp0MailPrioritizer.vsto"',
    "if errorlevel 1 (",
    "    echo.",
    "    echo [실패] 보안 경고가 표시되면 [설치] 버튼을 클릭하세요.",
    "    pause",
    "    exit /b 1",
    ")",
    "echo.",
    "echo [완료] 설치 완료. Outlook을 재시작하세요.",
    "pause"
)
[System.IO.File]::WriteAllLines($installBat, $installLines, [System.Text.Encoding]::GetEncoding(949))

# Write uninstall.bat
$uninstallBat = Join-Path $PublishDir "uninstall.bat"
$uninstallLines = @(
    "@echo off",
    "echo MailPrioritizer VSTO Add-in 제거 중...",
    "echo.",
    '"%CommonProgramFiles%\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" /uninstall "%~dp0MailPrioritizer.vsto"',
    "if errorlevel 1 (",
    "    echo [실패] 제거에 실패했습니다.",
    "    pause",
    "    exit /b 1",
    ")",
    "echo.",
    "echo [완료] 제거 완료.",
    "pause"
)
[System.IO.File]::WriteAllLines($uninstallBat, $uninstallLines, [System.Text.Encoding]::GetEncoding(949))

Write-Host "    + install.bat"
Write-Host "    + uninstall.bat"
Write-Host "  완료." -ForegroundColor Green

$publishedFiles = (Get-ChildItem $PublishDir -File).Count
Write-Host ""
Write-Host "====================================================" -ForegroundColor Green
Write-Host "  배포 완료!  ($publishedFiles 개 파일)" -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  출력 폴더: $PublishDir"
Write-Host ""
Write-Host "  배포 방법:"
Write-Host "    1. publish\ 폴더 전체를 공유 드라이브 / USB에 복사"
Write-Host "    2. 대상 PC에서 install.bat 실행 (또는 .vsto 더블클릭)"
Write-Host "    3. 보안 경고 대화상자에서 [설치] 클릭"
Write-Host ""
Write-Host "  제거 방법:"
Write-Host "    uninstall.bat 실행"
Write-Host ""

if ($Open) {
    Start-Process explorer.exe $PublishDir
}
