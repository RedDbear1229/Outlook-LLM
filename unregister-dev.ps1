#Requires -Version 5.0
<#
.SYNOPSIS
    Outlook에서 MailPrioritizer Add-in 등록을 제거합니다 (개발 환경용).

.DESCRIPTION
    register-dev.ps1로 등록된 HKCU 레지스트리 키를 삭제합니다.
    Outlook을 재시작하면 리본 탭이 사라집니다.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$AddinKey = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Addins\MailPrioritizer"

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  MailPrioritizer — 개발 환경 등록 해제" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""

# Outlook 실행 중 경고
$outlookProc = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
if ($outlookProc) {
    Write-Host "[경고] Outlook이 실행 중입니다. 종료 후 진행하는 것을 권장합니다." -ForegroundColor Yellow
    $ans = Read-Host "계속하시겠습니까? [Y/N]"
    if ($ans -notmatch '^[Yy]') {
        Write-Host "취소됨." -ForegroundColor Gray
        exit 0
    }
}

if (Test-Path $AddinKey) {
    Remove-Item -Path $AddinKey -Recurse -Force
    Write-Host "[완료] 레지스트리 키 삭제: $AddinKey" -ForegroundColor Green
} else {
    Write-Host "[스킵] 레지스트리 키 없음 (이미 미등록 상태)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Outlook을 재시작하면 리본 탭이 사라집니다." -ForegroundColor Cyan
Write-Host ""
