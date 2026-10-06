# PDF 일괄 변환 시작: 도우미가 떠 있으면 그대로 쓰고, 없으면 창 없이 띄운 뒤 브라우저로 화면을 연다.
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$helper = Join-Path $here 'PDF변환_도우미.ps1'
$sess = Join-Path $env:LOCALAPPDATA 'PDF변환\session.json'

function Get-Live {
    if (-not (Test-Path -LiteralPath $sess)) { return $null }
    try { $s = [IO.File]::ReadAllText($sess, [Text.Encoding]::UTF8) | ConvertFrom-Json } catch { return $null }
    try {
        $h = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/health" -f $s.port) -TimeoutSec 2
        if ($h.ok -and $h.pid -eq $s.pid) { return $s }
    } catch {}
    return $null
}

$s = Get-Live
if (-not $s) {
    $ps = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    Start-Process -FilePath $ps -WindowStyle Hidden -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-WindowStyle', 'Hidden', '-File', ('"' + $helper + '"'))
    for ($i = 0; $i -lt 50 -and -not $s; $i++) { Start-Sleep -Milliseconds 300; $s = Get-Live }
}
if (-not $s) {
    Add-Type -AssemblyName System.Windows.Forms
    [void][System.Windows.Forms.MessageBox]::Show("PDF 변환 도우미를 시작하지 못했습니다.`n회사 보안 정책이 PowerShell 스크립트 실행을 막고 있을 수 있습니다.", 'PDF 일괄 변환')
    exit 1
}
$url = "http://127.0.0.1:{0}/?t={1}" -f $s.port, $s.token
if ($env:PDFCONV_NO_OPEN -eq '1') { Write-Output $url; exit 0 }   # 자동 시험용: 브라우저를 열지 않고 주소만 출력
Start-Process $url
