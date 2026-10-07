# PDF 일괄 변환 도우미 (Windows PowerShell 5.1) - 파이썬 설치 없이 윈도우 내장 PowerShell만으로 동작
# 한글(HWP/HWPX)·워드(DOC/DOCX)·파워포인트(PPT/PPTX)를 각 프로그램의 자동화 기능으로 PDF로 바꾼다.
# 화면(PDF변환.html)도 이 도우미가 http://127.0.0.1:<포트>/ 에서 직접 제공한다(이 PC 밖에서는 접속 불가).
# 보안: 실행할 때마다 만드는 일회용 비밀번호(X-Token)와 출처(Origin) 확인. 사용자가 선택 창으로 고른 폴더·파일과
#       다운로드 폴더 안에서만 읽고 쓴다. 원본은 읽기 전용으로 열고 고치거나 지우지 않는다.
#       화면에 끌어다 놓은 파일은 브라우저가 원래 위치를 알려 주지 않으므로, 도우미 전용 임시 폴더
#       (%LOCALAPPDATA%\PDF변환\받은파일)에 사본을 받아 변환하고, 도우미가 꺼질 때 그 사본만 지운다.
# 실행: powershell -NoProfile -ExecutionPolicy Bypass -STA -File PDF변환_도우미.ps1   (보통은 PDF변환_시작.ps1이 띄운다)
# 제작: 해무기(henmoogi), 2026 · MIT License
[CmdletBinding(PositionalBinding = $false)]
param([int]$Port = 43141, [int]$IdleMinutes = 30, [int]$TimeoutSec = 120, [string]$Allow = '')   # -Allow '폴더1;폴더2': 자동 시험용으로 폴더를 미리 허용(명령줄 전용)
$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
Add-Type -AssemblyName System.Windows.Forms
$VERSION = '2026.10.07'
$HERE = Split-Path -Parent $MyInvocation.MyCommand.Path
$HTML = Join-Path $HERE 'PDF변환.html'
$SESS_DIR = Join-Path $env:LOCALAPPDATA 'PDF변환'
$SESS_FILE = Join-Path $SESS_DIR 'session.json'
$STAGE_ROOT = Join-Path $SESS_DIR '받은파일'
$STAGE = Join-Path $STAGE_ROOT ([string]$PID)   # 끌어다 놓은 파일의 사본(이 도우미 전용, 꺼질 때 지움)
$TOKEN = [guid]::NewGuid().ToString('N') + [guid]::NewGuid().ToString('N')
$DOWNLOADS = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path
$EXT = @{ '.hwp' = 'hwp'; '.hwpx' = 'hwp'; '.doc' = 'word'; '.docx' = 'word'; '.ppt' = 'ppt'; '.pptx' = 'ppt' }

# 작업 스레드와 나눠 쓰는 상태
$S = [hashtable]::Synchronized(@{ allowed = (New-Object System.Collections.ArrayList); allowedFiles = (New-Object System.Collections.ArrayList);
        job = $null; lastJob = $null; stop = $false; lastDir = '';
        itemStarted = $null; enginePid = 0; timedOut = $false; creating = $null; closeStarted = $null; last = Get-Date })

function Send-Json($ctx, $obj, [int]$status = 200) {
    $b = [Text.Encoding]::UTF8.GetBytes(($obj | ConvertTo-Json -Depth 8 -Compress))
    $r = $ctx.Response; $r.StatusCode = $status; $r.ContentType = 'application/json; charset=utf-8'
    $r.Headers.Add('Cache-Control', 'no-store'); $r.ContentLength64 = $b.Length; $r.OutputStream.Write($b, 0, $b.Length); $r.Close()
}
function Send-Html($ctx) {
    $b = [IO.File]::ReadAllBytes($HTML)
    $r = $ctx.Response; $r.ContentType = 'text/html; charset=utf-8'; $r.Headers.Add('Cache-Control', 'no-store')
    $r.ContentLength64 = $b.Length; $r.OutputStream.Write($b, 0, $b.Length); $r.Close()
}
function Read-Body($req) {
    if (-not $req.HasEntityBody) { return @{} }
    $sr = New-Object IO.StreamReader($req.InputStream, [Text.Encoding]::UTF8); $t = $sr.ReadToEnd(); $sr.Close()
    if (-not $t) { return @{} }
    return ($t | ConvertFrom-Json)
}
function Get-Q($req, [string]$name) {
    foreach ($pair in $req.Url.Query.TrimStart('?').Split('&')) {
        $i = $pair.IndexOf('='); if ($i -lt 0) { continue }
        if ($pair.Substring(0, $i) -eq $name) { return [uri]::UnescapeDataString($pair.Substring($i + 1)) }
    }
    return ''
}
function Norm([string]$p) { return [IO.Path]::GetFullPath($p).TrimEnd('\') }
function Is-Under([string]$full, [string]$root) {
    $r = Norm $root
    return ($full -ieq $r -or $full.StartsWith($r + '\', [StringComparison]::OrdinalIgnoreCase))
}
function Test-Allowed([string]$p) {   # 고른 폴더와 다운로드 폴더(저장 위치·폴더 목록용)
    if (-not $p) { return $false }
    try { $full = Norm $p } catch { return $false }
    foreach ($root in @($S.allowed) + @($DOWNLOADS)) { if (Is-Under $full $root) { return $true } }
    return $false
}
function Test-Source([string]$p) {    # 변환할 문서: 고른 폴더 안, 하나씩 고른 파일, 끌어다 놓아 받은 사본만
    if (-not $p) { return $false }
    try { $full = Norm $p } catch { return $false }
    if (Is-Under $full $STAGE) { return $true }
    foreach ($f in @($S.allowedFiles)) { if ($full -ieq $f) { return $true } }
    foreach ($root in @($S.allowed)) { if (Is-Under $full $root) { return $true } }
    return $false
}
function Is-Doc([string]$name) { return ($EXT.ContainsKey([IO.Path]::GetExtension($name).ToLower()) -and -not ([IO.Path]::GetFileName($name)).StartsWith('~$')) }
# 화면이 보낸 상대 경로를 안전한 폴더\파일 이름으로(.. 와 드라이브 문자, 쓸 수 없는 글자를 걸러 저장 위치 밖으로 못 나가게)
$BADCH = [IO.Path]::GetInvalidFileNameChars()
function Safe-Rel([string]$rel) {
    $out = New-Object System.Collections.ArrayList
    foreach ($part in ($rel -split '[\\/]+')) {
        $q = -join @($part.ToCharArray() | ForEach-Object { if ($BADCH -contains $_) { '_' } else { $_ } })
        $q = $q.Trim().TrimEnd('.')
        if ($q -and $q -ne '..') { [void]$out.Add($q) }
    }
    return ($out -join '\')
}
function New-Owner {   # 선택 창이 다른 창 뒤에 숨지 않게 맨 앞에 뜨는 투명한 주인 창
    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost = $true; $owner.ShowInTaskbar = $false; $owner.Opacity = 0; $owner.StartPosition = 'CenterScreen'
    $owner.Show(); $owner.Activate()
    return $owner
}
function Pick-Folder([string]$desc, [string]$start) {
    $owner = New-Owner
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = $desc; $dlg.ShowNewFolderButton = $true
    if ($start -and (Test-Path -LiteralPath $start)) { $dlg.SelectedPath = $start } elseif ($S.lastDir -and (Test-Path -LiteralPath $S.lastDir)) { $dlg.SelectedPath = $S.lastDir }
    $res = $dlg.ShowDialog($owner); $owner.Close(); $owner.Dispose()
    if ($res -eq [System.Windows.Forms.DialogResult]::OK) { $S.lastDir = $dlg.SelectedPath; return $dlg.SelectedPath }
    return $null
}
function Pick-Files {
    $owner = New-Owner
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = '변환할 문서를 고르세요 (Ctrl이나 Shift를 누른 채 누르면 여러 개)'
    $dlg.Filter = '한글·워드·파워포인트 문서|*.hwp;*.hwpx;*.doc;*.docx;*.ppt;*.pptx|모든 파일|*.*'
    $dlg.Multiselect = $true
    if ($S.lastDir -and (Test-Path -LiteralPath $S.lastDir)) { $dlg.InitialDirectory = $S.lastDir }
    $res = $dlg.ShowDialog($owner); $owner.Close(); $owner.Dispose()
    if ($res -eq [System.Windows.Forms.DialogResult]::OK -and $dlg.FileNames.Count) { $S.lastDir = Split-Path -Parent $dlg.FileNames[0]; return @($dlg.FileNames) }
    return @()
}
function Clear-Stage { if (Test-Path -LiteralPath $STAGE) { Remove-Item -LiteralPath $STAGE -Recurse -Force -ErrorAction SilentlyContinue } }
function Get-Targets([string]$folder, [bool]$rec) {
    $opt = @{ LiteralPath = $folder; File = $true; ErrorAction = 'SilentlyContinue' }
    if ($rec) { $opt.Recurse = $true }
    @(Get-ChildItem @opt | Where-Object { Is-Doc $_.Name } | Sort-Object FullName)
}
function Get-Apps {
    @{ hwp = (Test-Path 'Registry::HKEY_CLASSES_ROOT\HWPFrame.HwpObject'); word = (Test-Path 'Registry::HKEY_CLASSES_ROOT\Word.Application');
        ppt = (Test-Path 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application') }
}
function Get-HwpModule {
    $v = (Get-ItemProperty -Path 'HKCU:\Software\HNC\HwpAutomation\Modules' -ErrorAction SilentlyContinue).FilePathCheckerModule
    return [bool]($v -and (Test-Path -LiteralPath $v))
}

# 변환할 목록 만들기: 화면 목록(원본 경로 + 목록에 보이는 상대 경로)으로 저장 위치 계산, 같은 이름 충돌 구분, 이미 있는 PDF 건너뛰기
function Plan-Items($list, [string]$outMode, [string]$outFolder, [bool]$overwrite) {
    $outRoot = $null
    if ($outMode -eq 'downloads') {
        $base = Join-Path $DOWNLOADS ('PDF변환_' + (Get-Date -Format 'yyyyMMdd_HHmm')); $outRoot = $base; $n = 2
        while (Test-Path -LiteralPath $outRoot) { $outRoot = "$base`_$n"; $n++ }
    } elseif ($outMode -eq 'folder') { $outRoot = Norm $outFolder }
    $items = New-Object System.Collections.ArrayList
    $seen = @{}
    foreach ($x in $list) {
        $src = Norm ([string]$x.path)
        if ($seen.ContainsKey($src.ToLower())) { continue }
        $seen[$src.ToLower()] = 1
        $f = Get-Item -LiteralPath $src
        $rel = Safe-Rel ([string]$x.rel); if (-not $rel) { $rel = $f.Name }
        $relDir = Split-Path -Parent $rel
        $dstDir = if ($outMode -eq 'beside') { $f.DirectoryName } elseif ($relDir) { Join-Path $outRoot $relDir } else { $outRoot }
        [void]$items.Add(@{ i = $items.Count; rel = $rel; src = $f.FullName; kind = $EXT[$f.Extension.ToLower()]; ext = $f.Extension.TrimStart('.').ToLower();
                size = $f.Length; dstDir = $dstDir; base = [IO.Path]::GetFileNameWithoutExtension($f.Name); dst = ''; status = '대기'; msg = ''; sec = $null })
    }
    # 같은 폴더에 이름만 같고 확장자가 다른 문서(계약서.hwp, 계약서.docx)는 PDF 이름 뒤에 _확장자를 붙여 구분하고,
    # 그래도 겹치면(다른 곳에서 고른 같은 이름의 문서) _2, _3을 붙인다
    $exts = @{}   # 같은 저장 위치·같은 이름에 확장자가 몇 종류인지
    foreach ($it in $items) { $k = ($it.dstDir + '\' + $it.base).ToLower(); if (-not $exts[$k]) { $exts[$k] = @{} }; $exts[$k][$it.ext] = 1 }
    $used = @{}
    foreach ($it in $items) {
        $stem = if ($exts[($it.dstDir + '\' + $it.base).ToLower()].Count -gt 1) { "$($it.base)_$($it.ext)" } else { $it.base }
        $dst = Join-Path $it.dstDir "$stem.pdf"; $n = 2
        while ($used.ContainsKey($dst.ToLower())) { $dst = Join-Path $it.dstDir "$stem`_$n.pdf"; $n++ }
        $used[$dst.ToLower()] = 1
        $it.dst = $dst
        if (-not $overwrite -and (Test-Path -LiteralPath $it.dst)) { $it.status = '건너뜀'; $it.msg = '같은 이름의 PDF가 이미 있습니다' }
    }
    return @{ items = $items; outRoot = $outRoot }
}

# ---------- 작업 스레드(변환) ----------
$WORKER = {
    param($S)
    $ErrorActionPreference = 'Stop'
    function Wait-Stable([string]$p, [int]$timeout) {
        $dl = (Get-Date).AddSeconds($timeout); $last = -1; $stable = 0
        while ((Get-Date) -lt $dl) {
            if (Test-Path -LiteralPath $p) {
                $sz = (Get-Item -LiteralPath $p).Length
                if ($sz -gt 0 -and $sz -eq $last) { $stable++; if ($stable -ge 3) { return $true } } else { $stable = 0 }
                $last = $sz
            }
            Start-Sleep -Milliseconds 500
        }
        return (Test-Path -LiteralPath $p)
    }
    function New-Engine([string]$kind) {
        $pname = @{ hwp = 'Hwp'; word = 'WINWORD'; ppt = 'POWERPNT' }[$kind]
        $before = @(Get-Process -Name $pname -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        $S.creating = @{ pname = $pname; before = $before; started = Get-Date }   # 띄우다 멈추면 주 스레드가 새로 뜬 것만 닫는다
        if ($kind -eq 'hwp') {
            $app = New-Object -ComObject HWPFrame.HwpObject
            try { [void]$app.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule') } catch {}
            try { [void]$app.SetMessageBoxMode(0x00214411) } catch {}   # 확인 대화상자를 기본 단추로 자동 응답
            try { $app.XHwpWindows.Item(0).Visible = $false } catch {}
        } elseif ($kind -eq 'word') {
            $app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
        } else {
            $app = New-Object -ComObject PowerPoint.Application
        }
        $S.creating = $null
        $after = @(Get-Process -Name $pname -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        $new = @($after | Where-Object { $before -notcontains $_ })
        $myPid = 0; if ($new.Count -eq 1) { $myPid = $new[0] }
        $shared = $false
        if ($kind -eq 'ppt') {
            # 파워포인트는 한 번에 하나만 뜨는 프로그램이다. 사용자가 쓰는 중(발표 자료가 열려 있거나 화면에 보임)이면
            # 그 파워포인트를 빌려 쓰고 끝나도 닫지 않는다. 아무것도 안 열린 숨은 파워포인트는 도우미 몫으로 보고 닫는다.
            try { $shared = ($app.Presentations.Count -gt 0) -or ($app.Visible -ne 0) } catch { $shared = $true }
            if (-not $shared -and $myPid -eq 0 -and $after.Count -eq 1) { $myPid = $after[0] }
            if ($shared) { $myPid = 0 }
        }
        $eng = @{ kind = $kind; app = $app; pid = $myPid; shared = $shared }
        $S.enginePid = $myPid
        return $eng
    }
    function Close-Engine($eng) {
        if (-not $eng) { return }
        if (-not $eng.shared -and $eng.pid) { $S.enginePid = $eng.pid; $S.closeStarted = Get-Date }   # 닫다가 멈추면 주 스레드가 강제로 닫는다
        if (-not $eng.shared) { try { [void]$eng.app.Quit() } catch {} }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($eng.app) } catch {}
        # 닫은 프로그램이 완전히 꺼질 때까지 잠깐 기다린다(다음 작업이 꺼지는 중인 프로그램에 붙지 않게)
        if (-not $eng.shared -and $eng.pid) {
            for ($w = 0; $w -lt 20 -and (Get-Process -Id $eng.pid -ErrorAction SilentlyContinue); $w++) { Start-Sleep -Milliseconds 250 }
        }
        $S.closeStarted = $null
        $S.enginePid = 0
    }
    function Convert-Word($app, [string]$src, [string]$dst) {
        $m = [Type]::Missing
        # 가짜 암호를 넘겨, 암호 문서는 입력 창 대신 오류로 끝나게 한다(암호 없는 문서는 그대로 열림)
        $doc = $app.Documents.Open($src, $false, $true, $false, '__no_password__', '', $false, '', '', 0, $m, $false)
        try { [void]$doc.ExportAsFixedFormat($dst, 17, $false, 0, 0, 1, 1, 0, $true, $true, 1, $true, $true, $false) }
        finally { try { [void]$doc.Close(0) } catch {} }
    }
    function Convert-Ppt($app, [string]$src, [string]$dst) {
        $pres = $app.Presentations.Open($src + '::__no_password__::', -1, 0, 0)   # 읽기 전용, 창 없이
        try { try { [void]$pres.ExportAsFixedFormat($dst, 2, 2) } catch { [void]$pres.SaveAs($dst, 32) } }
        finally { try { [void]$pres.Close() } catch {} }
    }
    function Convert-Hwp($app, [string]$src, [string]$dst) {
        # 포맷은 ""로 두어 한글이 실제 형식을 알아서 판단(확장자와 내용이 다른 문서 대비), 잠금·버전 경고 없이 열기
        $opened = $app.Open($src, '', 'lock:false;forceopen:true;versionwarning:false;')
        if (-not $opened) { throw '문서를 열지 못했습니다(암호가 걸렸거나 손상된 문서일 수 있습니다).' }
        try {
            # 1순위: 'PDF로 인쇄'(한글 자체 Hancom PDF)에 인쇄 방식을 '보통(0)'으로 지정한다.
            #   문서에 '2쪽 모아 찍기' 같은 인쇄 설정이 저장돼 있으면 'PDF로 저장'은 그 설정을 따라 쪽을 0.707배로 줄여 버린다.
            #   인쇄 경로라 저장이 잠긴 배포용 문서도 이 방법으로 된다.
            try {
                $pp = $app.HParameterSet.HPrint
                [void]$app.HAction.GetDefault('PrintToPDF', $pp.HSet)
                $pp.PrintMethod = 0; $pp.ZoomX = 100; $pp.ZoomY = 100; $pp.Range = 0; $pp.NumCopy = 1; $pp.filename = $dst
                if ($app.HAction.Execute('PrintToPDF', $pp.HSet)) { [void](Wait-Stable $dst 60) }
            } catch {}
            # 2순위: 'PDF로 저장' 동작, 3순위: SaveAs
            if (-not (Test-Path -LiteralPath $dst)) {
                $pset = $app.HParameterSet.HFileOpenSave
                [void]$app.HAction.GetDefault('FileSaveAsPdf', $pset.HSet)
                $pset.filename = $dst; $pset.Format = 'PDF'; $pset.Attributes = 0
                $ok = $app.HAction.Execute('FileSaveAsPdf', $pset.HSet)
                if (-not $ok -or -not (Test-Path -LiteralPath $dst)) { try { [void]$app.SaveAs($dst, 'PDF', '') } catch {} }
            }
            if (-not (Test-Path -LiteralPath $dst)) {
                # 배포용 문서(저장 잠금, 인쇄만 허용)는 인쇄 방식 PDF로 우회. 별도 변환기가 비동기로 쓰므로 완성될 때까지 기다린다
                try {
                    $x = $app.XHwpDocuments.Item(0).XHwpPrint
                    try { $x.filename = $dst } catch {}
                    try { [void]$x.RunToPDF($dst) } catch { try { [void]$x.RunToPDF() } catch {} }
                } catch {}
                [void](Wait-Stable $dst 90)
            }
            if (-not (Test-Path -LiteralPath $dst)) {
                $locked = $false; try { $locked = ($app.EditMode -eq 0) } catch {}
                if ($locked) { throw '배포용·보호 문서라 저장과 인쇄가 모두 잠겨 있습니다.' }
            }
        } finally { try { [void]$app.Clear(1) } catch {} }
    }
    function Friendly([string]$m) {
        if ($m -match '암호|password|Password|0x800A1520') { return '암호가 걸린 문서라 열 수 없습니다.' }
        if ($m -match 'RPC|0x800706BA|0x80010108|원격 프로시저') { return '프로그램이 응답하지 않아 멈췄습니다.' }
        return $m
    }

    $job = $S.job
    $engines = @{}
    try {
        foreach ($k in @('hwp', 'word', 'ppt')) {
            $group = @($job.items | Where-Object { $_.kind -eq $k -and $_.status -eq '대기' })
            if ($group.Count -eq 0) { continue }
            foreach ($it in $group) {
                if ($S.stop) { break }
                $it.status = '변환 중'; $job.current = $it.i; $S.timedOut = $false; $S.itemStarted = Get-Date
                $t0 = Get-Date
                try {
                    if (-not (Test-Path -LiteralPath $it.dstDir)) { [void](New-Item -ItemType Directory -Force -Path $it.dstDir) }
                    if (-not $engines[$k]) { $engines[$k] = New-Engine $k }
                    $app = $engines[$k].app
                    if ($k -eq 'hwp') { Convert-Hwp $app $it.src $it.dst } elseif ($k -eq 'word') { Convert-Word $app $it.src $it.dst } else { Convert-Ppt $app $it.src $it.dst }
                    if (Test-Path -LiteralPath $it.dst) { $it.status = '완료' } else { $it.status = '실패'; $it.msg = 'PDF가 만들어지지 않았습니다.' }
                } catch {
                    $it.status = '실패'
                    $it.msg = if ($S.timedOut) { "$($S.timeoutSec)초 넘게 응답이 없어 멈췄습니다." } else { Friendly $_.Exception.Message }
                    $e = $engines[$k]
                    if ($e -and ($S.timedOut -or ($e.pid -and -not (Get-Process -Id $e.pid -ErrorAction SilentlyContinue)))) { $engines[$k] = $null }
                }
                $S.itemStarted = $null
                $it.sec = [math]::Round(((Get-Date) - $t0).TotalSeconds, 1)
            }
            if ($engines[$k]) { Close-Engine $engines[$k]; $engines[$k] = $null }
            if ($S.stop) { break }
        }
    } finally {
        foreach ($e in @($engines.Values)) { Close-Engine $e }
        foreach ($it in $job.items) { if ($it.status -eq '대기' -or $it.status -eq '변환 중') { $it.status = '취소'; $it.msg = '멈춤을 눌러 변환하지 않았습니다.' } }
        $job.current = -1; $job.finished = (Get-Date).ToString('s')
        $job.state = if ($S.stop) { 'stopped' } else { 'done' }
    }
}

$script:ps = $null; $script:rs = $null; $script:handle = $null
function Start-Worker {
    $script:rs = [runspacefactory]::CreateRunspace(); $script:rs.ApartmentState = 'STA'; $script:rs.ThreadOptions = 'ReuseThread'; $script:rs.Open()
    $script:ps = [powershell]::Create(); $script:ps.Runspace = $script:rs
    [void]$script:ps.AddScript($WORKER).AddArgument($S)
    $script:handle = $script:ps.BeginInvoke()
}
function Reap-Worker {
    if ($script:handle -and $script:handle.IsCompleted) {
        try { [void]$script:ps.EndInvoke($script:handle) } catch { if ($S.job) { $S.job.error = $_.Exception.Message; $S.job.state = 'done' } }
        $script:ps.Dispose(); $script:rs.Dispose(); $script:ps = $null; $script:rs = $null; $script:handle = $null
    }
}
function Job-Running { return ($S.job -and $S.job.state -eq 'running') }
function Job-View {
    $j = $S.job
    if (-not $j) { return @{ state = 'none' } }
    $items = @(foreach ($it in $j.items) { [ordered]@{ i = $it.i; rel = $it.rel; kind = $it.kind; size = $it.size; dst = $it.dst; status = $it.status; msg = $it.msg; sec = $it.sec } })
    $c = @{}; foreach ($it in $j.items) { $c[$it.status] = 1 + [int]$c[$it.status] }
    return [ordered]@{ state = $j.state; outDir = $j.outDir; current = $j.current; total = $j.items.Count; counts = $c; items = $items; error = $j.error }
}
function Begin-Job($items, [string]$outDir) {
    $S.stop = $false
    $S.job = @{ state = 'running'; outDir = $outDir; items = $items; current = -1; started = (Get-Date).ToString('s'); finished = $null; error = $null }
    Start-Worker
}

# ---------- HTTP ----------
$listener = $null
for ($p = $Port; $p -lt $Port + 10; $p++) {
    $l = New-Object System.Net.HttpListener; $l.Prefixes.Add("http://127.0.0.1:$p/")
    try { $l.Start(); $listener = $l; $Port = $p; break } catch { }
}
if (-not $listener) { Write-Host '도우미가 쓸 포트를 열지 못했습니다.'; exit 1 }
[void](New-Item -ItemType Directory -Force -Path $SESS_DIR)
# 지난번 도우미가 강제로 꺼져 남은 사본 정리(지금 살아 있는 도우미의 것은 그대로)
Get-ChildItem -LiteralPath $STAGE_ROOT -Directory -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -match '^\d+$' -and -not (Get-Process -Id ([int]$_.Name) -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -eq 'powershell' })
} | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
[IO.File]::WriteAllText($SESS_FILE, (@{ port = $Port; token = $TOKEN; pid = $PID; version = $VERSION } | ConvertTo-Json -Compress), [Text.Encoding]::UTF8)
$S.timeoutSec = $TimeoutSec
foreach ($a in $Allow.Split(';')) { if ($a -and (Test-Path -LiteralPath $a)) { [void]$S.allowed.Add($a) } }
Write-Host "PDF 변환 도우미 v$VERSION - http://127.0.0.1:$Port  (창을 닫으면 도우미가 꺼집니다)"
$SELF_ORIGIN = "http://127.0.0.1:$Port"   # PowerShell 변수는 대소문자를 구분하지 않으므로 이름을 겹치지 않게
$task = $null
try {
    while ($listener.IsListening) {
        if (-not $task) { $task = $listener.GetContextAsync() }
        if (-not $task.Wait(1000)) {
            Reap-Worker
            # 한 파일이 너무 오래 걸리면 도우미가 띄운 프로그램만 강제로 닫아 다음 파일로 넘어가게 한다
            if ($S.itemStarted -and $S.enginePid -and ((Get-Date) - $S.itemStarted).TotalSeconds -gt $TimeoutSec) {
                $S.timedOut = $true; $S.itemStarted = $null
                try { Stop-Process -Id $S.enginePid -Force -ErrorAction Stop } catch {}
            }
            # 프로그램을 띄우다 멈춘 경우: 그 사이 새로 뜬 같은 이름의 프로세스만 닫는다(원래 떠 있던 것은 그대로)
            $cr = $S.creating
            if ($cr -and ((Get-Date) - $cr.started).TotalSeconds -gt $TimeoutSec) {
                $S.timedOut = $true; $S.creating = $null
                Get-Process -Name $cr.pname -ErrorAction SilentlyContinue | Where-Object { $cr.before -notcontains $_.Id } | ForEach-Object { try { Stop-Process -Id $_.Id -Force } catch {} }
            }
            # 닫다가 멈춘 경우(15초): 도우미가 띄운 프로그램만 강제로 닫는다
            if ($S.closeStarted -and $S.enginePid -and ((Get-Date) - $S.closeStarted).TotalSeconds -gt 15) {
                $S.closeStarted = $null
                try { Stop-Process -Id $S.enginePid -Force -ErrorAction Stop } catch {}
            }
            if (-not (Job-Running) -and ((Get-Date) - $S.last).TotalMinutes -gt $IdleMinutes) { break }
            continue
        }
        $ctx = $task.Result; $task = $null; $S.last = Get-Date
        try {
            $req = $ctx.Request; $path = $req.Url.AbsolutePath
            if ($path -eq '/health') { Send-Json $ctx @{ ok = $true; version = $VERSION; pid = $PID }; continue }
            if ($path -eq '/' -or $path -eq '/index.html') {
                if ((Get-Q $req 't') -ne $TOKEN) { Send-Json $ctx @{ ok = $false; error = '실행 아이콘(01 PDF변환_실행)으로 다시 열어 주세요.' } 403; continue }
                Send-Html $ctx; continue
            }
            if (-not $path.StartsWith('/api/')) { Send-Json $ctx @{ ok = $false; error = 'NOT_FOUND' } 404; continue }
            $reqOrigin = $req.Headers['Origin']
            if (($reqOrigin -and $reqOrigin -ne $SELF_ORIGIN) -or $req.Headers['X-Token'] -ne $TOKEN -or $req.HttpMethod -ne 'POST') {
                Send-Json $ctx @{ ok = $false; error = 'FORBIDDEN' } 403; continue
            }
            $body = if ($path -eq '/api/upload') { $null } else { Read-Body $req }   # 받은 파일 본문은 아래에서 바로 디스크로
            switch ($path) {
                '/api/health' {
                    Send-Json $ctx @{ ok = $true; version = $VERSION; apps = (Get-Apps); hwpModule = (Get-HwpModule); downloads = $DOWNLOADS; running = (Job-Running) }
                }
                '/api/pick' {
                    if ($body.purpose -ne 'output' -and $body.mode -eq 'files') {
                        $files = @(); $skipped = 0
                        foreach ($p in (Pick-Files)) {
                            $f = Get-Item -LiteralPath $p -ErrorAction SilentlyContinue
                            if ($f -and (Is-Doc $f.Name)) {
                                [void]$S.allowedFiles.Add((Norm $f.FullName))
                                $files += [ordered]@{ path = $f.FullName; name = $f.Name; kind = $EXT[$f.Extension.ToLower()]; size = $f.Length }
                            } else { $skipped++ }
                        }
                        if ($files.Count -eq 0 -and $skipped -eq 0) { Send-Json $ctx @{ ok = $true; cancel = $true } } else { Send-Json $ctx @{ ok = $true; files = $files; skipped = $skipped } }
                        break
                    }
                    $desc = if ($body.purpose -eq 'output') { 'PDF를 저장할 폴더를 고르세요' } else { '변환할 문서(한글·워드·파워포인트)가 있는 폴더를 고르세요' }
                    $sel = Pick-Folder $desc $body.start
                    if ($sel) { [void]$S.allowed.Add($sel); Send-Json $ctx @{ ok = $true; path = $sel } } else { Send-Json $ctx @{ ok = $true; cancel = $true } }
                }
                '/api/scan' {
                    if (-not (Test-Allowed $body.folder)) { Send-Json $ctx @{ ok = $false; error = '선택 창으로 고른 폴더만 쓸 수 있습니다.' } 403; break }
                    $files = Get-Targets (Norm $body.folder) ([bool]$body.recursive)
                    $root = Norm $body.folder
                    $list = @(foreach ($f in $files) { [ordered]@{ rel = $f.FullName.Substring($root.Length).TrimStart('\'); path = $f.FullName; kind = $EXT[$f.Extension.ToLower()]; size = $f.Length } })
                    Send-Json $ctx @{ ok = $true; files = $list }
                }
                '/api/upload' {
                    # 화면에 끌어다 놓은 파일 하나를 받아 임시 폴더에 사본으로 저장(X-Rel: 놓은 폴더 기준 상대 경로, X-Group: 놓은 차례)
                    $rel = Safe-Rel ([uri]::UnescapeDataString([string]$req.Headers['X-Rel']))
                    $grp = [string]$req.Headers['X-Group']; if ($grp -notmatch '^\d{1,6}$') { $grp = '0' }
                    if (-not $rel -or -not (Is-Doc $rel)) { Send-Json $ctx @{ ok = $false; error = '한글·워드·파워포인트 문서만 받을 수 있습니다.' } 400; break }
                    if ($req.ContentLength64 -gt 2GB) { Send-Json $ctx @{ ok = $false; error = '2GB가 넘는 파일은 받을 수 없습니다.' } 413; break }
                    $dst = Norm (Join-Path (Join-Path $STAGE $grp) $rel)
                    if (-not (Is-Under $dst $STAGE)) { Send-Json $ctx @{ ok = $false; error = '파일 이름이 올바르지 않습니다.' } 400; break }
                    [void](New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dst))
                    $fs = [IO.File]::Create($dst)
                    try { $req.InputStream.CopyTo($fs) } finally { $fs.Close() }
                    Send-Json $ctx @{ ok = $true; path = $dst; size = (Get-Item -LiteralPath $dst).Length }
                }
                '/api/discard' {
                    # 목록 비우기: 끌어다 놓아 받은 사본을 지운다(변환 중이면 그대로)
                    if (Job-Running) { Send-Json $ctx @{ ok = $false; error = '변환 중에는 비울 수 없습니다.' } 409; break }
                    Clear-Stage; $S.job = $null; Send-Json $ctx @{ ok = $true }
                }
                '/api/start' {
                    if (Job-Running) { Send-Json $ctx @{ ok = $false; error = '이미 변환 중입니다.' } 409; break }
                    if ($body.retry) {
                        $prev = $S.job
                        if (-not $prev) { Send-Json $ctx @{ ok = $false; error = '다시 할 작업이 없습니다.' } 400; break }
                        $items = New-Object System.Collections.ArrayList
                        foreach ($it in $prev.items) { if ($it.status -eq '실패' -or $it.status -eq '취소') { $c = $it.Clone(); $c.status = '대기'; $c.msg = ''; $c.sec = $null; $c.i = $items.Count; [void]$items.Add($c) } }
                        if ($items.Count -eq 0) { Send-Json $ctx @{ ok = $false; error = '실패한 파일이 없습니다.' } 400; break }
                        Begin-Job $items $prev.outDir; Send-Json $ctx @{ ok = $true; outDir = $prev.outDir; total = $items.Count }; break
                    }
                    $mode = [string]$body.outMode
                    if (@('downloads', 'beside', 'folder') -notcontains $mode) { $mode = 'downloads' }
                    if ($mode -eq 'folder' -and -not (Test-Allowed $body.outFolder)) { Send-Json $ctx @{ ok = $false; error = '저장 폴더를 폴더 선택 창으로 골라 주세요.' } 403; break }
                    $list = @($body.items | Where-Object { $_ })
                    if ($list.Count -eq 0) { Send-Json $ctx @{ ok = $false; error = '변환할 문서가 없습니다.' } 400; break }
                    $bad = $null; $staged = $false
                    foreach ($x in $list) {
                        $p = [string]$x.path
                        if (-not (Test-Source $p) -or -not (Is-Doc $p) -or -not (Test-Path -LiteralPath $p -PathType Leaf)) { $bad = $p; break }
                        if (Is-Under (Norm $p) $STAGE) { $staged = $true }
                    }
                    if ($bad) { Send-Json $ctx @{ ok = $false; error = ('쓸 수 없는 파일이 있습니다: ' + [IO.Path]::GetFileName($bad) + ' (목록을 비우고 다시 골라 주세요)') } 403; break }
                    if ($mode -eq 'beside' -and $staged) { Send-Json $ctx @{ ok = $false; error = '끌어다 놓은 파일은 원본 옆에 저장할 수 없습니다. 저장 위치를 바꿔 주세요.' } 400; break }
                    $plan = Plan-Items $list $mode ([string]$body.outFolder) ([bool]$body.overwrite)
                    if ($plan.items.Count -eq 0) { Send-Json $ctx @{ ok = $false; error = '변환할 파일이 없습니다.' } 400; break }
                    $outDir = if ($mode -eq 'beside') { '' } else { $plan.outRoot }
                    Begin-Job $plan.items $outDir
                    Send-Json $ctx @{ ok = $true; outDir = $outDir; total = $plan.items.Count }
                }
                '/api/job' { Reap-Worker; Send-Json $ctx @{ ok = $true; job = (Job-View) } }
                '/api/stop' { $S.stop = $true; Send-Json $ctx @{ ok = $true } }
                '/api/open' {
                    # 결과 폴더 열기: 화면이 경로를 보내지 않고, 도우미가 아는 마지막 작업의 저장 위치만 연다
                    $target = $null; $j = $S.job
                    if ($j) {
                        if ($j.outDir) { $target = $j.outDir }
                        else { $first = @($j.items | Where-Object { $_.status -eq '완료' -or $_.status -eq '건너뜀' }) | Select-Object -First 1; if ($first) { $target = $first.dstDir } }
                    }
                    if (-not $target -or -not (Test-Path -LiteralPath $target)) { Send-Json $ctx @{ ok = $false; error = '열 결과 폴더가 없습니다.' } 404; break }
                    Start-Process explorer.exe -ArgumentList ('"' + $target + '"'); Send-Json $ctx @{ ok = $true }
                }
                '/api/quit' { Send-Json $ctx @{ ok = $true }; if (-not (Job-Running)) { $listener.Stop() } }
                default { Send-Json $ctx @{ ok = $false; error = 'NOT_FOUND' } 404 }
            }
        } catch { try { Send-Json $ctx @{ ok = $false; error = $_.Exception.Message } 500 } catch {} }
    }
} finally {
    try { $listener.Close() } catch {}
    try { Remove-Item -LiteralPath $SESS_FILE -Force -ErrorAction SilentlyContinue } catch {}
    try { Clear-Stage } catch {}
}
