# HTML -> PPT 변환 런처 (드래그&드롭 또는 파일 선택)
$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "============================================" -ForegroundColor DarkGray
Write-Host "   PPT 메이커  ·  HTML -> PPT  (c) 2026 Sunny Ryu" -ForegroundColor White
Write-Host "============================================`n" -ForegroundColor DarkGray

# 1) 입력 파일 수집 (드롭된 인자 우선, 없으면 파일 선택창)
$files = @($args | Where-Object { $_ -and (Test-Path $_) -and ($_ -match '\.html?$') })
if (-not $files -or $files.Count -eq 0) {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'HTML 파일 (*.html;*.htm)|*.html;*.htm'
    $dlg.Multiselect = $true
    $dlg.Title = 'PPT로 변환할 HTML 파일 선택 (여러 개 가능)'
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host '취소되었습니다.' -ForegroundColor Yellow; Start-Sleep 1; exit
    }
    $files = $dlg.FileNames
}

# 2) Node 확인
$node = (Get-Command node -ErrorAction SilentlyContinue).Source
if (-not $node) {
    Write-Host 'Node.js가 설치돼 있지 않습니다. https://nodejs.org 에서 설치 후 다시 실행하세요.' -ForegroundColor Red
    Write-Host "`n아무 키나 누르면 닫힙니다..."; [void][System.Console]::ReadKey($true); exit
}

# 3) 변환
$outs = @()
foreach ($f in $files) {
    Write-Host "● 변환 중: $(Split-Path $f -Leaf)" -ForegroundColor Cyan
    & $node (Join-Path $dir 'html2pptx.js') $f
    $out = [System.IO.Path]::ChangeExtension($f, '.pptx')
    if (Test-Path $out) { $outs += $out }
    else { Write-Host "  ✗ 변환 실패: $(Split-Path $f -Leaf)" -ForegroundColor Red }
    Write-Host ''
}

# 4) 폰트 임베드 (PowerPoint 있을 때만)
if ($outs.Count -gt 0) {
    try {
        $ppt = New-Object -ComObject PowerPoint.Application
        Write-Host '폰트 임베드 중...' -ForegroundColor DarkCyan
        foreach ($o in $outs) {
            try { $p = $ppt.Presentations.Open($o, $false, $false, $false); $p.SaveAs($o, 24, -1); $p.Close() } catch {}
        }
        $ppt.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
        Write-Host '폰트 임베드 완료.' -ForegroundColor DarkCyan
    } catch {
        Write-Host '(PowerPoint이 없어 폰트 임베드는 건너뜀 — PPT는 정상 생성됨)' -ForegroundColor Yellow
    }
}

# 5) 결과 안내
Write-Host ''
if ($outs.Count -gt 0) {
    Write-Host "총 $($outs.Count)개 PPT 생성 완료:" -ForegroundColor Green
    $outs | ForEach-Object { Write-Host "   $_" -ForegroundColor Green }
    Start-Process explorer.exe ('/select,"' + $outs[0] + '"')
} else {
    Write-Host '생성된 PPT가 없습니다.' -ForegroundColor Red
}
Write-Host "`n아무 키나 누르면 닫힙니다..."
[void][System.Console]::ReadKey($true)
