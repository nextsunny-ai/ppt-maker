# HTML -> PPT  앱 런처: 로컬 서버 기동 + 테두리 없는 앱 창
$ErrorActionPreference='SilentlyContinue'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$port = 39217

function Test-Server { try { (Invoke-WebRequest "http://127.0.0.1:$port/ping" -TimeoutSec 1 -UseBasicParsing).StatusCode -eq 200 } catch { $false } }
function Find-Chrome {
  $c=@("$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
       "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
       "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe",
       "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
       "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe")
  foreach($p in $c){ if(Test-Path $p){ return $p } }; return $null
}

$node = (Get-Command node -ErrorAction SilentlyContinue).Source
if (-not $node) {
  Add-Type -AssemblyName System.Windows.Forms | Out-Null
  [System.Windows.Forms.MessageBox]::Show("Node.js가 필요합니다.`nhttps://nodejs.org 에서 설치 후 다시 실행하세요.","HTML → PPT") | Out-Null
  exit
}

if (-not (Test-Server)) {
  Start-Process -WindowStyle Hidden -FilePath $node -ArgumentList "`"$dir\server.js`"" -WorkingDirectory $dir
  for ($i=0; $i -lt 40; $i++) { Start-Sleep -Milliseconds 300; if (Test-Server) { break } }
}

$url = "http://127.0.0.1:$port/"
$fileArgs = @($args | Where-Object { $_ -and ($_ -match '\.html?$') -and (Test-Path $_) })
if ($fileArgs.Count -gt 0) {
  $q = ($fileArgs | ForEach-Object { 'file=' + [uri]::EscapeDataString((Resolve-Path $_).Path) }) -join '&'
  $url += "?$q"
}

$chrome = Find-Chrome
if ($chrome) {
  $profile = "$env:LOCALAPPDATA\HTML2PPT\chrome-profile"
  New-Item -ItemType Directory -Force $profile | Out-Null
  Start-Process $chrome -ArgumentList "--app=$url","--window-size=880,820","--user-data-dir=`"$profile`"","--no-first-run","--no-default-browser-check"
} else {
  Start-Process $url   # 폴백: 기본 브라우저
}
