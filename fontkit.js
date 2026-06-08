// Font detection + per-user install (no AI, no admin). Windows.
const puppeteer = require('puppeteer-core');
const { execFileSync, execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

function findChrome(){
  const c=[process.env.CHROME_PATH,
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
    (process.env.LOCALAPPDATA||'')+'\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe'].filter(Boolean);
  for(const p of c){ try{ if(fs.existsSync(p)) return p; }catch(e){} }
  return null;
}

// canonical family -> google-webfonts-helper slug + variants (for download)
const KNOWN = {
  'Inter':          {slug:'inter',           variants:'regular,500,600,700,800,900'},
  'Pretendard':     {slug:null},  // not on gwfh; handled via jsDelivr below
  'Spectral':       {slug:'spectral',        variants:'300,300italic,regular,italic,500'},
  'Cormorant Garamond':{slug:'cormorant-garamond', variants:'300,300italic,regular,italic,500,600,700'},
  'JetBrains Mono': {slug:'jetbrains-mono',   variants:'regular,500'},
  'Playfair Display':{slug:'playfair-display',variants:'regular,500,600,700'},
  'Poppins':        {slug:'poppins',         variants:'regular,500,600,700'},
  'Montserrat':     {slug:'montserrat',      variants:'regular,500,600,700'},
  'Noto Sans KR':   {slug:'noto-sans-kr',    variants:'regular,500,700'},
  'Roboto':         {slug:'roboto',          variants:'regular,500,700'},
  'Lato':           {slug:'lato',            variants:'regular,700'},
};

// 폰트 스택을 순서대로(=브라우저 우선순위) 보고 첫 '실제' 폰트를 정규화해 반환.
// (예전 버그: Pretendard를 먼저 검사해 "Inter, Pretendard…"를 Inter가 아닌 Pretendard로 오인)
function canonical(ff){
  ff=ff||'';
  const toks=ff.split(',').map(t=>t.replace(/['"]/g,'').trim()).filter(Boolean);
  const generic=/^(system-ui|ui-sans-serif|ui-serif|ui-monospace|sans-serif|serif|monospace|-apple-system|BlinkMacSystemFont|Segoe UI|Segoe|Apple SD Gothic Neo|Malgun Gothic|Arial|Helvetica( Neue)?|Georgia|Times( New Roman)?|SF Pro( Display| Text)?|SF Mono)$/i;
  for(const t of toks){
    if(generic.test(t)) continue;            // 시스템/대체용 generic은 건너뜀
    if(/Pretendard/i.test(t)) return 'Pretendard';
    if(/JetBrains/i.test(t)) return 'JetBrains Mono';
    if(/Spectral/i.test(t)) return 'Spectral';
    if(/Cormorant/i.test(t)) return 'Cormorant Garamond';
    if(/^Inter\b/i.test(t)) return 'Inter';
    if(/Playfair/i.test(t)) return 'Playfair Display';
    if(/Poppins/i.test(t)) return 'Poppins';
    if(/Montserrat/i.test(t)) return 'Montserrat';
    if(/Noto Sans KR/i.test(t)) return 'Noto Sans KR';
    if(/Noto Serif KR/i.test(t)) return 'Noto Serif KR';
    if(/IBM Plex Mono/i.test(t)) return 'IBM Plex Mono';
    if(/Roboto/i.test(t)) return 'Roboto';
    if(/Lato/i.test(t)) return 'Lato';
    return t;                                 // 알 수 없는 named 폰트 = 그대로
  }
  return null;
}

async function detectUsedFonts(htmlPath){
  const chrome=findChrome();
  if(!chrome) throw new Error('Chrome/Edge not found');
  const browser=await puppeteer.launch({executablePath:chrome, headless:'new', args:['--no-sandbox']});
  try{
    const page=await browser.newPage();
    await page.setViewport({width:1920,height:1080});
    await page.goto('file:///'+htmlPath.replace(/\\/g,'/'),{waitUntil:'networkidle0'});
    await page.evaluate(()=>{const d=document.querySelector('deck-stage');if(d)d.setAttribute('noscale','');});
    await page.emulateMediaType('print');
    await new Promise(r=>setTimeout(r,800));
    const fams=await page.evaluate(()=>{
      const set=new Set();
      document.querySelectorAll('section *').forEach(el=>{
        if([...el.childNodes].some(n=>n.nodeType===3&&n.textContent.trim()))
          set.add(getComputedStyle(el).fontFamily);
      });
      return [...set];
    });
    const out=new Map();
    fams.forEach(f=>{ const c=canonical(f); if(c) out.set(c,true); });
    return [...out.keys()];
  } finally { await browser.close(); }
}

function installedFamilies(){
  const ps = `Add-Type -AssemblyName System.Drawing; (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name -join "\`n"`;
  try{
    const out=execFileSync('powershell',['-NoProfile','-Command',ps],{encoding:'utf8',maxBuffer:1024*1024*8});
    return out.split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
  }catch(e){ return []; }
}

function isInstalled(family, installed){
  const f=family.toLowerCase();
  // exact-ish: a family is installed if some installed name equals it or starts with it + space (weight variants)
  return installed.some(n=>{ const x=n.toLowerCase(); return x===f || x.startsWith(f+' '); });
}

// returns [{name, installed, downloadable}]
function analyzeFonts(used){
  const installed=installedFamilies();
  return used.map(name=>({
    name,
    installed: isInstalled(name, installed),
    downloadable: (name==='Pretendard') || !!(KNOWN[name] && KNOWN[name].slug)
  }));
}

function dlFile(url, dest){
  execFileSync('curl',['-sL','-o',dest,url],{stdio:'ignore'});
  return fs.existsSync(dest) && fs.statSync(dest).size>1000;
}

// download + per-user install a known font family. returns true on success
function installFont(name){
  const tmp=path.join(os.tmpdir(),'h2p_fonts_'+Date.now());
  fs.mkdirSync(tmp,{recursive:true});
  const ttfs=[];
  if(name==='Pretendard'){
    // jsDelivr static weights
    const base='https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/public/static/';
    ['Pretendard-Regular.otf','Pretendard-Medium.otf','Pretendard-SemiBold.otf','Pretendard-Bold.otf','Pretendard-ExtraBold.otf','Pretendard-Black.otf'].forEach(fn=>{
      const d=path.join(tmp,fn); if(dlFile(base+fn,d)) ttfs.push(d);
    });
  } else if(KNOWN[name] && KNOWN[name].slug){
    const k=KNOWN[name];
    const zip=path.join(tmp,'f.zip');
    const url=`https://gwfh.mranftl.com/api/fonts/${k.slug}?download=zip&subsets=latin&variants=${k.variants}&formats=ttf`;
    if(dlFile(url,zip)){
      try{
        execFileSync('powershell',['-NoProfile','-Command',`Expand-Archive -LiteralPath '${zip}' -DestinationPath '${tmp}\\ex' -Force`],{stdio:'ignore'});
        const exdir=path.join(tmp,'ex');
        for(const fn of fs.readdirSync(exdir)) if(/\.(ttf|otf)$/i.test(fn)) ttfs.push(path.join(exdir,fn));
      }catch(e){}
    }
  }
  if(!ttfs.length) return false;
  // per-user install via PowerShell (copy + registry + AddFontResource)
  const list=ttfs.map(t=>t.replace(/'/g,"''")).map(t=>`'${t}'`).join(',');
  const ps=`
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System.Runtime.InteropServices;
public class GdiF { [DllImport("gdi32.dll")] public static extern int AddFontResource(string f); }
"@
$dest="$env:LOCALAPPDATA\\Microsoft\\Windows\\Fonts"; New-Item -ItemType Directory -Force $dest | Out-Null
$reg='HKCU:\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Fonts'
foreach($src in @(${list})){
  $fn=Split-Path $src -Leaf; $target=Join-Path $dest $fn
  Copy-Item $src $target -Force
  $base=[System.IO.Path]::GetFileNameWithoutExtension($fn)
  Set-ItemProperty -Path $reg -Name "$base (TrueType)" -Value $target
  [GdiF]::AddFontResource($target) | Out-Null
}
`;
  try{ execFileSync('powershell',['-NoProfile','-Command',ps],{stdio:'ignore'}); return true; }
  catch(e){ return false; }
}

module.exports = { detectUsedFonts, analyzeFonts, installFont, installedFamilies };
