// Local backend for PPT MAKER. Non-blocking: all heavy work runs in child processes.
const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFile, spawn } = require('child_process');
const { promisify } = require('util');
const pexec = promisify(execFile);

const DIR = __dirname;
const PORT = 39217;
const NODE = process.execPath;
const THUMB_DIR = path.join(os.tmpdir(), 'pptmaker_preview');   // 미리보기 썸네일 임시 저장

function send(res, code, body, type='application/json'){
  res.writeHead(code, {'Content-Type':type, 'Cache-Control':'no-store', 'Connection':'close'});
  res.end(typeof body==='string'||Buffer.isBuffer(body) ? body : JSON.stringify(body));
}
function readBody(req){ return new Promise(r=>{ let d=''; req.on('data',c=>d+=c); req.on('end',()=>{ try{r(JSON.parse(d||'{}'))}catch(e){r({})} }); }); }

async function pickFiles(){
  const ps = `[Console]::OutputEncoding=[System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Windows.Forms | Out-Null
$owner = New-Object System.Windows.Forms.Form
$owner.TopMost=$true; $owner.ShowInTaskbar=$false; $owner.Opacity=0; $owner.Size=New-Object System.Drawing.Size(1,1)
$owner.StartPosition='CenterScreen'; $owner.Show(); $owner.Activate(); $owner.BringToFront()
$d = New-Object System.Windows.Forms.OpenFileDialog
$d.Filter = 'HTML (*.html;*.htm)|*.html;*.htm'; $d.Multiselect = $true; $d.Title = 'PPT로 변환할 HTML 선택'
$r = $d.ShowDialog($owner); $owner.Close()
if ($r -eq [System.Windows.Forms.DialogResult]::OK) { $d.FileNames -join "\`n" }`;
  try{ const {stdout}=await pexec('powershell',['-NoProfile','-STA','-Command',ps],{encoding:'utf8',maxBuffer:1<<20}); return stdout.split(/\r?\n/).map(s=>s.trim()).filter(Boolean); }
  catch(e){ return []; }
}

function clearThumbs(){
  try{
    if(fs.existsSync(THUMB_DIR)){ for(const f of fs.readdirSync(THUMB_DIR)){ try{ fs.unlinkSync(path.join(THUMB_DIR,f)); }catch(e){} } }
    else fs.mkdirSync(THUMB_DIR,{recursive:true});
  }catch(e){}
}
// PowerPoint 한 세션에서: (옵션)폰트 임베드 + 각 슬라이드 PNG 썸네일 생성
// 반환: [{file, n, w, h, thumbs:[절대경로...]}]
async function finalize(list, embed){
  if(!list.length) return [];
  const arr=list.map(p=>`'${p.replace(/'/g,"''")}'`).join(',');
  const td=THUMB_DIR.replace(/'/g,"''");
  const emb=(embed!==false)?'$true':'$false';
  const ps=`
$ErrorActionPreference='SilentlyContinue'
$files=@(${arr}); $td='${td}'; $embed=${emb}
if(!(Test-Path $td)){ New-Item -ItemType Directory -Path $td -Force | Out-Null }
$out=@()
try{
  $a=New-Object -ComObject PowerPoint.Application
  for($fi=0; $fi -lt $files.Count; $fi++){
    $f=$files[$fi]
    try{
      $p=$a.Presentations.Open($f,$false,$false,$false)
      if($embed){ $p.SaveAs($f,24,-1) }
      $w=$p.PageSetup.SlideWidth; $h=$p.PageSetup.SlideHeight
      $tw=800; $th=[int][math]::Round($tw*$h/$w)
      $thumbs=@()
      for($si=1; $si -le $p.Slides.Count; $si++){
        $png=Join-Path $td ("f{0}_s{1}.png" -f $fi,$si)
        $p.Slides.Item($si).Export($png,'PNG',$tw,$th) | Out-Null
        $thumbs+=$png
      }
      $out += [pscustomobject]@{ file=$f; n=$p.Slides.Count; w=[int]$w; h=[int]$h; thumbs=$thumbs }
      $p.Close()
    }catch{}
  }
  $a.Quit()
}catch{}
@{files=@($out)} | ConvertTo-Json -Depth 5 -Compress
`;
  try{ const {stdout}=await pexec('powershell',['-NoProfile','-Command',ps],{maxBuffer:1<<24});
    const j=JSON.parse(stdout.trim()); return j.files||[]; }
  catch(e){ return []; }
}

const MIME={'.html':'text/html;charset=utf-8','.css':'text/css','.js':'text/javascript','.svg':'image/svg+xml','.png':'image/png'};

const server = http.createServer(async (req,res)=>{
  const u=new URL(req.url,'http://localhost');
  try{
    if(u.pathname==='/'){ return send(res,200,fs.readFileSync(path.join(DIR,'ui','index.html')),'text/html;charset=utf-8'); }
    if(u.pathname.startsWith('/ui/')){
      const fp=path.join(DIR, u.pathname.replace(/\.\./g,''));
      if(fs.existsSync(fp)) return send(res,200,fs.readFileSync(fp), MIME[path.extname(fp)]||'application/octet-stream');
      return send(res,404,'not found','text/plain');
    }
    if(u.pathname==='/favicon.ico'){ res.writeHead(204); return res.end(); }
    if(u.pathname==='/thumb'){
      const p=path.normalize(u.searchParams.get('p')||'');
      if(p.startsWith(THUMB_DIR) && fs.existsSync(p)) return send(res,200,fs.readFileSync(p),'image/png');
      return send(res,404,'no','text/plain');
    }
    if(u.pathname==='/ping'){ return send(res,200,{ok:true}); }
    if(u.pathname==='/pick'){ return send(res,200,{paths:await pickFiles()}); }
    if(u.pathname==='/analyze' && req.method==='POST'){
      const {path:hp}=await readBody(req);
      if(!hp||!fs.existsSync(hp)) return send(res,400,{error:'파일 없음'});
      try{ const {stdout}=await pexec(NODE,[path.join(DIR,'fontcli.js'),'analyze',hp],{maxBuffer:1<<20}); return send(res,200,JSON.parse(stdout)); }
      catch(e){ return send(res,200,{fonts:[]}); }
    }
    if(u.pathname==='/install' && req.method==='POST'){
      const {names}=await readBody(req);
      const results=[];
      for(const n of (names||[])){
        try{ const {stdout}=await pexec(NODE,[path.join(DIR,'fontcli.js'),'install',n],{maxBuffer:1<<20}); const j=JSON.parse(stdout); results.push({name:n,ok:!!j.ok}); }
        catch(e){ results.push({name:n,ok:false}); }
      }
      return send(res,200,{results});
    }
    if(u.pathname==='/convert' && req.method==='POST'){
      const {paths,embed,orient}=await readBody(req);
      const or = (orient==='portrait'||orient==='세로') ? 'portrait' : 'landscape';
      clearThumbs();
      const outputs=[];
      for(const hp of (paths||[])){
        try{ const out=hp.replace(/\.[^.]+$/,'')+'.pptx'; await pexec(NODE,[path.join(DIR,'html2pptx.js'),hp,out,or],{maxBuffer:1<<22}); if(fs.existsSync(out)) outputs.push(out); }
        catch(e){}
      }
      // 폰트 임베드 + 미리보기 썸네일 (PowerPoint 한 세션)
      const fin = await finalize(outputs, embed);
      const previews = fin.map(o=>({ out:o.file, n:o.n, portrait:(o.h>o.w),
        thumbs:(o.thumbs||[]).map(t=>'/thumb?p='+encodeURIComponent(t)) }));
      return send(res,200,{outputs, previews});
    }
    if(u.pathname==='/reveal' && req.method==='POST'){
      const {path:p}=await readBody(req);
      try{
        if(p && fs.existsSync(p)){
          // 한글·공백 경로 안전: 따옴표로 감싸고 verbatim 전달 → 폴더 열고 파일 선택
          spawn('explorer.exe', ['/select,"'+p+'"'], {windowsVerbatimArguments:true, detached:true}).unref();
        } else if(p){
          spawn('explorer.exe', [path.dirname(p)], {detached:true}).unref();
        }
      }catch(e){}
      return send(res,200,{ok:true});
    }
    if(u.pathname==='/quit'){ send(res,200,{ok:true}); setTimeout(()=>process.exit(0),200); return; }
    return send(res,404,{error:'not found'});
  }catch(e){ return send(res,500,{error:String(e.message||e)}); }
});

server.keepAliveTimeout = 2000;
server.headersTimeout = 5000;
server.on('error', e=>{ if(e.code==='EADDRINUSE'){ console.error('이미 실행 중입니다.'); process.exit(0);} else throw e; });
server.listen(PORT, '127.0.0.1', ()=>console.log('PPT MAKER server on http://127.0.0.1:'+PORT));
