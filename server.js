// Local backend for PPT MAKER. Non-blocking: all heavy work runs in child processes.
const http = require('http');
const fs = require('fs');
const path = require('path');
const { execFile, spawn } = require('child_process');
const { promisify } = require('util');
const pexec = promisify(execFile);

const DIR = __dirname;
const PORT = 39217;
const NODE = process.execPath;

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

async function embedFonts(list){
  if(!list.length) return;
  const arr=list.map(p=>`'${p.replace(/'/g,"''")}'`).join(',');
  const ps=`try{ $a=New-Object -ComObject PowerPoint.Application; foreach($f in @(${arr})){ try{ $p=$a.Presentations.Open($f,$false,$false,$false); $p.SaveAs($f,24,-1); $p.Close() }catch{} }; $a.Quit() }catch{}`;
  try{ await pexec('powershell',['-NoProfile','-Command',ps],{maxBuffer:1<<20}); }catch(e){}
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
      const {paths,embed}=await readBody(req);
      const outputs=[];
      for(const hp of (paths||[])){
        try{ await pexec(NODE,[path.join(DIR,'html2pptx.js'),hp],{maxBuffer:1<<22}); const out=hp.replace(/\.[^.]+$/,'')+'.pptx'; if(fs.existsSync(out)) outputs.push(out); }
        catch(e){}
      }
      if(embed!==false) await embedFonts(outputs);
      return send(res,200,{outputs});
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
