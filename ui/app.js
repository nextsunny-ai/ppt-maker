const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const api = async (p, body) => {
  const o = body ? {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)} : {};
  return (await fetch(p,o)).json();
};
const baseName = p => p.replace(/^.*[\\/]/,'');
const state = { files:[], fonts:[] };

/* ---------- nav ---------- */
function showView(v){
  $$('.nav').forEach(b=>b.classList.toggle('active', b.dataset.view===v));
  $$('.view').forEach(s=>s.classList.toggle('active', s.dataset.view===v));
}
$$('.nav').forEach(b=>b.onclick=()=>showView(b.dataset.view));
$$('[data-view]').forEach(el=>{ if(!el.classList.contains('nav')&&!el.classList.contains('view')) el.onclick=()=>showView(el.dataset.view); });

/* ---------- theme ---------- */
function setTheme(t){
  document.documentElement.dataset.theme=t;
  localStorage.setItem('ppt_theme',t);
  $$('#themeSeg button').forEach(b=>b.classList.toggle('active', b.dataset.t===t));
}
$('#themeBtn').onclick=()=>setTheme(document.documentElement.dataset.theme==='dark'?'light':'dark');
$$('#themeSeg button').forEach(b=>b.onclick=()=>setTheme(b.dataset.t));
setTheme(localStorage.getItem('ppt_theme')||'dark');

/* ---------- embed sync ---------- */
$('#embedDefault').onchange=()=>{ $('#embedChk').checked=$('#embedDefault').checked; };

/* ---------- files ---------- */
async function pick(){ const {paths}=await api('/pick'); addFiles(paths||[]); }
function addFiles(paths){
  for(const p of paths){ if(p&&!state.files.includes(p)) state.files.push(p); }
  renderFiles();
  if(state.files.length) analyze();
}
function renderFiles(){
  const el=$('#filelist');
  const has=state.files.length>0;
  el.hidden=!has; $('#optbar').hidden=!has; $('#fontMini').hidden=!has;
  if(!has){ $('#fontrows').innerHTML='<div class="empty">먼저 [변환] 탭에서 HTML을 선택하세요.</div>'; $('#installAll').hidden=true; $('#fontNote').textContent=''; return; }
  el.innerHTML=state.files.map((f,i)=>`<span class="chip"><span class="nm" title="${f}">${baseName(f)}</span><span class="x" data-i="${i}">✕</span></span>`).join('');
  el.querySelectorAll('.x').forEach(x=>x.onclick=()=>{ state.files.splice(+x.dataset.i,1); renderFiles(); if(state.files.length) analyze(); });
}

/* ---------- fonts ---------- */
async function analyze(){
  // 즉시 작업중 표시
  const btn=$('#convertBtn'); btn.disabled=true; btn.classList.remove('ready');
  $('#results').hidden=true; $('#progress').hidden=false;
  $('#progressText').textContent='폰트 확인 중… (몇 초 걸립니다)';
  $('#barFill').style.width='10%';
  let pct=10; const tk=setInterval(()=>{pct=Math.min(pct+5,55);$('#barFill').style.width=pct+'%';},400);
  $('#fontMiniRows').innerHTML='<span class="fpill"><span class="spin"></span>&nbsp;분석 중…</span>';
  $('#fontrows').innerHTML='<div class="empty"><span class="spin"></span>&nbsp;폰트 분석 중…</div>';
  const merged=new Map();
  for(const f of state.files){
    const {fonts=[]}=await api('/analyze',{path:f});
    fonts.forEach(ft=>{ const c=merged.get(ft.name); if(!c||(c.installed&&!ft.installed)) merged.set(ft.name,ft); });
  }
  state.fonts=[...merged.values()];
  clearInterval(tk); $('#barFill').style.width='100%';
  setTimeout(()=>{ $('#progress').hidden=true; },350);
  renderFonts();
  btn.disabled=false; btn.classList.add('ready');
}
function renderFonts(){
  const mini=$('#fontMiniRows');
  mini.innerHTML = state.fonts.length ? state.fonts.map(f=>`<span class="fpill ${f.installed?'ok':'no'}"><span class="pip"></span>${f.name}${f.installed?'':' · 없음'}</span>`).join('')
    : '<span class="fpill ok"><span class="pip"></span>기본 폰트만 사용</span>';
  const rows=$('#fontrows');
  rows.innerHTML = state.fonts.length ? state.fonts.map(f=>{
    const st=f.installed?'<span class="status ok"><span class="pip"></span>설치됨</span>':'<span class="status no"><span class="pip"></span>없음</span>';
    const act=(!f.installed&&f.downloadable)?`<button class="mini" data-font="${f.name}">설치</button>`:(!f.installed?'<span class="note" style="margin:0">수동 설치</span>':'');
    return `<div class="fontrow"><span class="fn">${f.name}</span><div class="fr">${st}${act}</div></div>`;
  }).join('') : '<div class="empty">감지된 커스텀 폰트 없음 (기본 폰트만 사용)</div>';
  rows.querySelectorAll('.mini').forEach(b=>b.onclick=()=>installOne(b.dataset.font,b));
  const missing=state.fonts.filter(f=>!f.installed);
  $('#installAll').hidden=!missing.some(f=>f.downloadable);
  const manual=missing.filter(f=>!f.downloadable).map(f=>f.name);
  $('#fontNote').textContent = manual.length
    ? `수동 설치 필요: ${manual.join(', ')} — 폰트 파일을 받아 설치하세요. (나머지는 변환 시 임베드)`
    : (missing.length? '없는 폰트는 무료로 설치할 수 있습니다. 변환하면 PPT에 임베드되어 다른 PC에서도 동일합니다.'
                     : '필요한 폰트가 모두 설치되어 있습니다. 변환 시 PPT에 임베드됩니다.');
}
async function installOne(name,btn){ if(btn){btn.disabled=true;btn.innerHTML='<span class="spin"></span>';} await api('/install',{names:[name]}); await analyze(); }
$('#installAll').onclick=async()=>{
  const names=state.fonts.filter(f=>!f.installed&&f.downloadable).map(f=>f.name); if(!names.length)return;
  const b=$('#installAll'); b.disabled=true; b.innerHTML='<span class="spin"></span> 설치 중…';
  await api('/install',{names}); b.disabled=false; b.textContent='없는 폰트 모두 설치'; await analyze();
};

/* ---------- convert ---------- */
$('#convertBtn').onclick=async()=>{
  if(!state.files.length) return;
  const btn=$('#convertBtn'); btn.disabled=true; btn.classList.remove('ready');
  $('#progress').hidden=false; $('#results').hidden=true;
  const fill=$('#barFill'), txt=$('#progressText');
  txt.textContent='슬라이드 렌더링 중… (12장 기준 15~25초)'; fill.style.width='15%';
  let pct=15; const tick=setInterval(()=>{pct=Math.min(pct+4,88);fill.style.width=pct+'%';},600);
  const embed=$('#embedChk').checked;
  const {outputs=[]}=await api('/convert',{paths:state.files,embed});
  clearInterval(tick); fill.style.width='100%'; txt.textContent=embed?'폰트 임베드 완료':'변환 완료';
  setTimeout(()=>{ $('#progress').hidden=true; showResults(outputs); btn.disabled=false; },500);
};
function showResults(outs){
  $('#results').hidden=false;
  $('#resultrows').innerHTML = outs.length? outs.map(o=>`<div class="resultrow"><div class="rl"><span class="ic">P</span><span class="rn" title="${o}">${baseName(o)}</span></div><button class="btn ghost sm" data-open="${o}">폴더에서 보기</button></div>`).join('')
    : '<div class="empty">생성된 파일이 없습니다.</div>';
  $('#resultrows').querySelectorAll('[data-open]').forEach(b=>b.onclick=()=>api('/reveal',{path:b.dataset.open}));
}

/* ---------- drop + launch arg ---------- */
$('#drop').onclick=pick;
$('#drop').addEventListener('dragover',e=>{e.preventDefault();$('#drop').classList.add('drag');});
$('#drop').addEventListener('dragleave',()=>$('#drop').classList.remove('drag'));
$('#drop').addEventListener('drop',e=>{
  e.preventDefault();$('#drop').classList.remove('drag');
  // 브라우저 보안상 창 안 드롭은 파일 경로를 못 읽음 → 안내
  const h=$('#dropHint'); const orig=h.innerHTML;
  h.innerHTML='<b>여기를 클릭</b>해서 HTML 파일을 고르세요';
  h.style.color='var(--accent-2)';
  setTimeout(()=>{ h.innerHTML=orig; h.style.color=''; }, 3500);
});
(function(){ const q=new URLSearchParams(location.search).getAll('file'); if(q.length) addFiles(q); })();
