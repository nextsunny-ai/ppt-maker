/* PPT 메이커 웹 — 브라우저에서 HTML 덱 → 편집가능 .pptx. 데스크탑 엔진과 동일 추출(extract.js, iframe). */
const IN = 13.333/1920, PT = IN*72;
const DESKTOP_URL = 'https://github.com/nextsunny-ai/ppt-maker';
const $ = s=>document.querySelector(s);
const state = { files:[], orient:'landscape' };

/* ---------- 폰트 매핑 (데스크탑 pptFont와 동일) ---------- */
function firstFamily(fam){
  for(let p of String(fam||'').split(',')){
    p=p.trim().replace(/^["']|["']$/g,'');
    if(!p) continue;
    if(/^(serif|sans-serif|monospace|cursive|fantasy|system-ui|ui-(serif|sans-serif|monospace|rounded)|inherit|initial|unset|-apple-system|BlinkMacSystemFont|Segoe UI|Apple SD|Malgun|Arial|Helvetica|sans|Roboto)$/i.test(p)) continue;
    return p;
  }
  return '';
}
function pptFont(fam, w){
  fam=fam||''; w=w||400;
  const pick=(base,steps)=>{ let s=''; for(const [mw,suf] of steps){ if(w>=mw) s=suf; } return {face:base+s, bold:false}; };
  if(/Pretendard/i.test(fam)) return pick('Pretendard',[[0,''],[500,' Medium'],[600,' SemiBold'],[800,' ExtraBold'],[900,' Black']]);
  if(/JetBrains|Plex Mono|monospace/i.test(fam)) return pick('JetBrains Mono',[[0,''],[500,' Medium']]);
  if(/Cormorant/i.test(fam)) return {face:'Cormorant Garamond', bold:false};
  if(/Spectral/i.test(fam)) return pick('Spectral',[[0,' Light'],[400,''],[500,' Medium']]);
  if(/(^|[",\s])Inter([",\s]|$)/i.test(fam)) return pick('Inter',[[0,''],[500,' Medium'],[600,' SemiBold'],[800,' ExtraBold'],[900,' Black']]);
  const real=firstFamily(fam);
  if(real) return {face:real, bold:w>=600};
  return pick('Inter',[[0,''],[500,' Medium'],[600,' SemiBold'],[800,' ExtraBold'],[900,' Black']]);
}

/* ---------- 배경이미지 배치 (데스크탑 resolveBg와 동일) ---------- */
function resolveBg(a, nat, sizeStr, pxStr, pyStr){
  sizeStr=(sizeStr||'auto').trim(); let w,h; const ar=nat.w/nat.h;
  if(sizeStr==='cover'||sizeStr==='contain'){
    const aar=a.w/a.h;
    if((sizeStr==='cover')===(ar>aar)){ h=a.h; w=a.h*ar; } else { w=a.w; h=a.w/ar; }
  } else {
    const pt=sizeStr.split(/\s+/); const sx=pt[0]||'auto', sy=pt[1]||'auto';
    const toPx=(v,base)=> (v==='auto'||v==null)?null : (/%$/.test(v)? parseFloat(v)/100*base : parseFloat(v));
    let wpx=toPx(sx,a.w), hpx=toPx(sy,a.h);
    if(wpx==null&&hpx==null){ w=nat.w; h=nat.h; }
    else if(wpx==null){ h=hpx; w=hpx*ar; }
    else if(hpx==null){ w=wpx; h=wpx/ar; }
    else { w=wpx; h=hpx; }
  }
  const pos=(v,base,size)=>{ v=(v||'').trim();
    if(/%$/.test(v)) return parseFloat(v)/100*(base-size);
    if(v==='left'||v==='top') return 0;
    if(v==='right'||v==='bottom') return base-size;
    if(v==='center') return (base-size)/2;
    return parseFloat(v)||0; };
  return {x:a.x+pos(pxStr,a.w,w), y:a.y+pos(pyStr,a.h,h), w, h};
}

/* ---------- 한 슬라이드 그리기 (데스크탑 renderSlide 포팅, src=data URL) ---------- */
function renderSlide(pres, slide, data, natMap){
  const resolveSrc=(src)=>{ if(!src) return null; if(/^data:/i.test(src)) return {data:src}; if(/^https?:/i.test(src)) return {path:src}; if(/^#/.test(src)) return null; return null; };
  const drawImg=(o)=>{ try{
    let r=null;
    if(o.invert && /logo-white/i.test(o.src||'')){ /* 웹: 흑백 변형본 없음 → 원본 사용 */ }
    r=resolveSrc(o.src); if(!r) return;
    const io=Object.assign({}, r, {x:o.g.x*IN, y:o.g.y*IN, w:o.g.w*IN, h:o.g.h*IN, sizing:{type:'cover', w:o.g.w*IN, h:o.g.h*IN}, transparency:o.op<1?Math.round((1-o.op)*100):0});
    if(o.zoom && o.zoom.from>1.008) io.objectName='ZOOM|'+Math.round(o.zoom.from*100000)+'|'+Math.round((o.zoom.dur||5)*1000);
    slide.addImage(io);
    if(o.bright!==undefined && o.bright<0.99) slide.addShape(pres.shapes.RECTANGLE, {x:o.g.x*IN,y:o.g.y*IN,w:o.g.w*IN,h:o.g.h*IN, fill:{color:'000000',transparency:Math.round(o.bright*100)}, line:{type:'none'}});
    if(o.sat!==undefined && o.sat<0.97) slide.addShape(pres.shapes.RECTANGLE, {x:o.g.x*IN,y:o.g.y*IN,w:o.g.w*IN,h:o.g.h*IN, fill:{color:'8A8079',transparency:Math.round(o.sat*100)}, line:{type:'none'}});
  }catch(e){} };
  const drawBgImg=(o)=>{ try{
    const r=resolveSrc(o.src); if(!r||!r.data) return;
    const nat=natMap[o.src] || {w:o.area.w, h:o.area.h};
    const pl=resolveBg(o.area, nat, o.size, o.posx, o.posy);
    const bo={data:r.data, x:pl.x*IN, y:pl.y*IN, w:pl.w*IN, h:pl.h*IN};
    if(o.zoom && o.zoom.from>1.008) bo.objectName='ZOOM|'+Math.round(o.zoom.from*100000)+'|'+Math.round((o.zoom.dur||5)*1000);
    slide.addImage(bo);
  }catch(e){} };
  const drawRect=(o)=>{
    const opt={x:o.g.x*IN, y:o.g.y*IN, w:o.g.w*IN, h:o.g.h*IN};
    if(o.fill) opt.fill={color:o.fill.hex, transparency:Math.round((1-o.fill.a)*100)}; else opt.fill={type:'none'};
    if(o.line) opt.line={color:o.line.hex, width:Math.max(0.25,o.line.w*PT), transparency:Math.round((1-o.line.a)*100)};
    if(o.rad && o.rad>2){ const half=Math.min(o.g.w,o.g.h)/2; opt.rectRadius=Math.min(o.rad,half)*IN; slide.addShape(pres.shapes.ROUNDED_RECTANGLE, opt); }
    else slide.addShape(pres.shapes.RECTANGLE, opt);
  };
  const drawLine=(o)=>{ slide.addShape(pres.shapes.LINE, {x:o.x1*IN,y:o.y1*IN,w:(o.x2-o.x1)*IN,h:(o.y2-o.y1)*IN, line:{color:o.color,width:Math.max(0.25,o.w*PT),transparency:Math.round((1-o.a)*100)}}); };
  const drawGrad=(o)=>{
    const hh=n=>('0'+Math.max(0,Math.min(255,Math.round(n))).toString(16)).slice(-2);
    const enc=o.stops.map(s=>{ const h=(hh(s.c[0])+hh(s.c[1])+hh(s.c[2])).toUpperCase(); const pos=Math.round(Math.max(0,Math.min(1,s.p))*100000); const al=Math.round(Math.max(0,Math.min(1,s.c[3]===undefined?1:s.c[3]))*100000); return h+'@'+pos+'@'+al; }).join(';');
    const mid=o.stops.reduce((a,s)=>(s.c[3]||0)>(a.c[3]||0)?s:a, o.stops[0]);
    const fhex=(hh(mid.c[0])+hh(mid.c[1])+hh(mid.c[2])).toUpperCase();
    slide.addShape(pres.shapes.RECTANGLE, {x:o.g.x*IN,y:o.g.y*IN,w:o.g.w*IN,h:o.g.h*IN, fill:{color:fhex,transparency:Math.round((1-(mid.c[3]||1))*100*0.6+20)}, line:{type:'none'}, objectName:'GRAD|'+(o.flip?1:0)+'|'+enc});
  };
  const nonText=data.ops.filter(o=>o.k!=='text').sort((a,b)=>(a.order||0)-(b.order||0));
  for(const o of nonText){ if(o.k==='img') drawImg(o); else if(o.k==='bgimg') drawBgImg(o); else if(o.k==='line') drawLine(o); else if(o.k==='grad') drawGrad(o); else drawRect(o); }
  const texts=data.ops.filter(o=>o.k==='text').sort((a,b)=>(a.order||0)-(b.order||0));
  for(const o of texts){
    const arr=[];
    o.runs.forEach(r=>{
      if(r.br){ if(arr.length) arr[arr.length-1].options.breakLine=true; return; }
      let t=r.text; if(r.upper) t=t.toUpperCase();
      const pf=pptFont(r.fam,r.weight);
      const op={fontFace:pf.face, fontSize:Math.max(4,r.size), italic:r.italic, charSpacing:r.cs?Math.round(r.cs*10)/10:0};
      if(pf.bold) op.bold=true;
      if(r.outline){ op.outline={size:r.outline.w, color:r.outline.color}; op.color=(data.bg||'0E0D0B'); }
      else { op.color=r.color; op.transparency=r.alpha<1?Math.round((1-r.alpha)*100):0; }
      arr.push({text:t, options:op});
    });
    if(!arr.length) continue;
    const rsizes=o.runs.filter(r=>!r.br).map(r=>r.size||0);
    const maxSize=Math.max(...rsizes);
    const noWrap=maxSize>=40 || o.singleLine;
    const topt={x:o.g.x*IN, y:o.g.y*IN, w:o.g.w*IN+0.12, h:o.g.h*IN, align:o.align, valign:(o.valign||'top'), margin:0, wrap:!noWrap, autoFit:false};
    if(!o.singleLine && o.lhPt && maxSize<=24) topt.lineSpacing=o.lhPt; else topt.lineSpacingMultiple=o.lh;
    slide.addText(arr, topt);
  }
}

/* ---------- 그라데이션/줌 후처리 (jszip, 데스크탑 fixPptx 포팅) ---------- */
function buildTiming(anims){
  let id=5;
  const pars=anims.map(a=>{ const eId=id++, bId=id++;
    return `<p:par><p:cTn id="${eId}" presetID="6" presetClass="emph" presetSubtype="0" fill="hold" grpId="0" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animScale><p:cBhvr><p:cTn id="${bId}" dur="${a.dur}" fill="hold"/><p:tgtEl><p:spTgt spid="${a.spid}"/></p:tgtEl></p:cBhvr><p:from x="${a.from}" y="${a.from}"/><p:to x="100000" y="100000"/></p:animScale></p:childTnLst></p:cTn></p:par>`;
  }).join('');
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${pars}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
}
async function fixPptx(blob){
  try{
    const zip=await JSZip.loadAsync(blob);
    const slides=Object.keys(zip.files).filter(n=>/^ppt\/slides\/slide\d+\.xml$/.test(n));
    let changed=false;
    for(const sn of slides){
      let xml=await zip.file(sn).async('string');
      const hasGrad=xml.indexOf('name="GRAD|')>=0, hasZoom=xml.indexOf('name="ZOOM|')>=0;
      if(!hasGrad && !hasZoom) continue;
      if(hasGrad){
        xml=xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g,(sp)=>{
          const m=sp.match(/name="GRAD\|([01])\|([^"]*)"/); if(!m) return sp;
          const flip=m[1]==='1';
          let stops=m[2].split(';').map(s=>{const p=s.split('@');return{h:p[0],pos:parseInt(p[1],10),a:parseInt(p[2],10)};});
          if(flip) stops=stops.slice().reverse().map(s=>({h:s.h,pos:100000-s.pos,a:s.a}));
          stops.sort((a,b)=>a.pos-b.pos);
          const gs=stops.map(s=>`<a:gs pos="${s.pos}"><a:srgbClr val="${s.h}"><a:alpha val="${s.a}"/></a:srgbClr></a:gs>`).join('');
          const grad=`<a:gradFill rotWithShape="1"><a:gsLst>${gs}</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`;
          let out=sp.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, grad);
          out=out.replace(/(name=")GRAD\|[^"]*(")/, '$1Gradient$2');
          return out;
        });
      }
      if(hasZoom){
        const anims=[];
        xml=xml.replace(/<p:cNvPr id="(\d+)" name="ZOOM\|(\d+)\|(\d+)"/g,(m,id,from,dur)=>{ anims.push({spid:id,from:Math.min(parseInt(from,10),160000),dur:Math.max(300,Math.min(parseInt(dur,10),20000))}); return `<p:cNvPr id="${id}" name="Background"`; });
        if(anims.length && xml.indexOf('<p:timing>')<0) xml=xml.replace('</p:sld>', buildTiming(anims)+'</p:sld>');
      }
      zip.file(sn, xml); changed=true;
    }
    if(!changed) return blob;
    return await zip.generateAsync({type:'blob', mimeType:'application/vnd.openxmlformats-officedocument.presentationml.presentation'});
  }catch(e){ console.warn('후처리 건너뜀', e); return blob; }
}

/* ---------- 파일 → 에셋 맵 (data URL) + HTML 재작성 ---------- */
function readAsDataURL(file){ return new Promise((res,rej)=>{ const r=new FileReader(); r.onload=()=>res(r.result); r.onerror=rej; r.readAsDataURL(file); }); }
function readAsText(file){ return new Promise((res,rej)=>{ const r=new FileReader(); r.onload=()=>res(r.result); r.onerror=rej; r.readAsText(file); }); }
function imgNatural(dataURL){ return new Promise(res=>{ const im=new Image(); im.onload=()=>res({w:im.naturalWidth,h:im.naturalHeight}); im.onerror=()=>res(null); im.src=dataURL; }); }
function baseName(p){ return p.replace(/^.*[\\/]/,'').replace(/[?#].*$/,''); }

async function buildAssets(files, htmlPath){
  // 이미지/CSS/폰트 → dataURL. 키: 파일명(basename) + 상대경로 둘 다
  const map={}, natMap={};
  for(const f of files){
    const rel=(f.webkitRelativePath||f.name);
    if(/\.html?$/i.test(f.name)) continue;
    if(/\.(png|jpe?g|gif|webp|svg|avif|bmp)$/i.test(f.name)){
      try{ const d=await readAsDataURL(f); map[baseName(rel)]=d; map[rel]=d; }catch(e){}
    }
  }
  return {map, natMap};
}
function rewriteHtml(html, map){
  // src="..." 와 url(...) 의 로컬 경로를 dataURL로 (http/data는 유지)
  const lookup=(u)=>{ u=u.trim().replace(/^["']|["']$/g,''); if(/^(data:|https?:|#)/i.test(u)) return null; return map[baseName(u)]||map[u]||null; };
  html=html.replace(/(\s(?:src|href))\s*=\s*(["'])([^"']+)\2/gi,(m,attr,q,u)=>{ if(attr.trim()==='href' && !/\.(png|jpe?g|gif|webp|svg|avif)$/i.test(u)) return m; const d=lookup(u); return d?`${attr}=${q}${d}${q}`:m; });
  html=html.replace(/url\(\s*(["']?)([^"')]+)\1\s*\)/gi,(m,q,u)=>{ const d=lookup(u); return d?`url(${q}${d}${q})`:m; });
  return html;
}

/* ---------- 변환 ---------- */
function waitFrame(){ return new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))); }
async function convert(){
  const log=$('#log'); const setLog=t=>{ log.textContent=t; };
  if(!state.files.length){ setLog('HTML 파일이 포함된 폴더(또는 파일들)를 선택하세요.'); return; }
  $('#go').disabled=true; $('#result').hidden=true; $('#progress').hidden=false;
  try{
    const htmlFile=state.files.find(f=>/\.html?$/i.test(f.name));
    if(!htmlFile){ throw new Error('HTML 파일을 못 찾았습니다.'); }
    setLog('에셋 읽는 중…');
    const {map}=await buildAssets(state.files);
    let html=await readAsText(htmlFile);
    html=rewriteHtml(html, map);

    // iframe 준비 (1920 폭 / 세로는 충분히 크게)
    const portrait = state.orient==='portrait';
    const RW = 1920;  // 항상 1920 폭으로 추출 (덱은 1:1, 문서는 세로 슬라이스)
    const ifr=$('#stage'); ifr.style.width=RW+'px'; ifr.style.height='1080px';
    setLog('렌더링 중…');
    await new Promise(res=>{ ifr.onload=()=>res(); ifr.srcdoc=html; });
    const win=ifr.contentWindow, doc=ifr.contentDocument;
    // extract.js 주입
    await fetch('extract.js').then(r=>r.text()).then(src=>{ const s=doc.createElement('script'); s.textContent=src; doc.head.appendChild(s); });
    // 폰트·이미지 로딩 + 줌감지 + reveal 강제
    try{ await win.document.fonts.ready; }catch(e){}
    win.__detectZoom && win.__detectZoom();
    win.__forceReveal && win.__forceReveal();
    await new Promise(r=>setTimeout(r,900));
    await waitFrame();

    const N = win.__sectionCount();
    const pres=new PptxGenJS();
    // 자연 이미지 크기맵 (bgimg용)
    const natMap={};
    const pres_render=async()=>{};

    if(N>=1){
      pres.defineLayout({name:'D', width:13.333, height:7.5}); pres.layout='D';
      setLog(`슬라이드 ${N}장 변환 중…`);
      for(let i=1;i<=N;i++){
        const data=win.__extractSlide(i,{});
        // bgimg 자연크기 채우기
        if(data&&data.ops) for(const o of data.ops){ if(o.k==='bgimg'&&!natMap[o.src]){ const n=await imgNatural(o.src); if(n)natMap[o.src]=n; } }
        const slide=pres.addSlide(); if(data&&data.bg) slide.background={color:data.bg};
        if(data) renderSlide(pres, slide, data, natMap);
      }
    } else {
      const PAGE_H = portrait?1920:1080;
      pres.defineLayout({name:'DOC', width:portrait?7.5:13.333, height:portrait?13.333:7.5}); pres.layout='DOC';
      setLog('문서 모드(세로 슬라이스) 변환 중…');
      const full=win.__extractSlide(1,{mode:'doc',renderW:RW});
      const opBottom=o=> o.k==='line'?Math.max(o.y1,o.y2):(o.k==='bgimg'?o.area.y+o.area.h:(o.g?o.g.y+o.g.h:0));
      const anchorY=o=> o.k==='line'?Math.min(o.y1,o.y2):(o.k==='bgimg'?o.area.y:(o.g?o.g.y:0));
      const shift=(o,dy)=>{ const n=Object.assign({},o); if(o.k==='line'){n.y1=o.y1+dy;n.y2=o.y2+dy;} else if(o.k==='bgimg'){n.area=Object.assign({},o.area,{y:o.area.y+dy});} else if(o.g){n.g=Object.assign({},o.g,{y:o.g.y+dy});} return n; };
      if(full&&full.ops&&full.ops.length){
        for(const o of full.ops){ if(o.k==='bgimg'&&!natMap[o.src]){ const n=await imgNatural(o.src); if(n)natMap[o.src]=n; } }
        let H=full.height||0; for(const o of full.ops){ const b=opBottom(o); if(b>H)H=b; }
        const nS=Math.max(1,Math.ceil(H/PAGE_H));
        for(let p=0;p<nS;p++){ const y0=p*PAGE_H,y1=y0+PAGE_H;
          const ops=full.ops.filter(o=>{const ay=anchorY(o);return ay>=y0&&ay<y1;}).map(o=>shift(o,-y0));
          const slide=pres.addSlide(); if(full.bg) slide.background={color:full.bg};
          renderSlide(pres, slide, {bg:full.bg, ops}, natMap);
        }
      } else { pres.addSlide(); }
    }

    setLog('PPTX 만드는 중…');
    let blob=await pres.write({outputType:'blob'});
    blob=await fixPptx(blob);
    const name=htmlFile.name.replace(/\.[^.]+$/,'')+'.pptx';
    const url=URL.createObjectURL(blob);
    const a=$('#dl'); a.href=url; a.download=name; a.textContent='⬇ '+name+' 다운로드';
    $('#progress').hidden=true; $('#result').hidden=false;
    // 폰트 경고
    const customFont = /Cormorant|Spectral|Playfair|Montserrat|Poppins|Lora|EB Garamond/i.test(html) || /@font-face/i.test(html);
    $('#fontnote').hidden = !customFont;
  }catch(e){
    $('#progress').hidden=true; setLog('오류: '+e.message); $('#result').hidden=false;
  }
  $('#go').disabled=false;
}

/* ---------- UI ---------- */
function setFiles(list){ state.files=[...list]; const names=state.files.filter(f=>/\.html?$/i.test(f.name)).map(f=>f.name);
  $('#picked').textContent = state.files.length? `${state.files.length}개 파일 · HTML: ${names.join(', ')||'없음'}` : '';
  $('#go').disabled=!names.length;
}
window.addEventListener('DOMContentLoaded',()=>{
  $('#folder').addEventListener('change',e=>setFiles(e.target.files));
  $('#filesinp').addEventListener('change',e=>setFiles(e.target.files));
  document.querySelectorAll('#orient button').forEach(b=>b.onclick=()=>{ state.orient=b.dataset.o; document.querySelectorAll('#orient button').forEach(x=>x.classList.toggle('on',x===b)); });
  $('#go').onclick=convert;
  const drop=$('#drop');
  drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('over');});
  drop.addEventListener('dragleave',()=>drop.classList.remove('over'));
  drop.addEventListener('drop',async e=>{ e.preventDefault(); drop.classList.remove('over');
    const items=e.dataTransfer.items; if(items&&items[0]&&items[0].webkitGetAsEntry){ const fl=[]; const walk=(entry)=>new Promise(res=>{ if(entry.isFile) entry.file(f=>{fl.push(f);res();}); else if(entry.isDirectory){ const rd=entry.createReader(); rd.readEntries(async ents=>{ for(const en of ents) await walk(en); res(); }); } else res(); });
      const roots=[...items].map(i=>i.webkitGetAsEntry()).filter(Boolean); for(const r of roots) await walk(r); setFiles(fl);
    } else setFiles(e.dataTransfer.files);
  });
});
