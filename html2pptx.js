// HTML deck -> native editable PPTX
//   node html2pptx.js <input.html> [output.pptx]
// Reads computed geometry from the rendered page (puppeteer + installed Chrome)
// and rebuilds each slide as native PowerPoint text/shapes/images (pptxgenjs).
const puppeteer = require('puppeteer-core');
const PptxGenJS = require('pptxgenjs');
const path = require('path');
const fs = require('fs');

// 이미지 파일의 자연 크기 읽기 (PNG/JPEG/GIF) — 배경이미지 비율 계산용
function imgSize(file){
  try{
    const fd=fs.openSync(file,'r'); const buf=Buffer.alloc(64); const n=fs.readSync(fd,buf,0,64,0); fs.closeSync(fd);
    if(buf[0]===0x89&&buf[1]===0x50) return {w:buf.readUInt32BE(16), h:buf.readUInt32BE(20)};          // PNG
    if(buf[0]===0x47&&buf[1]===0x49) return {w:buf.readUInt16LE(6), h:buf.readUInt16LE(8)};            // GIF
    if(buf[0]===0xFF&&buf[1]===0xD8){                                                                   // JPEG
      const b=fs.readFileSync(file); let o=2;
      while(o<b.length){ if(b[o]!==0xFF){o++;continue;} const m=b[o+1];
        if(m>=0xC0&&m<=0xCF&&m!==0xC4&&m!==0xC8&&m!==0xCC){ return {h:b.readUInt16BE(o+5), w:b.readUInt16BE(o+7)}; }
        o+=2+b.readUInt16BE(o+2); }
    }
  }catch(e){}
  return null;
}
// CSS background-size/position → 실제 이미지 배치(인치 기준 px) 계산
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

// ---- locate an installed Chrome / Edge ----
function findChrome(){
  const c = [
    process.env.CHROME_PATH,
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
    (process.env.LOCALAPPDATA||'')+'\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
  ].filter(Boolean);
  for(const p of c){ try{ if(fs.existsSync(p)) return p; }catch(e){} }
  return null;
}

const HTML = process.argv[2];
if(!HTML){ console.error('Usage: node html2pptx.js <input.html> [output.pptx]'); process.exit(1); }
const OUT = process.argv[3] || path.join(path.dirname(HTML), path.basename(HTML).replace(/\.[^.]+$/,'') + '.pptx');
const ORIENT = /portrait|세로/i.test(process.argv[4]||'') ? 'portrait' : 'landscape';   // 문서 모드 슬라이스 방향 (덱 모드엔 영향 없음)
const CHROME = findChrome();
if(!CHROME){ console.error('Chrome/Edge를 찾을 수 없습니다. Chrome을 설치하거나 CHROME_PATH 환경변수를 설정하세요.'); process.exit(1); }

const IN = 13.333 / 1920;   // px -> inch  (슬라이드 폭이 1920px HTML과 1:1)
const PT = IN * 72;         // px -> pt  (위치·글자크기와 동일한 척도. 선 두께도 이 값으로 HTML 그대로 재현)
function fileUrl(p){ return 'file:///' + p.replace(/\\/g,'/'); }

function pptFont(fam, w, italic){
  fam=fam||''; w=w||400;
  const pick=(base,steps)=>{ let s=''; for(const [mw,suf] of steps){ if(w>=mw) s=suf; } return (base+s); };
  if(/Pretendard/i.test(fam)) return pick('Pretendard',[[0,''],[500,' Medium'],[600,' SemiBold'],[800,' ExtraBold'],[900,' Black']]);
  if(/JetBrains|Plex Mono|monospace/i.test(fam)) return pick('JetBrains Mono',[[0,''],[500,' Medium']]);
  if(/Cormorant/i.test(fam)) return 'Cormorant Garamond Light'; // 설치된 정확한 패밀리(인용문=light weight)
  if(/Spectral/i.test(fam)) return pick('Spectral',[[0,' Light'],[400,''],[500,' Medium']]);
  return pick('Inter',[[0,''],[500,' Medium'],[600,' SemiBold'],[800,' ExtraBold'],[900,' Black']]);
}

async function extractSlide(page, idx, opts){
  return await page.evaluate((idx, opts) => {
    opts = opts || {};
    let sec;
    if(opts.mode==='doc'){ sec = document.body; }
    else { const secs=[...document.querySelectorAll('section')]; sec = secs[idx-1]; }
    if(!sec) return null;
    const SR = sec.getBoundingClientRect();
    const scale = opts.mode==='doc' ? 1 : SR.width/1920;
    const rel = (el)=>{const r=el.getBoundingClientRect();return{x:(r.left-SR.left)/scale,y:(r.top-SR.top)/scale,w:r.width/scale,h:r.height/scale};};
    const _ccv=document.createElement('canvas'); _ccv.width=_ccv.height=1; const _ccx=_ccv.getContext('2d',{willReadFrequently:true});
    function col(c){
      if(!c||c==='transparent'||c==='rgba(0, 0, 0, 0)'||c==='none') return null;
      const h=n=>('0'+Math.round(n).toString(16)).slice(-2);
      const m=c.match(/^rgba?\(([^)]+)\)/);
      if(m){ const p=m[1].split(/[,\s\/]+/).filter(s=>s!=='').map(s=>parseFloat(s));
        return {hex:(h(p[0])+h(p[1])+h(p[2])).toUpperCase(), a:p[3]===undefined?1:p[3]}; }
      // rgb 외 색(oklch/oklab/hsl/lab/color()/named 등) → canvas 픽셀로 실제 sRGB 변환
      try{
        _ccx.clearRect(0,0,1,1); _ccx.fillStyle='#000'; _ccx.fillStyle=c; _ccx.fillRect(0,0,1,1);
        const d=_ccx.getImageData(0,0,1,1).data; const a=d[3]/255;
        if(a<=0) return null;
        return {hex:(h(d[0])+h(d[1])+h(d[2])).toUpperCase(), a};
      }catch(e){ return null; }
    }
    const isBlock = (d)=>/(block|flex|grid|list-item|table)/.test(d) || d==='inline-block' || d==='inline-flex';
    // 인라인 구문요소: flex/grid 컨테이너의 자식이면 display가 block으로 blockify되지만 의미상 인라인
    const PHRASING=new Set(['B','STRONG','I','EM','SMALL','SPAN','A','SUP','SUB','MARK','U','CODE','ABBR','TIME','LABEL','FONT','CITE','Q','S','INS','DEL','BDI','BDO']);
    const isBlockEl=(el)=>{
      if(!isBlock(getComputedStyle(el).display)) return false;
      const par=el.parentElement;
      if(par && /(flex|grid)/.test(getComputedStyle(par).display) && PHRASING.has(el.tagName)){
        // 부모가 직접 텍스트노드를 가진 '텍스트 줄'이고, 자신이 칩(패딩/테두리/배경)이 아닐 때만 인라인 취급
        const hasText=[...par.childNodes].some(n=>n.nodeType===3 && n.textContent.trim());
        const cs=getComputedStyle(el); const cb=col(cs.backgroundColor);
        const looksChip=(parseFloat(cs.paddingLeft)||0)>2 || cs.borderLeftStyle!=='none' || (cb&&cb.a>0.02);
        if(hasText && !looksChip) return false;
      }
      return true;
    };
    // 부모보다 폰트가 훨씬 큰 인라인 강조(예: 작은 제목 옆 거대한 이탤릭) = 독립 박스로 분리
    const bigAccent=(el)=>false; // (비활성) 거대 인라인 분리는 상단 겹침 유발 → 되돌림
    // 최상위 콤마로 background 레이어 분리(괄호 안 콤마 보호)
    const splitTop=(s)=>{ const out=[]; let d=0,cur=''; for(const ch of s){ if(ch==='(')d++; else if(ch===')')d--; if(ch===','&&d===0){out.push(cur);cur='';} else cur+=ch; } if(cur.trim())out.push(cur); return out; };
    // 여러 그라데이션 레이어 = 알파 합성(평균 X). {hex,a} 반환
    const gradComposite=(bi)=>{
      const layers=splitTop(bi).filter(s=>/gradient/.test(s));
      if(!layers.length) return null;
      let cr=null,ca=0;
      for(let li=layers.length-1; li>=0; li--){            // 아래 레이어부터 위로 합성
        const ps=[...layers[li].matchAll(/rgba?\(([^)]+)\)/g)].map(m=>m[1].split(',').map(Number));
        if(!ps.length) continue;
        let r=0,gg=0,b=0,wsum=0,al=0;
        ps.forEach(p=>{const a=p[3]===undefined?1:p[3]; r+=p[0]*a; gg+=p[1]*a; b+=p[2]*a; wsum+=a; al+=a;});
        const la=al/ps.length, lc= wsum>0?[r/wsum,gg/wsum,b/wsum]:[ps[0][0],ps[0][1],ps[0][2]];
        if(cr===null){ cr=lc; ca=la; }
        else { const oA=la+ca*(1-la); const mx=i=> oA>0?(lc[i]*la+cr[i]*ca*(1-la))/oA:lc[i]; cr=[mx(0),mx(1),mx(2)]; ca=oA; }
      }
      if(cr===null) return null;
      const h=n=>('0'+Math.max(0,Math.min(255,Math.round(n))).toString(16)).slice(-2);
      return {hex:(h(cr[0])+h(cr[1])+h(cr[2])).toUpperCase(), a:ca};
    };
    // 선형 그라데이션 파싱 → {angle, stops:[{c:[r,g,b,a],p}]}
    const parseLinearGrad=(s)=>{
      const m=s.match(/linear-gradient\((.*)\)\s*$/);
      if(!m) return null;
      const parts=splitTop(m[1]); let angle=180, i=0;
      const first=(parts[0]||'').trim();
      if(/deg\s*$/.test(first)){ angle=parseFloat(first); i=1; }
      else if(/^to\s/i.test(first)){ angle=/bottom/i.test(first)?180:/top/i.test(first)?0:/right/i.test(first)?90:270; i=1; }
      const stops=[];
      for(;i<parts.length;i++){
        const cstr=(parts[i].match(/rgba?\([^)]+\)/)||[])[0];
        if(!cstr) continue;
        const c=col(cstr); if(!c) continue;
        const pm=parts[i].match(/([\d.]+)%/);
        stops.push({c:[parseInt(c.hex.slice(0,2),16),parseInt(c.hex.slice(2,4),16),parseInt(c.hex.slice(4,6),16),c.a], p:pm?parseFloat(pm[1])/100:null});
      }
      if(!stops.length) return null;
      if(stops[0].p==null) stops[0].p=0; if(stops[stops.length-1].p==null) stops[stops.length-1].p=1;
      for(let k=1;k<stops.length-1;k++){ if(stops[k].p==null){ let a=k-1; while(stops[a].p==null)a--; let b=k+1; while(b<stops.length&&stops[b].p==null)b++; stops[k].p=stops[a].p+((stops[b].p-stops[a].p)*(k-a)/(b-a)); } }
      return {angle, stops};
    };
    // 그라데이션 → op: 단일 수직 선형이면 밴드(grad), 그 외엔 평면 합성(rect)
    const makeGradOp=(bi,g,order)=>{
      const layers=splitTop(bi).filter(s=>/gradient/.test(s));
      if(layers.length===1 && /linear-gradient/.test(layers[0])){
        const lg=parseLinearGrad(layers[0]);
        if(lg && lg.stops.length>=2){ const a=((lg.angle%360)+360)%360;
          if(Math.abs(a-180)<=8 || a<=8 || Math.abs(a-360)<=8) return {k:'grad',order,g,flip:(a<=8||Math.abs(a-360)<=8),stops:lg.stops};
        }
      }
      const cc=gradComposite(bi); return cc? {k:'rect',order,g,fill:cc,line:null} : null;
    };
    const ops=[];
    const allEls=[...sec.querySelectorAll('*')];      // DOM(=paint) 순서
    const idxOf=(el)=>{ const i=allEls.indexOf(el); return i<0?9999:i; };

    sec.querySelectorAll('img').forEach(img=>{
      const g=rel(img); if(g.w<3||g.h<3) return;
      const cs=getComputedStyle(img);
      const filt=cs.filter||'';
      const invert=/invert\(/.test(filt);
      const fm=filt.match(/brightness\(([\d.]+)\)/);
      const sm=filt.match(/saturate\(([\d.]+)\)/);
      ops.push({k:'img', order:idxOf(img), src:img.getAttribute('src'), g, fit:cs.objectFit, op:parseFloat(cs.opacity),
        bright: invert?1:(fm?parseFloat(fm[1]):1), invert, sat: sm?parseFloat(sm[1]):1});
    });

    // 커스텀 이미지 컴포넌트(<image-slot src>) + src 속성 가진 커스텀 요소도 이미지로 처리 (shadow DOM 안에 그려짐)
    sec.querySelectorAll('image-slot[src], [data-img-src], picture').forEach(el=>{
      if(el.tagName==='PICTURE'){ if(el.querySelector('img')) return; }
      const src=el.getAttribute('src')||el.getAttribute('data-img-src'); if(!src) return;
      const g=rel(el); if(g.w<3||g.h<3) return;
      ops.push({k:'img', order:idxOf(el), src, g, fit:(el.getAttribute('fit')||'cover'), op:1, bright:1, invert:false});
    });

    sec.querySelectorAll('*').forEach(el=>{
      if(el.tagName==='IMG'||el.tagName==='BR'||el.tagName==='IMAGE-SLOT') return;
      const cs=getComputedStyle(el);
      if(cs.display==='none'||cs.visibility==='hidden') return;
      const g=rel(el);
      const bg=col(cs.backgroundColor);
      const ord=idxOf(el);
      // 얇은 선/바 (bg색 또는 그라데이션으로 그린 divider) → 선으로
      const biEl=cs.backgroundImage;
      const lineCol = (bg && bg.a>0.05) ? bg : ((biEl && biEl!=='none' && /gradient/.test(biEl)) ? gradComposite(biEl) : null);
      if(g.h>0 && g.h<=4 && g.w>=8 && lineCol && lineCol.a>0.05){ ops.push({k:'line',order:ord,x1:g.x,y1:g.y+g.h/2,x2:g.x+g.w,y2:g.y+g.h/2,color:lineCol.hex,a:lineCol.a,w:Math.max(g.h,1)}); return; }
      if(g.w>0 && g.w<=4 && g.h>=8 && lineCol && lineCol.a>0.05){ ops.push({k:'line',order:ord,x1:g.x+g.w/2,y1:g.y,x2:g.x+g.w/2,y2:g.y+g.h,color:lineCol.hex,a:lineCol.a,w:Math.max(g.w,1)}); return; }
      if(g.w<4||g.h<4) return;
      // border-radius (둥근 모서리) 측정
      const rad=Math.min((parseFloat(cs.borderTopLeftRadius)||0)/scale, Math.min(g.w,g.h)/2);
      const W=s=>parseFloat(cs.getPropertyValue('border-'+s+'-width'))||0;
      const C=s=>col(cs.getPropertyValue('border-'+s+'-color'));
      const wT=W('top'),wB=W('bottom'),wL=W('left'),wR=W('right'); const cT=C('top');
      const uniformBorder = wT>0.4 && Math.abs(wT-wB)<0.6 && Math.abs(wT-wL)<0.6 && Math.abs(wT-wR)<0.6 && cT && cT.a>0.05;
      if(uniformBorder && rad>2){
        // 둥근 테두리 박스 = 하나의 rounded rect (채움+외곽선)
        ops.push({k:'rect', order:ord, g, fill:(bg&&bg.a>0.02?{hex:bg.hex,a:bg.a}:null), line:{hex:cT.hex,a:cT.a,w:wT}, rad});
      } else {
        if(bg && bg.a>0.02){ ops.push({k:'rect', order:ord, g, fill:{hex:bg.hex,a:bg.a}, line:null, rad:rad>2?rad:0}); }
        // 테두리: 면별로 선 추출 (한쪽만 있어도 잡음)
        const sides=[['top',g.x,g.y,g.x+g.w,g.y],['bottom',g.x,g.y+g.h,g.x+g.w,g.y+g.h],['left',g.x,g.y,g.x,g.y+g.h],['right',g.x+g.w,g.y,g.x+g.w,g.y+g.h]];
        for(const [s,x1,y1,x2,y2] of sides){ const w=W(s),c=C(s); if(w>0.4 && c && c.a>0.05) ops.push({k:'line',order:ord,x1,y1,x2,y2,color:c.hex,a:c.a,w}); }
      }
      const bi=cs.backgroundImage;
      // 래스터 배경이미지(url) → 이미지로 (헤더 로고 등 CSS background 처리)
      if(bi && bi!=='none' && !/gradient/.test(bi)){
        const um=bi.match(/url\((['"]?)([^'")]+)\1\)/);
        if(um){
          const bl=parseFloat(cs.borderLeftWidth)||0, bt=parseFloat(cs.borderTopWidth)||0, br=parseFloat(cs.borderRightWidth)||0, bb=parseFloat(cs.borderBottomWidth)||0;
          const area={x:g.x+bl/scale, y:g.y+bt/scale, w:g.w-(bl+br)/scale, h:g.h-(bt+bb)/scale};
          ops.push({k:'bgimg', order:ord, area, src:um[2], size:cs.backgroundSize, posx:cs.backgroundPositionX, posy:cs.backgroundPositionY});
        }
      }
      if(bi && bi!=='none' && /gradient/.test(bi) && g.w>800 && g.h>500){
        const go=makeGradOp(bi, g, ord); if(go) ops.push(go);
      }
      // 절대배치 가상요소(::before/::after) 배경 = 오버레이·구분선·악센트 바
      for(const pe of ['::before','::after']){
        let cs2; try{ cs2=getComputedStyle(el,pe); }catch(e){ continue; }
        if(!cs2 || cs2.content==='none' || cs2.content==='') continue;
        if(!/absolute|fixed/.test(cs2.position)) continue;       // 흐름배치 마커는 위치 추정 불가 → 제외
        const pbg=col(cs2.backgroundColor);
        const pbi=cs2.backgroundImage;
        const hasGrad=pbi && pbi!=='none' && /gradient/.test(pbi);
        if(!(pbg&&pbg.a>0.03) && !hasGrad) continue;
        const blw=parseFloat(cs.borderLeftWidth)||0, btw=parseFloat(cs.borderTopWidth)||0, brw=parseFloat(cs.borderRightWidth)||0, bbw=parseFloat(cs.borderBottomWidth)||0;
        const pr=el.getBoundingClientRect();
        const cbL=pr.left+blw, cbT=pr.top+btw, cbW=pr.width-blw-brw, cbH=pr.height-btw-bbw;
        const num=(v,base)=>{ if(v==null||v==='auto')return null; if(/%$/.test(v))return parseFloat(v)/100*base; return parseFloat(v); };
        const L=num(cs2.left,cbW), R=num(cs2.right,cbW), T=num(cs2.top,cbH), B=num(cs2.bottom,cbH);
        const Wd=num(cs2.width,cbW), Hd=num(cs2.height,cbH);
        let px,pw,py,ph;
        if(L!=null&&R!=null){px=cbL+L;pw=cbW-L-R;} else if(Wd!=null){px=cbL+(L!=null?L:(R!=null?cbW-R-Wd:0));pw=Wd;} else {px=cbL;pw=cbW;}
        if(T!=null&&B!=null){py=cbT+T;ph=cbH-T-B;} else if(Hd!=null){py=cbT+(T!=null?T:(B!=null?cbH-B-Hd:0));ph=Hd;} else {py=cbT;ph=cbH;}
        // 가상요소 transform(scale/translate) 적용 — getBoundingClientRect가 없으므로 직접 (예: scaleX(.32)로 길이 다른 악센트 바)
        const tm=(cs2.transform||'none').match(/matrix\(([^)]+)\)/);
        if(tm){
          const m=tm[1].split(',').map(parseFloat); const a=m[0],d=m[3],e=m[4],f=m[5];
          const to=(cs2.transformOrigin||'0px 0px').split(' ').map(parseFloat);
          const oX=px+(to[0]||0), oY=py+(to[1]||0);
          px=oX+(px-oX)*a+e; pw=pw*a;
          py=oY+(py-oY)*d+f; ph=ph*d;
        }
        const gp={x:(px-SR.left)/scale, y:(py-SR.top)/scale, w:pw/scale, h:ph/scale};
        if(gp.w<1||gp.h<1) continue;
        let mx=ord; el.querySelectorAll('*').forEach(d=>{const i=idxOf(d); if(i>mx)mx=i;});
        const ordP = pe==='::after'? mx+0.5 : ord-0.5;
        if(pbg && pbg.a>0.03){
          if(gp.h>0&&gp.h<=4&&gp.w>=8){ ops.push({k:'line',order:ordP,x1:gp.x,y1:gp.y+gp.h/2,x2:gp.x+gp.w,y2:gp.y+gp.h/2,color:pbg.hex,a:pbg.a,w:Math.max(gp.h,1)}); continue; }
          if(gp.w>0&&gp.w<=4&&gp.h>=8){ ops.push({k:'line',order:ordP,x1:gp.x+gp.w/2,y1:gp.y,x2:gp.x+gp.w/2,y2:gp.y+gp.h,color:pbg.hex,a:pbg.a,w:Math.max(gp.w,1)}); continue; }
          ops.push({k:'rect',order:ordP,g:gp,fill:pbg,line:null});
        } else if(hasGrad){
          const go=makeGradOp(pbi, gp, ordP); if(go) ops.push(go);
        }
      }
    });

    const walker=document.createTreeWalker(sec, NodeFilter.SHOW_TEXT);
    const groups=new Map(); let tn;
    while((tn=walker.nextNode())){
      if(!tn.textContent.trim()) continue;
      let blk=tn.parentElement;
      while(blk && blk!==sec){ if(isBlockEl(blk)||bigAccent(blk)) break; blk=blk.parentElement; }
      if(!blk) blk=tn.parentElement;
      if(!groups.has(blk)) groups.set(blk,[]);
      groups.get(blk).push(tn);
    }
    groups.forEach((nodes, blk)=>{
      const cs=getComputedStyle(blk);
      let g=rel(blk); if(g.w<4||g.h<4) return;
      const hasBlockChild=[...blk.children].some(c=>isBlockEl(c));
      // 실제 텍스트(Range) 위치 측정 — 찌부러진 박스(overflow)·혼합콘텐츠에서 박스 top이 실제 글자 위치와 다름
      let l=1e9,t=1e9,r=-1e9,bt=-1e9,ok=false;
      for(const n of nodes){ const rg=document.createRange(); rg.selectNodeContents(n);
        for(const c of rg.getClientRects()){ if(c.width<=0||c.height<=0) continue; ok=true; l=Math.min(l,c.left); t=Math.min(t,c.top); r=Math.max(r,c.right); bt=Math.max(bt,c.bottom); } }
      let vAlign='top';
      const fsBlk=parseFloat(cs.fontSize)||0;
      let txRect=null;
      if(ok){
        txRect={x:(l-SR.left)/scale, y:(t-SR.top)/scale, w:(r-l)/scale, h:(bt-t)/scale};
        if(hasBlockChild && fsBlk<48) g=txRect;                       // 작은 라벨+본문 혼합(캡션 등)만 텍스트 박스로. 큰 제목 혼합(THREE+인라인블록)은 박스 유지
        else if(!hasBlockChild && fsBlk<48 && txRect.h > g.h*1.5){ g={x:g.x, y:txRect.y, w:g.w, h:txRect.h}; }  // 작은 본문의 찌부러진(overflow) 박스만 세로 보정
      }
      const runs=[];
      function walk(node){
        node.childNodes.forEach(ch=>{
          if(ch.nodeType===3){
            const t=ch.textContent.replace(/\s+/g,' ');
            if(!t.trim() && t!==' ') return;
            const es=getComputedStyle(ch.parentElement);
            const c=col(es.color)||{hex:'FAF8F4',a:1};
            const tsw=parseFloat(es.getPropertyValue('-webkit-text-stroke-width'))||0;
            const tsc=col(es.getPropertyValue('-webkit-text-stroke-color'));
            const fillCol=col(es.getPropertyValue('-webkit-text-fill-color'));
            const hollow = tsw>0 && (!fillCol || fillCol.a<0.05);
            runs.push({text:t, fam:es.fontFamily, weight:parseInt(es.fontWeight)||400,
              size:Math.round(parseFloat(es.fontSize)*0.5*10)/10, italic:es.fontStyle==='italic',
              color:c.hex, alpha:c.a,
              outline: hollow? {w:Math.max(0.5,tsw*0.5), color:(tsc?tsc.hex:'FAF8F4')} : null,
              cs:(parseFloat(es.letterSpacing)||0)*0.5, upper:es.textTransform==='uppercase'});
          } else if(ch.nodeType===1){
            if(ch.tagName==='BR'){ runs.push({br:true}); }
            else if(ch.tagName==='IMG'){}
            else if(isBlockEl(ch)){}
            else if(bigAccent(ch)){} // 독립 박스로 별도 처리
            else walk(ch);
          }
        });
      }
      walk(blk);
      if(!runs.length) return;
      const lhpx=parseFloat(cs.lineHeight);
      const lh = lhpx/parseFloat(cs.fontSize) || 1.2;
      const lhPt = isNaN(lhpx)? null : Math.round(lhpx*0.5*10)/10;   // HTML 줄높이를 절대값(pt)으로
      const hasBr=runs.some(r=>r.br);
      const maxPx=Math.max(...runs.filter(r=>!r.br).map(r=>(r.size||0)*2));
      // 실제 렌더링 줄 수로 판정 (박스 높이 X — overflow로 찌부러진 박스 오인 방지)
      let nLines=1;
      try{
        const rg=document.createRange(); rg.selectNodeContents(blk);
        const tops=[...rg.getClientRects()].filter(c=>c.width>0&&c.height>0).map(c=>c.top).sort((a,b)=>a-b);
        if(tops.length){ let prev=tops[0]; const gap=Math.max(6,(lhpx||16)*0.5);
          for(const t of tops){ if(t-prev>gap){ nLines++; prev=t; } } }
      }catch(e){}
      const singleLine = !hasBr && nLines<=1;   // 원본이 한 줄이면 PPT도 한 줄 유지
      // 큰 제목이 줄바꿈으로 감긴(wrapped, br 없음) 다줄 = 실제 글자 박스 + 세로 중앙정렬 → PowerPoint 행간 차이로 처지는 것 교정 (br로 나뉜 제목은 박스 그대로 두어 회귀 방지)
      if(txRect && fsBlk>=48 && nLines>1 && !hasBr && !hasBlockChild){ g={x:g.x, y:txRect.y, w:g.w, h:txRect.h}; vAlign='middle'; }
      // 패딩 보정: 테두리박스가 아닌 콘텐츠박스에서 텍스트 시작 (헤더 로고 자리 padding-left 등)
      // (혼합콘텐츠로 이미 실제 텍스트 Range를 쓴 경우는 패딩 보정 생략)
      const pL=(parseFloat(cs.paddingLeft)||0)/scale, pT=(parseFloat(cs.paddingTop)||0)/scale,
            pR=(parseFloat(cs.paddingRight)||0)/scale, pB=(parseFloat(cs.paddingBottom)||0)/scale;
      const gt = (!hasBlockChild && vAlign==='top' && (pL||pT||pR||pB))? {x:g.x+pL, y:g.y+pT, w:Math.max(4,g.w-pL-pR), h:Math.max(4,g.h-pT-pB)} : g;
      ops.push({k:'text', order:idxOf(blk), g:gt, valign:vAlign,
        align: cs.textAlign==='center'?'center':(cs.textAlign==='right'?'right':'left'),
        lh, lhPt, singleLine, runs});
    });

    let secBg = col(getComputedStyle(sec).backgroundColor);
    if((!secBg||secBg.a<0.02) && opts.mode==='doc'){ const hb=col(getComputedStyle(document.documentElement).backgroundColor); if(hb&&hb.a>0.02) secBg=hb; }
    return {bg: secBg? secBg.hex : null, ops, height: SR.height/scale};
  }, idx, opts);
}

(async ()=>{
  console.log('변환 시작:', path.basename(HTML));
  const browser=await puppeteer.launch({executablePath:CHROME, headless:'new', args:['--no-sandbox','--force-device-scale-factor=1']});
  const page=await browser.newPage();
  await page.setViewport({width:1920,height:1080,deviceScaleFactor:1});
  try{ await page.goto(fileUrl(HTML), {waitUntil:'domcontentloaded', timeout:60000}); }
  catch(e){ await page.goto(fileUrl(HTML), {waitUntil:'load', timeout:60000}).catch(()=>{}); }
  // 폰트·이미지 로딩 대기 (networkidle 의존 X = CDN/무거운 덱에서도 안전)
  try{ await page.evaluate(()=>document.fonts && document.fonts.ready); }catch(e){}
  try{ await page.evaluate(()=>Promise.all([...document.images].filter(i=>!i.complete).map(i=>new Promise(r=>{i.onload=i.onerror=r;setTimeout(r,3000);})))); }catch(e){}
  // deck-stage(버치형)일 때만 noscale+print로 펼침. 일반 덱은 화면(screen) 그대로 추출.
  const hasDeckStage = await page.evaluate(()=>{ const d=document.querySelector('deck-stage'); if(d){ d.setAttribute('noscale',''); return true; } return false; });
  if(hasDeckStage) await page.emulateMediaType('print');
  await new Promise(r=>setTimeout(r,1800));

  const N = await page.evaluate(()=>document.querySelectorAll('section').length);
  console.log('슬라이드 수:', N);

  const pres=new PptxGenJS();

  function renderSlide(slide, data){
    const drawImg=(o)=>{
      try{
        let srcRel=o.src;
        if(o.invert && /logo-white/i.test(srcRel)) srcRel=srcRel.replace(/logo-white/i,'logo-black');
        const src=srcRel.startsWith('http')? srcRel : path.join(path.dirname(HTML), srcRel);
        slide.addImage({path:src, x:o.g.x*IN, y:o.g.y*IN, w:o.g.w*IN, h:o.g.h*IN,
          sizing:{type:'cover', w:o.g.w*IN, h:o.g.h*IN}, transparency:o.op<1?Math.round((1-o.op)*100):0});
        if(o.bright!==undefined && o.bright<0.99){
          slide.addShape(pres.shapes.RECTANGLE, {x:o.g.x*IN, y:o.g.y*IN, w:o.g.w*IN, h:o.g.h*IN,
            fill:{color:'000000', transparency:Math.round(o.bright*100)}, line:{type:'none'}});
        }
        // saturate(<1)는 PPT 적용 불가 → 중성 회색으로 (1-sat)만큼 덮어 '탈채도' 근사 (밝은 배경색으로 덮으면 어두운 음영까지 하얘지므로 회색 사용)
        if(o.sat!==undefined && o.sat<0.97){
          slide.addShape(pres.shapes.RECTANGLE, {x:o.g.x*IN, y:o.g.y*IN, w:o.g.w*IN, h:o.g.h*IN,
            fill:{color:'8A8079', transparency:Math.round(o.sat*100)}, line:{type:'none'}});
        }
      }catch(e){}
    };
    const drawBgImg=(o)=>{
      try{
        let u=o.src, file;
        if(/^file:/i.test(u)) file=decodeURIComponent(u.replace(/^file:\/+/i,'')).replace(/\//g,'\\');
        else if(/^https?:/i.test(u)) file=u;
        else file=path.join(path.dirname(HTML), u);
        const nat=imgSize(file) || {w:o.area.w, h:o.area.h};
        const pl=resolveBg(o.area, nat, o.size, o.posx, o.posy);
        slide.addImage({path:file, x:pl.x*IN, y:pl.y*IN, w:pl.w*IN, h:pl.h*IN});
      }catch(e){}
    };
    const drawRect=(o)=>{
      const opt={x:o.g.x*IN, y:o.g.y*IN, w:o.g.w*IN, h:o.g.h*IN};
      if(o.fill) opt.fill={color:o.fill.hex, transparency:Math.round((1-o.fill.a)*100)}; else opt.fill={type:'none'};
      if(o.line) opt.line={color:o.line.hex, width:Math.max(0.25,o.line.w*PT), transparency:Math.round((1-o.line.a)*100)};
      if(o.rad && o.rad>2){
        const half=Math.min(o.g.w,o.g.h)/2;
        opt.rectRadius=Math.min(o.rad, half)*IN;   // 둥근 모서리 (pill은 half로 캡)
        slide.addShape(pres.shapes.ROUNDED_RECTANGLE, opt);
      } else slide.addShape(pres.shapes.RECTANGLE, opt);
    };
    const drawLine=(o)=>{
      slide.addShape(pres.shapes.LINE, {x:o.x1*IN, y:o.y1*IN, w:(o.x2-o.x1)*IN, h:(o.y2-o.y1)*IN,
        line:{color:o.color, width:Math.max(0.25,o.w*PT), transparency:Math.round((1-o.a)*100)}});
    };
    // 수직 선형 그라데이션 = 가로 띠(밴드) 여러 개로 근사 (위 투명→아래 어둠 등 정확히 재현)
    const drawGrad=(o)=>{
      const N=16, stops=o.stops;
      const sampleAt=(t)=>{ let s0=stops[0], s1=stops[stops.length-1];
        for(let k=0;k<stops.length-1;k++){ if(t>=stops[k].p && t<=stops[k+1].p){ s0=stops[k]; s1=stops[k+1]; break; } }
        const span=(s1.p-s0.p)||1, f=Math.min(1,Math.max(0,(t-s0.p)/span));
        return [0,1,2,3].map(i=>s0.c[i]+(s1.c[i]-s0.c[i])*f); };
      const hh=n=>('0'+Math.max(0,Math.min(255,Math.round(n))).toString(16)).slice(-2);
      for(let k=0;k<N;k++){
        const t0=k/N, tm=(k+0.5)/N, tt=o.flip?1-tm:tm;
        const c=sampleAt(tt); if(c[3]<0.02) continue;
        slide.addShape(pres.shapes.RECTANGLE, {x:o.g.x*IN, y:(o.g.y+o.g.h*t0)*IN, w:o.g.w*IN, h:(o.g.h/N+0.4)*IN,
          fill:{color:(hh(c[0])+hh(c[1])+hh(c[2])).toUpperCase(), transparency:Math.round((1-c[3])*100)}, line:{type:'none'}});
      }
    };
    // 배경·이미지·도형·선 = DOM(쌓임) 순서대로. 텍스트는 항상 그 위.
    const nonText=data.ops.filter(o=>o.k!=='text').sort((a,b)=>(a.order||0)-(b.order||0));
    for(const o of nonText){ if(o.k==='img') drawImg(o); else if(o.k==='bgimg') drawBgImg(o); else if(o.k==='line') drawLine(o); else if(o.k==='grad') drawGrad(o); else drawRect(o); }
    const texts=data.ops.filter(o=>o.k==='text').sort((a,b)=>(a.order||0)-(b.order||0));
    for(const o of texts){
      const arr=[];
      o.runs.forEach(r=>{
        if(r.br){ if(arr.length) arr[arr.length-1].options.breakLine=true; return; }
        let t=r.text; if(r.upper) t=t.toUpperCase();
        const op={fontFace:pptFont(r.fam,r.weight,r.italic), fontSize:Math.max(4,r.size), italic:r.italic,
          charSpacing:r.cs?Math.round(r.cs*10)/10:0};
        if(r.outline){ op.outline={size:r.outline.w, color:r.outline.color}; op.color=(data.bg||'0E0D0B'); }
        else { op.color=r.color; op.transparency=r.alpha<1?Math.round((1-r.alpha)*100):0; }
        arr.push({text:t, options:op});
      });
      if(!arr.length) continue;
      const rsizes=o.runs.filter(r=>!r.br).map(r=>r.size||0);
      const maxSize=Math.max(...rsizes);
      const minSize=Math.min(...rsizes.filter(s=>s>0));
      const uniformSize = minSize>0 && maxSize/minSize < 1.8;   // 한 줄 안 크기 편차 작음(혼합 거대 인라인 아님)
      const noWrap=maxSize>=40 || o.singleLine;
      const topt={x:o.g.x*IN, y:o.g.y*IN, w:o.g.w*IN+0.12, h:o.g.h*IN,
        align:o.align, valign:(o.valign||'top'), margin:0, wrap:!noWrap, autoFit:false};
      // 줄간격: 다줄 '본문'만 절대 pt(정확). 제목은 배수(첫 줄 위치 보존 — 줄간격 줄이면 제목 첫 줄이 위로 끌려 라벨과 겹침)
      if(!o.singleLine && o.lhPt && maxSize<=24) topt.lineSpacing=o.lhPt; else topt.lineSpacingMultiple=o.lh;
      slide.addText(arr, topt);
    }
  }

  if(N>=1){
    // 덱 모드 (section 단위) — 기존 동작 그대로
    pres.defineLayout({name:'D', width:13.333, height:7.5});
    pres.layout='D';
    for(let i=1;i<=N;i++){
      const data=await extractSlide(page, i);
      const slide=pres.addSlide();
      if(data && data.bg) slide.background={color:data.bg};
      if(!data){ console.log('  슬라이드',i,'(빈 데이터)'); continue; }
      renderSlide(slide, data);
      console.log('  슬라이드',i,'완료');
    }
  } else {
    // 문서 모드 (section 없음) — 긴 페이지를 세로 길이로 잘라 편집가능 슬라이드로
    const portrait = ORIENT==='portrait';
    const RW = portrait?1080:1920, PAGE_H = portrait?1920:1080;
    console.log('문서 모드(슬라이드 섹션 없음) — '+(portrait?'세로':'가로')+'형으로 슬라이스');
    await page.setViewport({width:RW, height:PAGE_H, deviceScaleFactor:1});
    await new Promise(r=>setTimeout(r,500));
    try{ await page.evaluate(()=>Promise.all([...document.images].filter(i=>!i.complete).map(i=>new Promise(r=>{i.onload=i.onerror=r;setTimeout(r,3000);})))); }catch(e){}
    const full = await extractSlide(page, 1, {mode:'doc', renderW:RW});
    pres.defineLayout({name:'DOC', width:portrait?7.5:13.333, height:portrait?13.333:7.5});
    pres.layout='DOC';
    const opBottom=(o)=> o.k==='line'? Math.max(o.y1,o.y2) : (o.k==='bgimg'? o.area.y+o.area.h : (o.g? o.g.y+o.g.h : 0));
    const anchorY=(o)=> o.k==='line'? Math.min(o.y1,o.y2) : (o.k==='bgimg'? o.area.y : (o.g? o.g.y : 0));
    const shift=(o,dy)=>{ const n=Object.assign({},o);
      if(o.k==='line'){ n.y1=o.y1+dy; n.y2=o.y2+dy; }
      else if(o.k==='bgimg'){ n.area=Object.assign({},o.area,{y:o.area.y+dy}); }
      else if(o.g){ n.g=Object.assign({},o.g,{y:o.g.y+dy}); }
      return n; };
    if(!full || !full.ops || !full.ops.length){
      console.log('변환할 내용을 못 찾음 (빈 슬라이드 1장 생성)');
      const s=pres.addSlide(); if(full&&full.bg) s.background={color:full.bg};
    } else {
      let H=full.height||0; for(const o of full.ops){ const b=opBottom(o); if(b>H)H=b; }
      const nS=Math.max(1, Math.ceil(H/PAGE_H));
      console.log('총 높이 '+Math.round(H)+'px → '+nS+'장');
      for(let p=0;p<nS;p++){
        const y0=p*PAGE_H, y1=y0+PAGE_H;
        const ops=full.ops.filter(o=>{ const ay=anchorY(o); return ay>=y0 && ay<y1; }).map(o=>shift(o,-y0));
        const slide=pres.addSlide();
        if(full.bg) slide.background={color:full.bg};
        renderSlide(slide, {bg:full.bg, ops});
        console.log('  슬라이드',(p+1),'완료');
      }
    }
  }
  await browser.close();
  await pres.writeFile({fileName:OUT});
  console.log('PPT 생성 완료:', OUT);
})().catch(e=>{ console.error('오류:', e.message); process.exit(1); });
