// HTML deck -> native editable PPTX
//   node html2pptx.js <input.html> [output.pptx]
// Reads computed geometry from the rendered page (puppeteer + installed Chrome)
// and rebuilds each slide as native PowerPoint text/shapes/images (pptxgenjs).
const puppeteer = require('puppeteer-core');
const PptxGenJS = require('pptxgenjs');
const path = require('path');
const fs = require('fs');

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
const CHROME = findChrome();
if(!CHROME){ console.error('Chrome/Edge를 찾을 수 없습니다. Chrome을 설치하거나 CHROME_PATH 환경변수를 설정하세요.'); process.exit(1); }

const IN = 13.333 / 1920;   // px -> inch
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

async function extractSlide(page, idx){
  return await page.evaluate((idx) => {
    const secs=[...document.querySelectorAll('section')];
    const sec = secs[idx-1];
    if(!sec) return null;
    const SR = sec.getBoundingClientRect();
    const scale = SR.width/1920;
    const rel = (el)=>{const r=el.getBoundingClientRect();return{x:(r.left-SR.left)/scale,y:(r.top-SR.top)/scale,w:r.width/scale,h:r.height/scale};};
    function col(c){
      if(!c||c==='transparent'||c==='rgba(0, 0, 0, 0)') return null;
      const m=c.match(/rgba?\(([^)]+)\)/); if(!m) return null;
      const p=m[1].split(',').map(s=>parseFloat(s));
      const h=n=>('0'+Math.round(n).toString(16)).slice(-2);
      return {hex:(h(p[0])+h(p[1])+h(p[2])).toUpperCase(), a:p[3]===undefined?1:p[3]};
    }
    const isBlock = (d)=>/(block|flex|grid|list-item|table)/.test(d) || d==='inline-block' || d==='inline-flex';
    // 부모보다 폰트가 훨씬 큰 인라인 강조(예: 작은 제목 옆 거대한 이탤릭) = 독립 박스로 분리
    const bigAccent=(el)=>false; // (비활성) 거대 인라인 분리는 상단 겹침 유발 → 되돌림
    const ops=[];
    const allEls=[...sec.querySelectorAll('*')];      // DOM(=paint) 순서
    const idxOf=(el)=>{ const i=allEls.indexOf(el); return i<0?9999:i; };

    sec.querySelectorAll('img').forEach(img=>{
      const g=rel(img); if(g.w<3||g.h<3) return;
      const cs=getComputedStyle(img);
      const filt=cs.filter||'';
      const invert=/invert\(/.test(filt);
      const fm=filt.match(/brightness\(([\d.]+)\)/);
      ops.push({k:'img', order:idxOf(img), src:img.getAttribute('src'), g, fit:cs.objectFit, op:parseFloat(cs.opacity),
        bright: invert?1:(fm?parseFloat(fm[1]):1), invert});
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
      // 얇은 선/바 (bg로 그린 divider) → 선으로
      if(g.h>0 && g.h<=4 && g.w>=8 && bg && bg.a>0.05){ ops.push({k:'line',order:ord,x1:g.x,y1:g.y+g.h/2,x2:g.x+g.w,y2:g.y+g.h/2,color:bg.hex,a:bg.a,w:Math.max(g.h,1)}); return; }
      if(g.w>0 && g.w<=4 && g.h>=8 && bg && bg.a>0.05){ ops.push({k:'line',order:ord,x1:g.x+g.w/2,y1:g.y,x2:g.x+g.w/2,y2:g.y+g.h,color:bg.hex,a:bg.a,w:Math.max(g.w,1)}); return; }
      if(g.w<4||g.h<4) return;
      // bg 채움
      if(bg && bg.a>0.02){ ops.push({k:'rect', order:ord, g, fill:{hex:bg.hex,a:bg.a}, line:null}); }
      // 테두리: 면별로 선 추출 (한쪽만 있어도 잡음)
      const W=s=>parseFloat(cs.getPropertyValue('border-'+s+'-width'))||0;
      const C=s=>col(cs.getPropertyValue('border-'+s+'-color'));
      const sides=[['top',g.x,g.y,g.x+g.w,g.y],['bottom',g.x,g.y+g.h,g.x+g.w,g.y+g.h],['left',g.x,g.y,g.x,g.y+g.h],['right',g.x+g.w,g.y,g.x+g.w,g.y+g.h]];
      for(const [s,x1,y1,x2,y2] of sides){ const w=W(s),c=C(s); if(w>0.4 && c && c.a>0.05) ops.push({k:'line',order:ord,x1,y1,x2,y2,color:c.hex,a:c.a,w}); }
      const bi=cs.backgroundImage;
      if(bi && bi!=='none' && /gradient/.test(bi) && g.w>800 && g.h>500){
        const ps=[...bi.matchAll(/rgba?\(([^)]+)\)/g)].map(m=>m[1].split(',').map(Number));
        if(ps.length){
          let r=0,gg=0,b=0,sa=0,asum=0;
          ps.forEach(p=>{const al=p[3]===undefined?1:p[3]; r+=p[0]*al; gg+=p[1]*al; b+=p[2]*al; sa+=al; asum+=al;});
          const h=n=>('0'+Math.round(n).toString(16)).slice(-2);
          ops.push({k:'rect', order:idxOf(el), g, fill:{hex:(h(r/sa)+h(gg/sa)+h(b/sa)).toUpperCase(), a: asum/ps.length}, line:null});
        }
      }
    });

    const walker=document.createTreeWalker(sec, NodeFilter.SHOW_TEXT);
    const groups=new Map(); let tn;
    while((tn=walker.nextNode())){
      if(!tn.textContent.trim()) continue;
      let blk=tn.parentElement;
      while(blk && blk!==sec){ if(isBlock(getComputedStyle(blk).display)||bigAccent(blk)) break; blk=blk.parentElement; }
      if(!blk) blk=tn.parentElement;
      if(!groups.has(blk)) groups.set(blk,[]);
      groups.get(blk).push(tn);
    }
    groups.forEach((nodes, blk)=>{
      const cs=getComputedStyle(blk);
      const g=rel(blk); if(g.w<4||g.h<4) return;
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
            else if(isBlock(getComputedStyle(ch).display)){}
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
      ops.push({k:'text', order:idxOf(blk), g,
        align: cs.textAlign==='center'?'center':(cs.textAlign==='right'?'right':'left'),
        lh, lhPt, singleLine, runs});
    });

    const secBg = col(getComputedStyle(sec).backgroundColor);
    return {bg: secBg? secBg.hex : null, ops};
  }, idx);
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
  pres.defineLayout({name:'D', width:13.333, height:7.5});
  pres.layout='D';

  for(let i=1;i<=N;i++){
    const data=await extractSlide(page, i);
    const slide=pres.addSlide();
    if(data && data.bg) slide.background={color:data.bg};
    if(!data){ console.log('  슬라이드',i,'(빈 데이터)'); continue; }
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
      }catch(e){}
    };
    const drawRect=(o)=>{
      const opt={x:o.g.x*IN, y:o.g.y*IN, w:o.g.w*IN, h:o.g.h*IN};
      if(o.fill) opt.fill={color:o.fill.hex, transparency:Math.round((1-o.fill.a)*100)}; else opt.fill={type:'none'};
      if(o.line) opt.line={color:o.line.hex, width:Math.max(0.5,o.line.w*0.75), transparency:Math.round((1-o.line.a)*100)};
      slide.addShape(pres.shapes.RECTANGLE, opt);
    };
    const drawLine=(o)=>{
      slide.addShape(pres.shapes.LINE, {x:o.x1*IN, y:o.y1*IN, w:(o.x2-o.x1)*IN, h:(o.y2-o.y1)*IN,
        line:{color:o.color, width:Math.max(0.5,o.w*0.75), transparency:Math.round((1-o.a)*100)}});
    };
    // 배경·이미지·도형·선 = DOM(쌓임) 순서대로. 텍스트는 항상 그 위.
    const nonText=data.ops.filter(o=>o.k!=='text').sort((a,b)=>(a.order||0)-(b.order||0));
    for(const o of nonText){ if(o.k==='img') drawImg(o); else if(o.k==='line') drawLine(o); else drawRect(o); }
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
      const maxSize=Math.max(...o.runs.filter(r=>!r.br).map(r=>r.size||0));
      const noWrap=maxSize>=40 || o.singleLine;
      slide.addText(arr, {x:o.g.x*IN, y:o.g.y*IN, w:o.g.w*IN+0.12, h:o.g.h*IN,
        align:o.align, valign:'top', margin:0, lineSpacingMultiple:o.lh, wrap:!noWrap, autoFit:false});
    }
    console.log('  슬라이드',i,'완료');
  }
  await browser.close();
  await pres.writeFile({fileName:OUT});
  console.log('PPT 생성 완료:', OUT);
})().catch(e=>{ console.error('오류:', e.message); process.exit(1); });
