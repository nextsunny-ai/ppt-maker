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
  if(/Spectral|Cormorant/i.test(fam)) return pick('Spectral',[[0,' Light'],[400,''],[500,' Medium']]);
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
    const ops=[];

    sec.querySelectorAll('img').forEach(img=>{
      const g=rel(img); if(g.w<3||g.h<3) return;
      const cs=getComputedStyle(img);
      const filt=cs.filter||'';
      const invert=/invert\(/.test(filt);
      const fm=filt.match(/brightness\(([\d.]+)\)/);
      ops.push({k:'img', src:img.getAttribute('src'), g, fit:cs.objectFit, op:parseFloat(cs.opacity),
        bright: invert?1:(fm?parseFloat(fm[1]):1), invert});
    });

    sec.querySelectorAll('*').forEach(el=>{
      if(el.tagName==='IMG'||el.tagName==='BR') return;
      const cs=getComputedStyle(el);
      if(cs.display==='none'||cs.visibility==='hidden') return;
      const g=rel(el); if(g.w<4||g.h<4) return;
      const bg=col(cs.backgroundColor);
      const bw=parseFloat(cs.borderTopWidth)||0;
      const bc=col(cs.borderTopColor);
      const hasBorder = bw>0 && bc && bc.a>0.02 &&
        (parseFloat(cs.borderRightWidth)||0)>0 && (parseFloat(cs.borderBottomWidth)||0)>0 && (parseFloat(cs.borderLeftWidth)||0)>0;
      const hasBg = bg && bg.a>0.02;
      if(hasBorder||hasBg){
        ops.push({k:'rect', g,
          fill: hasBg? {hex:bg.hex, a:bg.a} : null,
          line: hasBorder? {hex:bc.hex, a:bc.a, w:bw} : null});
      }
      const bi=cs.backgroundImage;
      if(bi && bi!=='none' && /gradient/.test(bi) && g.w>800 && g.h>500){
        const ps=[...bi.matchAll(/rgba?\(([^)]+)\)/g)].map(m=>m[1].split(',').map(Number));
        if(ps.length){
          let r=0,gg=0,b=0,sa=0,asum=0;
          ps.forEach(p=>{const al=p[3]===undefined?1:p[3]; r+=p[0]*al; gg+=p[1]*al; b+=p[2]*al; sa+=al; asum+=al;});
          const h=n=>('0'+Math.round(n).toString(16)).slice(-2);
          ops.push({k:'rect', g, fill:{hex:(h(r/sa)+h(gg/sa)+h(b/sa)).toUpperCase(), a: asum/ps.length}, line:null});
        }
      }
    });

    const walker=document.createTreeWalker(sec, NodeFilter.SHOW_TEXT);
    const groups=new Map(); let tn;
    while((tn=walker.nextNode())){
      if(!tn.textContent.trim()) continue;
      let blk=tn.parentElement;
      while(blk && blk!==sec){ if(isBlock(getComputedStyle(blk).display)) break; blk=blk.parentElement; }
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
            else walk(ch);
          }
        });
      }
      walk(blk);
      if(!runs.length) return;
      ops.push({k:'text', g,
        align: cs.textAlign==='center'?'center':(cs.textAlign==='right'?'right':'left'),
        lh: parseFloat(cs.lineHeight)/parseFloat(cs.fontSize) || 1.2, runs});
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
  await page.goto(fileUrl(HTML), {waitUntil:'networkidle0'});
  await page.evaluate(()=>{ const d=document.querySelector('deck-stage'); if(d) d.setAttribute('noscale',''); });
  await page.emulateMediaType('print');
  await new Promise(r=>setTimeout(r,1500));

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
    const imgs=data.ops.filter(o=>o.k==='img');
    const rects=data.ops.filter(o=>o.k==='rect').sort((a,b)=>(b.g.w*b.g.h)-(a.g.w*a.g.h));
    const texts=data.ops.filter(o=>o.k==='text');
    for(const o of imgs){
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
    }
    for(const o of rects){
      const opt={x:o.g.x*IN, y:o.g.y*IN, w:o.g.w*IN, h:o.g.h*IN};
      if(o.fill) opt.fill={color:o.fill.hex, transparency:Math.round((1-o.fill.a)*100)}; else opt.fill={type:'none'};
      if(o.line) opt.line={color:o.line.hex, width:Math.max(0.5,o.line.w*0.75), transparency:Math.round((1-o.line.a)*100)};
      slide.addShape(pres.shapes.RECTANGLE, opt);
    }
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
      const noWrap=maxSize>=40;
      slide.addText(arr, {x:o.g.x*IN, y:o.g.y*IN, w:o.g.w*IN+0.12, h:o.g.h*IN,
        align:o.align, valign:'top', margin:0, lineSpacingMultiple:o.lh, wrap:!noWrap, autoFit:false});
    }
    console.log('  슬라이드',i,'완료');
  }
  await browser.close();
  await pres.writeFile({fileName:OUT});
  console.log('PPT 생성 완료:', OUT);
})().catch(e=>{ console.error('오류:', e.message); process.exit(1); });
