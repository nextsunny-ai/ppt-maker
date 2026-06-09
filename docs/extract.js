// 추출 엔진 — 데스크탑 html2pptx.js의 page.evaluate 본문과 동일 (iframe 안에서 실행).
// iframe(deck HTML)에 주입되어 window.__extractSlide(idx, opts) 로 호출됨.
window.__extractSlide = function(idx, opts){
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
      try{
        _ccx.clearRect(0,0,1,1); _ccx.fillStyle='#000'; _ccx.fillStyle=c; _ccx.fillRect(0,0,1,1);
        const d=_ccx.getImageData(0,0,1,1).data; const a=d[3]/255;
        if(a<=0) return null;
        return {hex:(h(d[0])+h(d[1])+h(d[2])).toUpperCase(), a};
      }catch(e){ return null; }
    }
    const isBlock = (d)=>/(block|flex|grid|list-item|table)/.test(d) || d==='inline-block' || d==='inline-flex';
    const PHRASING=new Set(['B','STRONG','I','EM','SMALL','SPAN','A','SUP','SUB','MARK','U','CODE','ABBR','TIME','LABEL','FONT','CITE','Q','S','INS','DEL','BDI','BDO']);
    const isBlockEl=(el)=>{
      if(!isBlock(getComputedStyle(el).display)) return false;
      const par=el.parentElement;
      if(par && /(flex|grid)/.test(getComputedStyle(par).display) && PHRASING.has(el.tagName)){
        const hasText=[...par.childNodes].some(n=>n.nodeType===3 && n.textContent.trim());
        const cs=getComputedStyle(el); const cb=col(cs.backgroundColor);
        const looksChip=(parseFloat(cs.paddingLeft)||0)>2 || cs.borderLeftStyle!=='none' || (cb&&cb.a>0.02);
        if(hasText && !looksChip) return false;
      }
      return true;
    };
    const bigAccent=(el)=>false;
    const splitTop=(s)=>{ const out=[]; let d=0,cur=''; for(const ch of s){ if(ch==='(')d++; else if(ch===')')d--; if(ch===','&&d===0){out.push(cur);cur='';} else cur+=ch; } if(cur.trim())out.push(cur); return out; };
    const gradComposite=(bi)=>{
      const layers=splitTop(bi).filter(s=>/gradient/.test(s));
      if(!layers.length) return null;
      let cr=null,ca=0;
      for(let li=layers.length-1; li>=0; li--){
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
    const allEls=[...sec.querySelectorAll('*')];
    const idxOf=(el)=>{ const i=allEls.indexOf(el); return i<0?9999:i; };
    const zoomOf=(el)=>{ let e=el; for(let i=0;i<3 && e;i++){ const z=e.getAttribute&&e.getAttribute('data-ppt-zoom'); if(z){const p=z.split(','); const f=parseFloat(p[0]), d=parseFloat(p[1]); if(f>1.008) return {from:f, dur:d};} e=e.parentElement; } return null; };
    sec.querySelectorAll('img').forEach(img=>{
      const g=rel(img); if(g.w<3||g.h<3) return;
      const cs=getComputedStyle(img);
      const filt=cs.filter||'';
      const invert=/invert\(/.test(filt);
      const fm=filt.match(/brightness\(([\d.]+)\)/);
      const sm=filt.match(/saturate\(([\d.]+)\)/);
      ops.push({k:'img', order:idxOf(img), src:img.getAttribute('src'), g, fit:cs.objectFit, op:parseFloat(cs.opacity),
        bright: invert?1:(fm?parseFloat(fm[1]):1), invert, sat: sm?parseFloat(sm[1]):1, zoom:zoomOf(img)});
    });
    sec.querySelectorAll('image-slot[src], [data-img-src], picture').forEach(el=>{
      if(el.tagName==='PICTURE'){ if(el.querySelector('img')) return; }
      const src=el.getAttribute('src')||el.getAttribute('data-img-src'); if(!src) return;
      const g=rel(el); if(g.w<3||g.h<3) return;
      ops.push({k:'img', order:idxOf(el), src, g, fit:(el.getAttribute('fit')||'cover'), op:1, bright:1, invert:false, zoom:zoomOf(el)});
    });
    sec.querySelectorAll('*').forEach(el=>{
      if(el.tagName==='IMG'||el.tagName==='BR'||el.tagName==='IMAGE-SLOT') return;
      const cs=getComputedStyle(el);
      if(cs.display==='none'||cs.visibility==='hidden') return;
      const g=rel(el);
      const bg=col(cs.backgroundColor);
      const ord=idxOf(el);
      const biEl=cs.backgroundImage;
      const lineCol = (bg && bg.a>0.05) ? bg : ((biEl && biEl!=='none' && /gradient/.test(biEl)) ? gradComposite(biEl) : null);
      if(g.h>0 && g.h<=4 && g.w>=8 && lineCol && lineCol.a>0.05){ ops.push({k:'line',order:ord,x1:g.x,y1:g.y+g.h/2,x2:g.x+g.w,y2:g.y+g.h/2,color:lineCol.hex,a:lineCol.a,w:Math.max(g.h,1)}); return; }
      if(g.w>0 && g.w<=4 && g.h>=8 && lineCol && lineCol.a>0.05){ ops.push({k:'line',order:ord,x1:g.x+g.w/2,y1:g.y,x2:g.x+g.w/2,y2:g.y+g.h,color:lineCol.hex,a:lineCol.a,w:Math.max(g.w,1)}); return; }
      if(g.w<4||g.h<4) return;
      const rad=Math.min((parseFloat(cs.borderTopLeftRadius)||0)/scale, Math.min(g.w,g.h)/2);
      const W=s=>parseFloat(cs.getPropertyValue('border-'+s+'-width'))||0;
      const C=s=>col(cs.getPropertyValue('border-'+s+'-color'));
      const wT=W('top'),wB=W('bottom'),wL=W('left'),wR=W('right'); const cT=C('top');
      const uniformBorder = wT>0.4 && Math.abs(wT-wB)<0.6 && Math.abs(wT-wL)<0.6 && Math.abs(wT-wR)<0.6 && cT && cT.a>0.05;
      if(uniformBorder && rad>2){
        ops.push({k:'rect', order:ord, g, fill:(bg&&bg.a>0.02?{hex:bg.hex,a:bg.a}:null), line:{hex:cT.hex,a:cT.a,w:wT}, rad});
      } else {
        if(bg && bg.a>0.02){ ops.push({k:'rect', order:ord, g, fill:{hex:bg.hex,a:bg.a}, line:null, rad:rad>2?rad:0}); }
        const sides=[['top',g.x,g.y,g.x+g.w,g.y],['bottom',g.x,g.y+g.h,g.x+g.w,g.y+g.h],['left',g.x,g.y,g.x,g.y+g.h],['right',g.x+g.w,g.y,g.x+g.w,g.y+g.h]];
        for(const [s,x1,y1,x2,y2] of sides){ const w=W(s),c=C(s); if(w>0.4 && c && c.a>0.05) ops.push({k:'line',order:ord,x1,y1,x2,y2,color:c.hex,a:c.a,w}); }
      }
      const bi=cs.backgroundImage;
      if(bi && bi!=='none' && !/gradient/.test(bi)){
        const um=bi.match(/url\((['"]?)([^'")]+)\1\)/);
        if(um){
          const bl=parseFloat(cs.borderLeftWidth)||0, bt=parseFloat(cs.borderTopWidth)||0, br=parseFloat(cs.borderRightWidth)||0, bb=parseFloat(cs.borderBottomWidth)||0;
          const area={x:g.x+bl/scale, y:g.y+bt/scale, w:g.w-(bl+br)/scale, h:g.h-(bt+bb)/scale};
          ops.push({k:'bgimg', order:ord, area, src:um[2], size:cs.backgroundSize, posx:cs.backgroundPositionX, posy:cs.backgroundPositionY, zoom:zoomOf(el)});
        }
      }
      if(bi && bi!=='none' && /gradient/.test(bi) && g.w>800 && g.h>500){
        const go=makeGradOp(bi, g, ord); if(go) ops.push(go);
      }
      for(const pe of ['::before','::after']){
        let cs2; try{ cs2=getComputedStyle(el,pe); }catch(e){ continue; }
        if(!cs2 || cs2.content==='none' || cs2.content==='') continue;
        if(!/absolute|fixed/.test(cs2.position)) continue;
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
      let l=1e9,t=1e9,r=-1e9,bt=-1e9,ok=false;
      for(const n of nodes){ const rg=document.createRange(); rg.selectNodeContents(n);
        for(const c of rg.getClientRects()){ if(c.width<=0||c.height<=0) continue; ok=true; l=Math.min(l,c.left); t=Math.min(t,c.top); r=Math.max(r,c.right); bt=Math.max(bt,c.bottom); } }
      let vAlign='top';
      const fsBlk=parseFloat(cs.fontSize)||0;
      let txRect=null;
      if(ok){
        txRect={x:(l-SR.left)/scale, y:(t-SR.top)/scale, w:(r-l)/scale, h:(bt-t)/scale};
        if(hasBlockChild && fsBlk<48) g=txRect;
        else if(!hasBlockChild && fsBlk<48 && txRect.h > g.h*1.5){ g={x:g.x, y:txRect.y, w:g.w, h:txRect.h}; }
      }
      const runs=[];
      function walk(node){
        node.childNodes.forEach(ch=>{
          if(ch.nodeType===3){
            const t=ch.textContent.replace(/\s+/g,' ');
            if(!t.trim() && t!==' ') return;
            const es=getComputedStyle(ch.parentElement);
            const tsw=parseFloat(es.getPropertyValue('-webkit-text-stroke-width'))||0;
            const tsc=col(es.getPropertyValue('-webkit-text-stroke-color'));
            const fillCol=col(es.getPropertyValue('-webkit-text-fill-color'));
            const hollow = tsw>0 && (!fillCol || fillCol.a<0.05);
            let c = fillCol || col(es.color);
            if(!hollow && (!c || c.a<0.05)){
              const clip=(es.getPropertyValue('-webkit-background-clip')||es.getPropertyValue('background-clip')||'');
              if(/text/.test(clip)){ const gc=gradComposite(es.backgroundImage); if(gc && gc.a>0.05) c=gc; }
            }
            if(!c) c={hex:'FAF8F4',a:1};
            runs.push({text:t, fam:es.fontFamily, weight:parseInt(es.fontWeight)||400,
              size:Math.round(parseFloat(es.fontSize)*0.5*10)/10, italic:es.fontStyle==='italic',
              color:c.hex, alpha:c.a,
              outline: hollow? {w:Math.max(0.5,tsw*0.5), color:(tsc?tsc.hex:'FAF8F4')} : null,
              cs:(parseFloat(es.letterSpacing)||0)*0.5, upper:es.textTransform==='uppercase'});
          } else if(ch.nodeType===1){
            if(ch.tagName==='BR'){ runs.push({br:true}); }
            else if(ch.tagName==='IMG'){}
            else if(isBlockEl(ch)){}
            else if(bigAccent(ch)){}
            else walk(ch);
          }
        });
      }
      walk(blk);
      if(!runs.length) return;
      const lhpx=parseFloat(cs.lineHeight);
      const lh = lhpx/parseFloat(cs.fontSize) || 1.2;
      const lhPt = isNaN(lhpx)? null : Math.round(lhpx*0.5*10)/10;
      const hasBr=runs.some(r=>r.br);
      const maxPx=Math.max(...runs.filter(r=>!r.br).map(r=>(r.size||0)*2));
      let nLines=1;
      try{
        const rg=document.createRange(); rg.selectNodeContents(blk);
        const tops=[...rg.getClientRects()].filter(c=>c.width>0&&c.height>0).map(c=>c.top).sort((a,b)=>a-b);
        if(tops.length){ let prev=tops[0]; const gap=Math.max(6,(lhpx||16)*0.5);
          for(const t of tops){ if(t-prev>gap){ nLines++; prev=t; } } }
      }catch(e){}
      const singleLine = !hasBr && nLines<=1;
      if(txRect && fsBlk>=48 && nLines>1 && !hasBr && !hasBlockChild){ g={x:g.x, y:txRect.y, w:g.w, h:txRect.h}; vAlign='middle'; }
      const pL=(parseFloat(cs.paddingLeft)||0)/scale, pT=(parseFloat(cs.paddingTop)||0)/scale,
            pR=(parseFloat(cs.paddingRight)||0)/scale, pB=(parseFloat(cs.paddingBottom)||0)/scale;
      const gt = (!hasBlockChild && vAlign==='top' && (pL||pT||pR||pB))? {x:g.x+pL, y:g.y+pT, w:Math.max(4,g.w-pL-pR), h:Math.max(4,g.h-pT-pB)} : g;
      ops.push({k:'text', order:idxOf(blk), g:gt, valign:vAlign,
        align: cs.textAlign==='center'?'center':(cs.textAlign==='right'?'right':'left'),
        lh, lhPt, singleLine, runs});
    });
    let secBg = col(getComputedStyle(sec).backgroundColor);
    if(!secBg || secBg.a<0.02){
      let el=sec.parentElement;
      while(el){ const b=col(getComputedStyle(el).backgroundColor); if(b && b.a>0.02){ secBg=b; break; } el=el.parentElement; }
      if(!secBg || secBg.a<0.02){ const hb=col(getComputedStyle(document.documentElement).backgroundColor); if(hb && hb.a>0.02) secBg=hb; }
    }
    return {bg: secBg? secBg.hex : null, ops, height: SR.height/scale};
};

// 줌 감지 (활성화 강제 전) — data-ppt-zoom 태깅. 데스크탑 엔진과 동일 로직.
window.__detectZoom = function(){
  const scaleOf=(el)=>{ const t=getComputedStyle(el).transform; const m=(t||'').match(/matrix\(([^)]+)\)/); if(m){const v=m[1].split(',').map(parseFloat); return +Math.hypot(v[0],v[1]).toFixed(4);} const m3=(t||'').match(/matrix3d\(([^)]+)\)/); if(m3){const v=m3[1].split(',').map(parseFloat); return +Math.hypot(v[0],v[1]).toFixed(4);} return 1; };
  const durOf=(el)=>{ const cs=getComputedStyle(el); let d=0; const tp=(cs.transitionProperty||'').split(',').map(s=>s.trim()); (cs.transitionDuration||'').split(',').forEach((s,i)=>{ const p=tp[i]||tp[0]||''; if(p==='transform'||p==='all') d=Math.max(d,parseFloat(s)||0); }); if((cs.animationName||'none')!=='none'){ (cs.animationDuration||'').split(',').forEach(s=>d=Math.max(d,parseFloat(s)||0)); } return d; };
  const cands=[...document.querySelectorAll('img,[class*="bg"],[class*="hero"],[class*="cover"],[style*="background-image"]')];
  const info=new Map();
  cands.forEach(el=>{ const d=durOf(el); if(d>=0.3) info.set(el,d); });
  if(!info.size) return 0;
  const tmp=document.createElement('style'); tmp.textContent='*{transition:none!important;animation:none!important}'; document.head.appendChild(tmp);
  const re=[]; document.querySelectorAll('section,.slide,[class*="slide"]').forEach(s=>['active','current','is-active','is-current','is-visible','visible','shown','seen','revealed'].forEach(c=>{ if(s.classList.contains(c)){s.classList.remove(c);re.push([s,c]);} }));
  void document.body.offsetHeight;
  let n=0;
  info.forEach((d,el)=>{ const s=scaleOf(el); if(s>1.008 && s<1.6){ el.setAttribute('data-ppt-zoom', s.toFixed(4)+','+d.toFixed(2)); n++; } });
  re.forEach(([s,c])=>s.classList.add(c)); tmp.remove();
  return n;
};

// reveal 강제 (모든 슬라이드 활성+최종상태). 데스크탑 엔진과 동일.
window.__forceReveal = function(){
  const cls=['active','current','is-active','is-current','is-visible','visible','in-view','inview','show','shown','revealed','seen','animated','aos-animate'];
  document.querySelectorAll('section, .slide, [class*="slide"], [class*="page"]').forEach(s=>cls.forEach(c=>s.classList.add(c)));
  const st=document.createElement('style');
  st.textContent='[class~="r"],[class~="r1"],[class~="r2"],[class~="r3"],[class*="reveal"],[class*="fade"],[class*="anim"],[data-reveal],[data-aos]{opacity:1!important;transform:none!important;filter:none!important;visibility:visible!important;clip-path:none!important}';
  document.head.appendChild(st);
};
window.__sectionCount = function(){ return document.querySelectorAll('section').length; };
window.__docHeight = function(){ return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight); };
