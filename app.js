/* ===========================================================================
   SP Group · Excel-to-Slide Studio — AI edition (studio UI)
   Flow:  Upload ▶ Studio (config lives in a popover; AI edits the deck-spec)
   AI authors the deck-spec (single source of truth); a deterministic engine
   renders preview + .pptx. In-browser engine by default; Azure OpenAI-ready.
   =========================================================================== */
"use strict";

const TEMPLATES = {
  sp:{primary:"0097A7",accent:"FFAB40",dark:"282F45",name:"SP Group"},
  light:{primary:"1C7293",accent:"00A896",dark:"21295C",name:"SP Group"},
  exec:{primary:"3454D1",accent:"CADCFC",dark:"1E2761",name:"SP Group"},
  energy:{primary:"2C5F2D",accent:"97BC62",dark:"14331A",name:"SP Group"},
};
const COLORNAMES={red:"D7263D",orange:"F06000",amber:"E8A317",gold:"C9A227",green:"2E9E5B",teal:"00838A",
  cyan:"0FB5C4",blue:"1F6FEB",navy:"12203A",purple:"6D2E9E",magenta:"B5179E",grey:"5C5C5C",gray:"5C5C5C",black:"12203A",lime:"8DC63F"};

const SAMPLE={
  fy:["FY25/26","FY26/27","FY27/28","FY28/29","FY29/30","FY30/31","FY31/32"],
  roeByFy:{"FY27/28":8.4,"FY28/29":9.1,"FY29/30":9.7,"FY30/31":10.3,"FY31/32":10.8},
  series:{
    npat:{title:"NPAT by Business Unit (S$M)",cats:["SPPG","SES","SPS","SPGT","SGSPAA","Others*"],vals:[550,400,100,60,100,50]},
    ebitda:{title:"EBITDA by Business Unit (S$M)",cats:["SPPG","SES","SPS","SPGT","SGSPAA","Others*"],vals:[600,80,80,100,100,40]},
    revenue:{title:"Revenue by Business Unit (S$M)",cats:["SPPG","SES","SPS","Others"],vals:[3000,500,400,100]},
    growth:{title:"Group revenue growth to 2031 (S$M)",cats:["FY24/25","FY25/26","FY26/27","FY27/28","FY28/29","FY29/30","FY30/31"],vals:[3500,3600,3800,4000,4800,5400,6500]},
    roe:{title:"Return on Equity trend (%)",cats:["FY27/28","FY28/29","FY29/30","FY30/31","FY31/32"],vals:[8.4,9.1,9.7,10.3,10.8]},
  },
  kfi:[
    {metric:"ROE",values:[null,null,8.4,9.1,9.7,10.3,10.8],unit:"%"},
    {metric:"ROIC",values:[null,null,6.8,7.1,7.5,7.9,8.2],unit:"%"},
    {metric:"ROTA",values:[null,null,null,null,null,null,null]},
    {metric:"3-Year Annualised TSR",values:[null,null,null,null,null,null,null]},
    {metric:"Gearing",values:[null,null,38.0,37.2,36.5,35.1,34.0],unit:"%"},
    {metric:"FFO Interest Cover",values:[null,null,4.2,4.4,4.6,4.9,5.1]},
    {metric:"FFO to Total Debt",values:[null,null,null,null,null,null,null]},
    {metric:"Moody's Standalone rating",values:[null,null,"a3","a2","a2","a2","a2"]},
    {metric:"Moody's Overall rating",values:[null,null,"Aa2","Aa1","Aa1","Aa1","Aa1"]},
    {metric:"S&P Net Debt / EBITDA",values:[null,null,null,null,null,null,null]},
  ],
};
const METRIC_ALIASES={npat:["npat","net profit","profit after tax","earnings"],ebitda:["ebitda"],
  revenue:["revenue","turnover","sales","topline","top line"],growth:["growth","2031","outlook","trajectory"],
  roe:["roe","return on equity","returns"]};

let STATE={wbName:null,isSample:false,workbook:null,spec:null,slides:[],cur:0,charts:[],history:[],aiEndpoint:""};

const $=id=>document.getElementById(id);
const show=(el,on=true)=>el.classList.toggle("hidden",!on);
function overlay(on,m="Working…"){$("ovmsg").textContent=m;show($("overlay"),on);}
const num=v=>typeof v==="number"&&isFinite(v);
const esc=s=>String(s).replace(/[&<>]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[c]));
let _uid=0; const uid=()=>"s"+(++_uid);

/* ---------- intake ---------- */
const drop=$("drop"),fileInput=$("file");
drop.addEventListener("click",()=>fileInput.click());
["dragover","dragenter"].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add("drag");}));
["dragleave","drop"].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.remove("drag");}));
drop.addEventListener("drop",ev=>{if(ev.dataTransfer.files[0])handleFile(ev.dataTransfer.files[0]);});
fileInput.addEventListener("change",()=>{if(fileInput.files[0])handleFile(fileInput.files[0]);});
$("sampleBtn").addEventListener("click",()=>{STATE.isSample=true;STATE.wbName="Sample data";STATE.workbook=sampleCatalog();launch();});
function handleFile(f){STATE.wbName=f.name;overlay(true,"Reading workbook…");const r=new FileReader();
  r.onload=e=>{try{const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});STATE.workbook=parseWorkbook(wb);STATE.isSample=false;launch();}
    catch(err){overlay(false);alert("Could not read this file: "+err.message);}};
  r.readAsArrayBuffer(f);}

function launch(){
  overlay(true,"Generating your deck…");
  try{const ing=ingestWorkbook(STATE.wbName,STATE.workbook,STATE.isSample);STATE.seriesId=ing.seriesId;STATE.versionId=ing.versionId;STATE.deckId=null;}catch(e){}
  const cfg={cutoff:defaultCutoff(),roe:Object.values(SAMPLE.roeByFy)[0],hideZero:true,lockKpi:true,
    selected:["title","kpi","npat","growth","kfi"]};
  STATE.spec=buildSpec(cfg); STATE.cur=0; STATE.history=[];
  setTimeout(()=>{
    show($("gate"),false); show($("workspace"),false); show($("intake"),false); show($("studio"),true);
    STATE.deckId=null; saveCurrentDeck(true); updateSrcLabel();
    initSettingsUI(cfg); syncBrandBar(); renderChips();
    requestAnimationFrame(()=>{ buildSlides();
      $("log").innerHTML=""; logMsg("a","Hi! I'm your Deck Copilot. Your starter deck is saved to the workspace — tell me what to change (slides, charts, data sources, template or branding), or tap a suggestion.");
      overlay(false); });
  },350);
}
function defaultCutoff(){return "2026-06-30";}

/* ---------- workbook parsing ---------- */
function cleanCell(v){if(v==null)return null;if(typeof v==="string"){const s=v.trim();return(s===""||s.startsWith("#"))?null:s;}if(typeof v==="number")return Math.round(v*1e4)/1e4;return v;}
function sampleCatalog(){return{notes:["Using illustrative sample figures."],sheets:[{name:"Sample",headers:SAMPLE.fy,grid:[]}],kfi:SAMPLE.kfi,fy:SAMPLE.fy};}
function parseWorkbook(wb){
  const cat={notes:[],sheets:[],kfi:SAMPLE.kfi,fy:SAMPLE.fy};
  wb.SheetNames.forEach(n=>{const grid=XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:true}).slice(0,60).map(r=>(r||[]).slice(0,15).map(cleanCell));
    const headers=(grid[2]||grid[0]||[]).map(c=>c==null?"":String(c));cat.sheets.push({name:n,headers,grid});});
  const kfiSheet=wb.SheetNames.find(n=>n.trim().toUpperCase()==="KFI");
  if(kfiSheet){const g=XLSX.utils.sheet_to_json(wb.Sheets[kfiSheet],{header:1,raw:true});
    const hdr=(g[2]||[]).slice(1,8).map(cleanCell).filter(Boolean);if(hdr.length)cat.fy=hdr;
    const rows=[];for(let i=3;i<23&&i<g.length;i++){const r=g[i]||[];const nm=cleanCell(r[0]);if(!nm)continue;
      rows.push({metric:String(nm),values:r.slice(1,8).map(cleanCell),unit:/roe|roic|rota|gearing|margin/i.test(nm)?"%":undefined});}
    if(rows.length){cat.kfi=rows;cat.notes.push(`Read ${rows.length} KFI rows from "${kfiSheet}".`);}
    if(!rows.some(r=>r.values.some(num)))cat.notes.push("Source ratio cells are blank/#DIV/0! — charts use illustrative figures.");
  }else cat.notes.push('No "KFI" sheet found — using illustrative figures.');
  return cat;}

/* ---------- slide factories ---------- */
function chartSlide(metric,type){const s=SAMPLE.series[metric]||SAMPLE.series.npat;
  return{id:uid(),type:"chart",title:s.title.replace(/ \(.*\)$/,""),
    chartType:type||(metric==="npat"?"doughnut":metric==="roe"?"line":"bar"),seriesName:metric.toUpperCase(),
    data:{categories:[...s.cats],values:[...s.vals]},source:{kind:"sample",metric}};}
function kpiSlide(spec){const r=(spec.keyFigures&&spec.keyFigures.roeByFy)||{};const ent=Object.entries(r).filter(([k,v])=>num(v));
  const head=spec.keyFigures?spec.keyFigures.headlineRoe:null;
  const cards=[{label:"Headline ROE",value:(num(head)?head:"–")+"%",note:ent.length?ent[0][0]:"",role:"kf:headlineRoe"}];
  if(ent.length>=2){const mid=ent[Math.floor((ent.length-1)/2)],last=ent[ent.length-1];
    cards.push({label:"ROE",value:mid[1]+"%",note:mid[0],role:"kf:roe@"+mid[0]});
    cards.push({label:"ROE",value:last[1]+"%",note:last[0],role:"kf:roe@"+last[0]});}
  cards.push({label:"Group NPAT",value:"S$1,260M",note:"illustrative"});
  return{id:uid(),type:"kpi",title:"Key figures",cards};}
function tableSlide(){return{id:uid(),type:"table",title:"Key Financial Indicators",
  fy:(STATE.workbook?STATE.workbook.fy:SAMPLE.fy),rows:JSON.parse(JSON.stringify(STATE.workbook?STATE.workbook.kfi:SAMPLE.kfi)),
  hideZero:true,source:{kind:STATE.workbook&&!STATE.isSample?"sheet":"sample",sheet:"KFI"}};}
function titleSlide(){return{id:uid(),type:"title",title:"Group Budget Deck",subtitle:"Auto-generated from the consolidation workbook"};}
function textSlide(t,bullets){return{id:uid(),type:"text",title:t||"Notes",bullets:bullets||["Add your narrative here."]};}

function buildSpec(cfg){
  const snap=snapFromCatalog(STATE.workbook);
  const roeByFy=(snap&&Object.keys(snap.roeByFy).length)?snap.roeByFy:SAMPLE.roeByFy;
  const headlineRoe=(snap&&num(snap.headlineRoe))?snap.headlineRoe:parseFloat(cfg.roe);const tpl=TEMPLATES.sp;
  const spec={meta:{source:STATE.wbName,cutoff:cfg.cutoff,currency:"S$M",generated:new Date().toISOString().slice(0,16).replace("T"," ")},
    theme:{primary:tpl.primary,accent:tpl.accent,dark:tpl.dark,brandName:"SP GROUP",template:"sp"},
    keyFigures:{headlineRoe,roeByFy},hideZero:cfg.hideZero,lockKpi:cfg.lockKpi,slides:[]};
  const map={title:()=>titleSlide(),kpi:()=>kpiSlide(spec),npat:()=>chartSlide("npat","doughnut"),growth:()=>chartSlide("growth","bar"),kfi:()=>tableSlide()};
  cfg.selected.forEach(id=>{if(map[id])spec.slides.push(map[id]());});
  return spec;}

/* ---------- rendering ---------- */
function fmt(v,unit){if(v==null)return "–";let s=num(v)?v.toLocaleString(undefined,{maximumFractionDigits:1}):String(v);if(unit==="%"&&num(v))s+="%";return s;}
function palette(theme){return["#"+theme.primary,"#"+theme.accent,"#"+theme.dark,"#78909C","#80DEEA","#FFD180","#B0BEC5","#546E7A"];}

function buildSlides(){
  STATE.charts.forEach(c=>{try{c.destroy()}catch(e){}});STATE.charts=[];const spec=STATE.spec;
  STATE.slides=spec.slides.map((sl,i)=>{const el=document.createElement("div");el.className="slide";renderSlide(el,sl,spec,i);return{slide:sl,el,label:(i+1)+". "+(sl.title||sl.type)};});
  if(STATE.cur>=STATE.slides.length)STATE.cur=Math.max(0,STATE.slides.length-1);
  renderThumbs();showSlide(STATE.cur);renderInspector();}
const LOGO="data:image/png;base64,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"; /* SP group logo data URI, injected at build */
function logoTag(){return `<img src="${LOGO}" style="position:absolute;top:40px;right:60px;height:60px" alt="SP group">`;}
function headerEl(el,sp,title,sub){const theme=sp.theme;el.insertAdjacentHTML("beforeend",logoTag());
  const h=document.createElement("div");h.style.cssText="position:absolute;left:70px;top:60px;width:980px;font-size:32px;font-weight:700;color:#"+theme.dark;h.textContent=title||"";el.appendChild(h);
  el.insertAdjacentHTML("beforeend",`<div style="position:absolute;left:70px;top:118px;width:1140px;height:3px;background:#${theme.primary}"></div>`);
  if(sub){const s=document.createElement("div");s.style.cssText="position:absolute;left:70px;top:128px;font-size:15px;color:#5C5C5C";s.textContent=sub;el.appendChild(s);}}
function footEl(el,sp){const f=document.createElement("div");f.style.cssText="position:absolute;left:70px;bottom:34px;font-size:11px;color:#8A939B";
  f.textContent=`Auto-generated from ${sp.meta.source} · cut-off ${sp.meta.cutoff} · figures in ${sp.meta.currency} · DRAFT`;el.appendChild(f);}

function renderSlide(el,sl,sp){
  const theme=sp.theme;
  if(sl.type==="title"){el.style.background="#fff";
    el.innerHTML=`${logoTag()}
      <div style="position:absolute;left:90px;top:300px;width:1010px;font-size:50px;font-weight:700;color:#${theme.dark}">${esc(sl.title)}</div>
      <div style="position:absolute;left:92px;top:392px;width:440px;height:5px;background:#${theme.primary}"></div>
      <div style="position:absolute;left:90px;top:412px;font-size:24px;color:#${theme.primary}">${esc(sl.subtitle||"")}</div>
      <div style="position:absolute;left:90px;top:640px;font-size:14px;color:#8A939B">Cut-off ${sp.meta.cutoff} · generated ${sp.meta.generated}</div>`;return;}
  headerEl(el,sp,sl.title,sl.subtitle);
  if(sl.type==="kpi"){let x=70;const w=275;(sl.cards||[]).slice(0,4).forEach(c=>{const d=document.createElement("div");
    d.style.cssText=`position:absolute;left:${x}px;top:250px;width:${w}px;height:200px;background:#E8F6F8;border:1.5px solid #${theme.primary};border-radius:10px;padding:22px`;
    d.innerHTML=`<div style="color:#${theme.primary};font-weight:700;font-size:18px">${esc(c.label)}</div>
      <div style="color:#${theme.dark};font-weight:700;font-size:48px;margin-top:8px">${esc(c.value)}</div>
      <div style="color:#5C5C5C;font-size:14px;margin-top:6px">${esc(c.note||"")}</div>`;el.appendChild(d);x+=w+22;});}
  else if(sl.type==="chart"){const cv=document.createElement("canvas");cv.style.cssText="position:absolute;left:90px;top:230px;width:1100px;height:420px";cv.width=1100;cv.height=420;el.appendChild(cv);
    const d=sl.data,pal=palette(theme);
    requestAnimationFrame(()=>{try{const ch=new Chart(cv,{type:sl.chartType==="column"?"bar":sl.chartType,
      data:{labels:d.categories,datasets:[{data:d.values,label:sl.seriesName||"",
        backgroundColor:(sl.chartType==="doughnut"||sl.chartType==="pie")?pal:"#"+theme.primary,
        borderColor:"#"+theme.primary,borderWidth:sl.chartType==="line"?3:0,fill:false,tension:.3,borderRadius:sl.chartType==="bar"?4:0}]},
      options:{responsive:false,animation:false,cutout:sl.chartType==="doughnut"?"55%":undefined,
        plugins:{legend:{display:sl.chartType==="doughnut"||sl.chartType==="pie",position:"right",labels:{font:{size:15}}},tooltip:{enabled:false}},
        scales:(sl.chartType==="doughnut"||sl.chartType==="pie")?{}:{y:{ticks:{font:{size:13}}},x:{ticks:{font:{size:13}}}}}});STATE.charts.push(ch);}catch(e){}});}
  else if(sl.type==="table"){const sub=document.createElement("div");sub.style.cssText="position:absolute;left:70px;top:128px;font-size:14px;color:#5C5C5C";
    const cm=sp._changeMap||{};const hasCh=Object.keys(cm).length>0;
    sub.textContent="Auto-extracted · #DIV/0! cleaned · "+(sl.hideZero?"zero/empty rows hidden":"all rows shown")+(hasCh?" · changed cells highlighted":"");el.appendChild(sub);
    let rows=sl.rows.slice();if(sl.hideZero)rows=rows.filter(r=>!r.values.every(v=>v==null||v===0));rows=rows.slice(0,11);
    let h=`<table style="position:absolute;left:70px;top:235px;width:1140px;border-collapse:collapse;font-size:17px"><tr style="background:#${theme.primary};color:#fff"><th style="text-align:left;padding:9px 12px">Metric</th>`;
    sl.fy.forEach(x=>h+=`<th style="padding:9px 8px">${esc(x||"")}</th>`);h+="</tr>";
    rows.forEach((r,i)=>{h+=`<tr style="background:${i%2?"#E1F4F6":"#fff"}"><td style="padding:8px 12px;color:#${theme.dark}">${esc(r.metric.slice(0,34))}</td>`;
      r.values.forEach((v,ci)=>{const ch=cm["kfi:"+r.metric+"@"+ci];const cell=ch?`background:${ch.dir==="up"?"#E6F4EA":"#FCE8E8"};font-weight:700`:"";
        const arw=ch?(ch.dir==="up"?' <span style="color:#1E8E3E">▲</span>':' <span style="color:#C5221F">▼</span>'):"";
        h+=`<td style="padding:8px 8px;text-align:center;color:#${theme.dark};${cell}">${esc(fmt(v,r.unit))}${arw}</td>`;});h+="</tr>";});
    h+="</table>";el.insertAdjacentHTML("beforeend",h);}
  else if(sl.type==="changes"){const sub=document.createElement("div");sub.style.cssText="position:absolute;left:70px;top:128px;font-size:14px;color:#5C5C5C";
    sub.textContent=sl.note||"Material movements versus the prior version";el.appendChild(sub);
    const list=(sl.changes||[]).slice(0,11);
    if(!list.length){el.insertAdjacentHTML("beforeend",`<div style="position:absolute;left:70px;top:300px;font-size:22px;color:#5C5C5C">No material changes versus the prior version.</div>`);}
    else{let h=`<table style="position:absolute;left:70px;top:235px;width:1140px;border-collapse:collapse;font-size:17px"><tr style="background:#${theme.primary};color:#fff"><th style="text-align:left;padding:9px 12px">Figure</th><th style="padding:9px 8px">Prior</th><th style="padding:9px 8px">New</th><th style="padding:9px 8px">Change</th></tr>`;
      list.forEach((c,i)=>{const up=c.dir==="up";h+=`<tr style="background:${i%2?"#E1F4F6":"#fff"}"><td style="padding:8px 12px;color:#${theme.dark}">${esc(c.label)}</td>
        <td style="padding:8px;text-align:center;color:#5C5C5C">${esc(c.oldText)}</td>
        <td style="padding:8px;text-align:center;color:#${theme.dark};font-weight:700">${esc(c.newText)}</td>
        <td style="padding:8px;text-align:center;font-weight:700;color:${up?"#1E8E3E":"#C5221F"}">${up?"▲":"▼"} ${esc(c.pctText)}</td></tr>`;});
      h+="</table>";el.insertAdjacentHTML("beforeend",h);}}
  else if(sl.type==="text"){let y=250;(sl.bullets||[]).forEach(b=>{const d=document.createElement("div");
    d.style.cssText=`position:absolute;left:90px;top:${y}px;width:1100px;font-size:24px;color:#${theme.dark}`;
    d.innerHTML=`<span style="color:#${theme.accent};font-weight:700">•</span> ${esc(b)}`;el.appendChild(d);y+=58;});}
  footEl(el,sp);}

const THUMB_W=94;
function scaleSlide(el){const w=$("frame").clientWidth;el.style.transform=`scale(${w/1280})`;}
function showSlide(i){if(!STATE.slides.length){$("frame").innerHTML="";$("pos").textContent="0 / 0";return;}
  i=Math.max(0,Math.min(i,STATE.slides.length-1));STATE.cur=i;
  const f=$("frame");f.innerHTML="";const el=STATE.slides[i].el;el.style.transform="";f.appendChild(el);scaleSlide(el);
  $("pos").textContent=`${i+1} / ${STATE.slides.length}`;
  document.querySelectorAll(".thumb").forEach((t,j)=>t.classList.toggle("active",j===i));renderInspector();}
function renderThumbs(){const wrap=$("thumbs");wrap.innerHTML="";
  STATE.slides.forEach((s,i)=>{const t=document.createElement("div");t.className="thumb";
    t.innerHTML=`<div class="num">${i+1}</div><div class="mini"></div>`;
    t.addEventListener("click",()=>showSlide(i));wrap.appendChild(t);
    const m=t.querySelector(".mini");const c=s.el.cloneNode(true);c.style.transform="scale("+(THUMB_W/1280)+")";c.style.transformOrigin="top left";m.appendChild(c);});}
$("prev").addEventListener("click",()=>showSlide((STATE.cur-1+STATE.slides.length)%STATE.slides.length));
$("next").addEventListener("click",()=>showSlide((STATE.cur+1)%STATE.slides.length));
window.addEventListener("resize",()=>{if(STATE.slides[STATE.cur])scaleSlide(STATE.slides[STATE.cur].el);});

/* ---------- inspector ---------- */
function renderInspector(){const box=$("insp");const s=STATE.slides[STATE.cur];if(!s){box.textContent="Select a slide.";return;}const sl=s.slide;
  if(sl.type==="chart"){const sheets=["sample"].concat((STATE.workbook?STATE.workbook.sheets.map(x=>x.name):[]));
    box.innerHTML=`<b>${esc(sl.title)}</b> · chart
      <div class="row"><label>Type</label><select id="iType">${["doughnut","bar","column","line","pie"].map(t=>`<option ${t===sl.chartType?"selected":""}>${t}</option>`).join("")}</select></div>
      <div class="row"><label>Source</label><select id="iSheet">${sheets.map(n=>`<option ${(sl.source.sheet===n||(sl.source.kind==="sample"&&n==="sample"))?"selected":""}>${esc(n)}</option>`).join("")}</select></div>
      <div class="row" id="colRow" style="${sl.source.kind==="sample"?"display:none":""}">
        <label>Cols</label><input id="iLbl" type="text" value="${sl.source.labelCol||"A"}" style="max-width:42px">
        <input id="iVal" type="text" value="${sl.source.valueCol||"C"}" style="max-width:42px">
        <input id="iR0" type="text" value="${sl.source.row0||4}" style="max-width:38px">
        <input id="iR1" type="text" value="${sl.source.row1||12}" style="max-width:38px"></div>`;
    $("iType").onchange=e=>aiApply([{op:"set_chart_type",target:STATE.cur,chartType:e.target.value}],"Inspector");
    $("iSheet").onchange=e=>{const v=e.target.value;aiApply([{op:"set_source",target:STATE.cur,source:v==="sample"?{kind:"sample",metric:sl.source.metric||"npat"}:{kind:"sheet",sheet:v,labelCol:"A",valueCol:"C",row0:4,row1:12}}],"Inspector");};
    if($("iLbl"))["iLbl","iVal","iR0","iR1"].forEach(id=>$(id).onchange=()=>aiApply([{op:"set_source",target:STATE.cur,source:{kind:"sheet",sheet:$("iSheet").value,labelCol:$("iLbl").value.toUpperCase(),valueCol:$("iVal").value.toUpperCase(),row0:+$("iR0").value,row1:+$("iR1").value}}],"Inspector"));
  }else if(sl.type==="table"){box.innerHTML=`<b>${esc(sl.title)}</b> · table<div class="row"><label>Hide 0</label><select id="iHz"><option value="1" ${sl.hideZero?"selected":""}>on</option><option value="0" ${!sl.hideZero?"selected":""}>off</option></select></div>`;
    $("iHz").onchange=e=>aiApply([{op:"toggle_hidezero",target:STATE.cur,value:e.target.value==="1"}],"Inspector");
  }else box.innerHTML=`<b>${esc(sl.title||sl.type)}</b> · ${sl.type} slide. Use the Copilot to edit.`;}
function colIdx(l){l=String(l).toUpperCase();let n=0;for(const ch of l)n=n*26+(ch.charCodeAt(0)-64);return n-1;}
function extractFromSheet(sheetName,labelCol,valueCol,r0,r1){const sh=STATE.workbook&&STATE.workbook.sheets.find(s=>s.name===sheetName);if(!sh)return{categories:[],values:[]};
  const li=colIdx(labelCol),vi=colIdx(valueCol),cats=[],vals=[];
  for(let r=r0-1;r<=r1-1&&r<sh.grid.length;r++){const row=sh.grid[r]||[];const lab=row[li],val=row[vi];if(lab==null&&val==null)continue;cats.push(lab==null?("Row"+(r+1)):String(lab));vals.push(num(val)?val:0);}
  return{categories:cats,values:vals};}

/* ---------- AI engine ---------- */
function specSummary(){const sp=STATE.spec;return{slides:sp.slides.map((s,i)=>({i,type:s.type,title:s.title})),theme:sp.theme,sheets:(STATE.workbook?STATE.workbook.sheets.map(s=>s.name):[])};}
async function aiInterpret(text){
  if(STATE.aiEndpoint){try{const res=await fetch(STATE.aiEndpoint,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({instruction:text,context:specSummary()})});const j=await res.json();if(j&&j.ops)return j;}catch(e){return{reply:"(Azure OpenAI endpoint unreachable — using local engine) "+localEngine(text).reply,ops:localEngine(text).ops};}}
  return localEngine(text);}
function findMetric(t){for(const k in METRIC_ALIASES)if(METRIC_ALIASES[k].some(a=>t.includes(a)))return k;return null;}
function findChartType(t){if(/\b(pie)\b/.test(t))return "pie";if(/doughnut|donut/.test(t))return "doughnut";if(/\bline\b|trend/.test(t))return "line";if(/\bbar\b|column\b/.test(t))return "bar";return null;}
function findTargetIdx(t){const m=t.match(/slide\s*#?(\d+)/);if(m){const i=+m[1]-1;if(i>=0&&i<STATE.spec.slides.length)return i;}
  for(let i=0;i<STATE.spec.slides.length;i++){const ti=(STATE.spec.slides[i].title||"").toLowerCase();if(ti&&t.includes(ti))return i;}return STATE.cur;}
function findColor(t){for(const n in COLORNAMES)if(new RegExp("\\b"+n+"\\b").test(t))return COLORNAMES[n];const hex=t.match(/#([0-9a-f]{6})/i);return hex?hex[1]:null;}
function localEngine(text){const t=" "+text.toLowerCase().trim()+" ";const ops=[];
  for(const k in TEMPLATES){if(new RegExp("\\b"+k+"\\b").test(t)||(k==="light"&&/light|minimal/.test(t))||(k==="exec"&&/executive|midnight/.test(t))||(k==="energy"&&/energy|green template/.test(t))){ops.push({op:"set_template",name:k});return{reply:`Applied the ${k} template.`,ops};}}
  let bm=text.match(/brand(?:\s*name)?\s*(?:to|=|:)?\s*["']?([a-z0-9 &]{2,30})["']?/i);
  if(/brand/.test(t)&&bm&&!/colou?r/.test(t)){ops.push({op:"set_brand",name:bm[1].trim().toUpperCase()});return{reply:`Brand set to “${bm[1].trim().toUpperCase()}”.`,ops};}
  if(/colou?r|accent|primary|theme/.test(t)){const c=findColor(t);if(c){const which=/accent/.test(t)?"accent":/dark|background/.test(t)?"dark":"primary";ops.push({op:"set_theme",key:which,color:c});return{reply:`Set ${which} colour to #${c}.`,ops};}}
  if(/\b(build|create|generate|make)\b/.test(t)&&/\b(deck|board pack|pack|slides|presentation|overview)\b/.test(t)){const goal=/profitab/.test(t)?"profitability":/growth|2031/.test(t)?"growth":/return|roe/.test(t)?"returns":/cash|balance/.test(t)?"cash":"board";ops.push({op:"build_deck",goal});return{reply:`Built a ${goal} deck from the workbook.`,ops};}
  if(/\b(remove|delete|drop)\b/.test(t)){const i=findTargetIdx(t);ops.push({op:"remove_slide",target:i});return{reply:`Removed slide ${i+1}.`,ops};}
  if(/\b(add|insert|new)\b/.test(t)){
    if(/kpi|key figure/.test(t)){ops.push({op:"add_slide",slideType:"kpi"});return{reply:"Added a KPI slide.",ops};}
    if(/table|kfi|indicator/.test(t)){ops.push({op:"add_slide",slideType:"table"});return{reply:"Added a Key Financial Indicators table.",ops};}
    if(/text|note|narrative|comment/.test(t)){ops.push({op:"add_slide",slideType:"text"});return{reply:"Added a notes slide.",ops};}
    const metric=findMetric(t)||"npat";const ct=findChartType(t)||(metric==="roe"?"line":metric==="npat"?"doughnut":"bar");ops.push({op:"add_slide",slideType:"chart",metric,chartType:ct});return{reply:`Added a ${ct} chart of ${metric}.`,ops};}
  const ct=findChartType(t);if(ct&&/\b(change|make|turn|convert|switch|set)\b/.test(t)){const i=findTargetIdx(t);ops.push({op:"set_chart_type",target:i,chartType:ct});return{reply:`Changed slide ${i+1} to a ${ct} chart.`,ops};}
  if(/\b(remap|source|from sheet|use sheet|point|pull from)\b/.test(t)){const i=findTargetIdx(t);const sm=t.match(/sheet\s+["']?([a-z0-9 _&]+?)["']?(?:[,.]| col| column| labels| rows|$)/);const sheet=sm?sm[1].trim():null;
    if(sheet){const real=STATE.workbook&&STATE.workbook.sheets.find(s=>s.name.toLowerCase()===sheet.toLowerCase());ops.push({op:"set_source",target:i,source:{kind:"sheet",sheet:real?real.name:sheet,labelCol:"A",valueCol:"C",row0:4,row1:12}});return{reply:`Pointed slide ${i+1} at sheet “${real?real.name:sheet}”.`,ops};}}
  if(/\b(rename|title|call it|relabel)\b/.test(t)){const m=text.match(/(?:to|as|:)\s+["']?(.+?)["']?\s*$/i);if(m){const i=findTargetIdx(t);ops.push({op:"rename_slide",target:i,title:m[1].trim()});return{reply:`Renamed slide ${i+1}.`,ops};}}
  if(/\b(move|reorder)\b/.test(t)){const i=findTargetIdx(t);const to=/\b(first|top|start)\b/.test(t)?0:/\b(last|end)\b/.test(t)?STATE.spec.slides.length-1:null;if(to!=null){ops.push({op:"move_slide",target:i,to});return{reply:`Moved slide ${i+1}.`,ops};}}
  if(/hide.*(zero|blank|empty)/.test(t)){ops.push({op:"toggle_hidezero",target:findTargetIdx(t),value:true});return{reply:"Hiding zero/blank rows.",ops};}
  if(/show all rows|show.*(zero|blank|empty)/.test(t)){ops.push({op:"toggle_hidezero",target:findTargetIdx(t),value:false});return{reply:"Showing all rows.",ops};}
  const rm=t.match(/roe.*?(\d{1,2}(?:\.\d)?)\s*%?/);if(rm&&/\b(set|change|use|lock)\b/.test(t)){ops.push({op:"set_roe",value:parseFloat(rm[1])});return{reply:`Locked headline ROE at ${rm[1]}% across the deck.`,ops};}
  return{reply:"I can add/remove/rename/reorder slides, change chart types, remap data sources, build a board pack, set ROE, and change template/brand colours. Try: “add a line chart of ROE” or “build a profitability board pack”.",ops:[]};}

function snapshot(){STATE.history.push(JSON.stringify({slides:STATE.spec.slides,theme:STATE.spec.theme,kf:STATE.spec.keyFigures}));if(STATE.history.length>40)STATE.history.shift();}
function GOALS(goal){const C=m=>chartSlide(m,m==="npat"?"doughnut":m==="roe"?"line":"bar");
  const sets={profitability:["title","kpi","npat","ebitda","kfi"],growth:["title","growth","revenue","kpi"],returns:["title","kpi","roe","kfi"],cash:["title","kpi","kfi","growth"],board:["title","kpi","npat","growth","kfi"]};
  return(sets[goal]||sets.board).map(k=>k==="title"?titleSlide():k==="kpi"?kpiSlide(STATE.spec):k==="kfi"?tableSlide():C(k));}
function applyOps(ops){const sp=STATE.spec;ops.forEach(o=>{try{
  if(o.op==="add_slide"){let s;if(o.slideType==="chart")s=chartSlide(o.metric||"npat",o.chartType);else if(o.slideType==="kpi")s=kpiSlide(sp);else if(o.slideType==="table")s=tableSlide();else if(o.slideType==="text")s=textSlide(o.title,o.bullets);else s=textSlide();const at=(o.target!=null)?o.target+1:sp.slides.length;sp.slides.splice(at,0,s);STATE.cur=at;}
  else if(o.op==="remove_slide")sp.slides.splice(o.target,1);
  else if(o.op==="rename_slide"){if(sp.slides[o.target])sp.slides[o.target].title=o.title;}
  else if(o.op==="move_slide"){const[s]=sp.slides.splice(o.target,1);sp.slides.splice(o.to,0,s);STATE.cur=o.to;}
  else if(o.op==="set_chart_type"){const s=sp.slides[o.target];if(s&&s.type==="chart")s.chartType=o.chartType;}
  else if(o.op==="set_source"){const s=sp.slides[o.target];if(s&&s.type==="chart"){s.source=o.source;if(o.source.kind==="sample"){const m=SAMPLE.series[o.source.metric]||SAMPLE.series.npat;s.data={categories:[...m.cats],values:[...m.vals]};}else s.data=extractFromSheet(o.source.sheet,o.source.labelCol,o.source.valueCol,o.source.row0,o.source.row1);}}
  else if(o.op==="toggle_hidezero"){const s=sp.slides[o.target];if(s&&s.type==="table")s.hideZero=o.value;else sp.hideZero=o.value;}
  else if(o.op==="set_roe"){sp.keyFigures.headlineRoe=o.value;sp.slides.forEach(s=>{if(s.type==="kpi")s.cards=kpiSlide(sp).cards;});}
  else if(o.op==="build_deck"){sp.slides=GOALS(o.goal);STATE.cur=0;}
  else if(o.op==="set_template"){const tpl=TEMPLATES[o.name];if(tpl){sp.theme.primary=tpl.primary;sp.theme.accent=tpl.accent;sp.theme.dark=tpl.dark;sp.theme.template=o.name;syncBrandBar();}}
  else if(o.op==="set_theme"){sp.theme[o.key]=o.color;syncBrandBar();}
  else if(o.op==="set_brand"){sp.theme.brandName=o.name;syncBrandBar();}
}catch(e){}});}
function aiApply(ops,who){if(!ops||!ops.length)return;snapshot();applyOps(ops);buildSlides();if(who==="Inspector")logMsg("a","Updated via inspector.","· "+ops[ops.length-1].op);}

function logMsg(role,text,ops){const l=$("log");const d=document.createElement("div");d.className="msg "+role;d.innerHTML=esc(text)+(ops?`<span class="ops">${esc(ops)}</span>`:"");l.appendChild(d);l.scrollTop=l.scrollHeight;}
async function sendAI(text){if(!text.trim())return;logMsg("u",text);$("ai").value="";const{reply,ops}=await aiInterpret(text);if(ops&&ops.length){snapshot();applyOps(ops);buildSlides();}logMsg("a",reply,ops&&ops.length?("✓ applied: "+ops.map(o=>o.op).join(", ")):null);}
$("send").addEventListener("click",()=>sendAI($("ai").value));
$("ai").addEventListener("keydown",e=>{if(e.key==="Enter")sendAI($("ai").value);});
$("undo").addEventListener("click",()=>{if(!STATE.history.length){logMsg("a","Nothing to undo.");return;}const p=JSON.parse(STATE.history.pop());STATE.spec.slides=p.slides;STATE.spec.theme=p.theme;STATE.spec.keyFigures=p.kf;syncBrandBar();buildSlides();logMsg("a","Reverted last change.");});

const CHIPS=["Build a profitability board pack","Add a line chart of ROE","Add an EBITDA by business unit chart","Change this slide to a pie chart","Use the Light Minimal template","Set accent colour to orange","Show all rows on the KFI table"];
function renderChips(){const c=$("chips");c.innerHTML="";CHIPS.forEach(t=>{const b=document.createElement("button");b.className="chip";b.textContent=t;b.onclick=()=>sendAI(t);c.appendChild(b);});}

/* ---------- theme + settings popovers ---------- */
function syncBrandBar(){const th=STATE.spec.theme;$("brandName").value=th.brandName;$("cPrimary").value="#"+th.primary;$("cAccent").value="#"+th.accent;$("cDark").value="#"+th.dark;$("tplSel").value=th.template||"sp";}
function wireBrandBar(){
  $("brandName").addEventListener("change",e=>aiApply([{op:"set_brand",name:e.target.value.toUpperCase()}]));
  $("cPrimary").addEventListener("input",e=>aiApply([{op:"set_theme",key:"primary",color:e.target.value.slice(1)}]));
  $("cAccent").addEventListener("input",e=>aiApply([{op:"set_theme",key:"accent",color:e.target.value.slice(1)}]));
  $("cDark").addEventListener("input",e=>aiApply([{op:"set_theme",key:"dark",color:e.target.value.slice(1)}]));
  $("tplSel").addEventListener("change",e=>aiApply([{op:"set_template",name:e.target.value}]));}
function initSettingsUI(cfg){
  $("cutoff").value=cfg.cutoff;const sel=$("roe");sel.innerHTML="";Object.entries(SAMPLE.roeByFy).forEach(([fy,v])=>{const o=document.createElement("option");o.value=v;o.textContent=`${v}%  ·  ${fy}`;sel.appendChild(o);});
  sel.value=cfg.roe;$("hideZero").checked=cfg.hideZero;$("lockKpi").checked=cfg.lockKpi;
}
function wireSettings(){
  $("cutoff").addEventListener("change",e=>{STATE.spec.meta.cutoff=e.target.value||"—";buildSlides();});
  $("roe").addEventListener("change",e=>aiApply([{op:"set_roe",value:parseFloat(e.target.value)}]));
  $("hideZero").addEventListener("change",e=>{STATE.spec.hideZero=e.target.checked;STATE.spec.slides.forEach(s=>{if(s.type==="table")s.hideZero=e.target.checked;});buildSlides();});
  $("lockKpi").addEventListener("change",e=>{STATE.spec.lockKpi=e.target.checked;});
}

/* ---------- popover plumbing ---------- */
const POPS={themeBtn:"themePop",setBtn:"setPop",aiBtn:"aiPop"};
function closePops(except){Object.values(POPS).forEach(p=>{if(p!==except)show($(p),false);});}
Object.entries(POPS).forEach(([btn,pop])=>{$(btn).addEventListener("click",e=>{e.stopPropagation();const open=!$(pop).classList.contains("hidden");closePops();show($(pop),!open);});});
document.addEventListener("click",e=>{if(!e.target.closest(".pop")&&!e.target.closest(".pill")&&!e.target.closest("#aiBtn"))closePops();});
$("aiPop").addEventListener("click",e=>e.stopPropagation());
$("themePop").addEventListener("click",e=>e.stopPropagation());
$("setPop").addEventListener("click",e=>e.stopPropagation());
$("aiEndpoint").addEventListener("change",e=>{STATE.aiEndpoint=e.target.value.trim();$("aiMode").textContent=STATE.aiEndpoint?"Azure OpenAI":"local engine";});

/* ---------- export ---------- */
function exportPptx(){
 try{
  if(typeof PptxGenJS==="undefined"){alert("PowerPoint engine is still loading. Please wait a moment and try again.");return;}
  const sp=STATE.spec,th=sp.theme;const p=new PptxGenJS();p.defineLayout({name:"W",width:13.333,height:7.5});p.layout="W";
  const foot=`Auto-generated from ${sp.meta.source} · cut-off ${sp.meta.cutoff} · figures in ${sp.meta.currency} · DRAFT`;
  const logoImg=s=>{try{s.addImage({data:LOGO,x:11.05,y:0.33,w:1.6,h:1.6*108/275});}catch(e){}};
  const hdr=(s,title,sub)=>{logoImg(s);
    s.addText(title||"",{x:0.7,y:0.5,w:9.8,h:0.7,fontSize:24,bold:true,color:th.dark,fontFace:"Arial",valign:"middle"});
    s.addShape("rect",{x:0.7,y:1.22,w:11.93,h:0.04,fill:{color:th.primary}});
    if(sub)s.addText(sub,{x:0.7,y:1.32,w:12,h:0.35,fontSize:12,color:"5C5C5C",fontFace:"Arial"});};
  const footer=s=>s.addText(foot,{x:0.7,y:7.05,w:12,h:0.3,fontSize:8,color:"8A939B"});
  sp.slides.forEach((sl,idx)=>{const s=p.addSlide();
    if(sl.type==="title"){logoImg(s);
      s.addText(sl.title,{x:0.9,y:2.9,w:11,h:1.1,fontSize:40,bold:true,color:th.dark,fontFace:"Arial"});
      s.addShape("rect",{x:0.92,y:4.05,w:4.6,h:0.06,fill:{color:th.primary}});
      s.addText(sl.subtitle||"",{x:0.9,y:4.2,w:11,h:0.5,fontSize:20,color:th.primary,fontFace:"Arial"});
      s.addText(`Cut-off ${sp.meta.cutoff} · generated ${sp.meta.generated}`,{x:0.9,y:6.6,w:11,h:0.4,fontSize:12,color:"8A939B",fontFace:"Arial"});return;}
    hdr(s,sl.title,sl.subtitle);
    if(sl.type==="kpi"){let x=0.7,w=2.95;(sl.cards||[]).slice(0,4).forEach(c=>{s.addShape("rect",{x,y:2.4,w,h:2.0,fill:{color:"E8F6F8"},line:{color:th.primary,width:1.5}});
      s.addText(c.label,{x:x+0.2,y:2.55,w:w-0.4,h:0.4,fontSize:13,bold:true,color:th.primary,fontFace:"Arial"});
      s.addText(String(c.value),{x:x+0.2,y:2.95,w:w-0.4,h:0.8,fontSize:34,bold:true,color:th.dark,fontFace:"Arial"});
      s.addText(c.note||"",{x:x+0.2,y:3.85,w:w-0.4,h:0.4,fontSize:10,color:"5C5C5C",fontFace:"Arial"});x+=w+0.18;});}
    else if(sl.type==="chart"){
      let img=null;try{const cv=STATE.slides[idx]&&STATE.slides[idx].el&&STATE.slides[idx].el.querySelector("canvas");if(cv)img=cv.toDataURL("image/png");}catch(e){}
      if(img){s.addImage({data:img,x:1.1,y:1.95,w:11.1,h:11.1*420/1100});}
      else{const d=sl.data;const isPie=sl.chartType==="doughnut"||sl.chartType==="pie";
        const ctype=isPie?(sl.chartType==="pie"?p.ChartType.pie:p.ChartType.doughnut):(sl.chartType==="line"?p.ChartType.line:p.ChartType.bar);
        const opt={x:isPie?2.6:0.8,y:1.9,w:isPie?7:11.7,h:4.5,showValue:true,chartColors:isPie?palette(th).map(c=>c.replace("#","")):[th.primary],showLegend:isPie,legendPos:"r"};
        if(isPie)opt.holeSize=sl.chartType==="doughnut"?55:0;if(!isPie)opt.barDir="col";
        s.addChart(ctype,[{name:sl.seriesName||"Series",labels:d.categories,values:d.values}],opt);}}
    else if(sl.type==="table"){const cm=sp._changeMap||{};let rows=sl.rows.slice();if(sl.hideZero)rows=rows.filter(r=>!r.values.every(v=>v==null||v===0));rows=rows.slice(0,11);
      const head=[{text:"Metric",options:{fill:th.primary,color:"FFFFFF",bold:true,align:"left"}}].concat(sl.fy.map(h=>({text:h||"",options:{fill:th.primary,color:"FFFFFF",bold:true,align:"center"}})));
      const body=rows.map((r,i)=>{const bg=i%2?"E1F4F6":"FFFFFF";return[{text:r.metric.slice(0,34),options:{fill:bg,color:th.dark,align:"left"}}].concat(r.values.map((v,ci)=>{const ch=cm["kfi:"+r.metric+"@"+ci];return{text:fmt(v,r.unit)+(ch?(ch.dir==="up"?" ▲":" ▼"):""),options:{fill:ch?(ch.dir==="up"?"E6F4EA":"FCE8E8"):bg,color:th.dark,bold:!!ch,align:"center"}};}));});
      s.addTable([head,...body],{x:0.7,y:1.95,w:11.95,fontSize:10,fontFace:"Arial",border:{type:"solid",color:"FFFFFF",pt:1},colW:[3.0,...sl.fy.map(()=>(11.95-3.0)/sl.fy.length)],rowH:0.34,valign:"middle"});}
    else if(sl.type==="changes"){const list=(sl.changes||[]).slice(0,11);
      if(!list.length){s.addText("No material changes versus the prior version.",{x:0.9,y:3,w:11,h:1,fontSize:18,color:"5C5C5C",fontFace:"Arial"});}
      else{const head=["Figure","Prior","New","Change"].map((t,i)=>({text:t,options:{fill:th.primary,color:"FFFFFF",bold:true,align:i?"center":"left"}}));
        const body=list.map((c,i)=>{const bg=i%2?"E1F4F6":"FFFFFF";const cc=c.dir==="up"?"1E8E3E":"C5221F";
          return[{text:c.label,options:{fill:bg,color:th.dark,align:"left"}},{text:c.oldText,options:{fill:bg,color:"5C5C5C",align:"center"}},{text:c.newText,options:{fill:bg,color:th.dark,bold:true,align:"center"}},{text:(c.dir==="up"?"▲ ":"▼ ")+c.pctText,options:{fill:bg,color:cc,bold:true,align:"center"}}];});
        s.addTable([head,...body],{x:0.7,y:1.95,w:11.95,fontSize:11,fontFace:"Arial",border:{type:"solid",color:"FFFFFF",pt:1},colW:[5.5,2.14,2.14,2.15],rowH:0.4,valign:"middle"});}}
    else if(sl.type==="text"){s.addText((sl.bullets||[]).map(b=>({text:b,options:{bullet:{code:"2022"},color:th.dark,fontSize:18,paraSpaceAfter:10}})),{x:0.9,y:2.3,w:11.5,h:4,fontFace:"Arial"});}
    footer(s);});
  overlay(true,"Building PowerPoint…");
  p.writeFile({fileName:`SP_Group_Deck_${sp.meta.cutoff}.pptx`}).then(()=>overlay(false)).catch(e=>{overlay(false);alert("Could not export: "+(e&&e.message||e));});
 }catch(e){overlay(false);alert("Could not export: "+(e&&e.message||e));}
}
$("dlBtn").addEventListener("click",exportPptx);

/* ---------- password gate ---------- */
const ACCESS_PW="CHANGE_ME";
function openLanding(){show($("gate"),false);show($("studio"),false);show($("intake"),false);show($("workspace"),true);renderWorkspace();}
function tryUnlock(){if($("pw").value===ACCESS_PW){try{sessionStorage.setItem("x2s_ok","1");}catch(e){}openLanding();$("pw").value="";}else{show($("pwErr"),true);$("pw").value="";$("pw").focus();}}
$("pwBtn").addEventListener("click",tryUnlock);
$("pw").addEventListener("keydown",e=>{if(e.key==="Enter")tryUnlock();});
try{if(sessionStorage.getItem("x2s_ok")==="1")openLanding();else $("pw").focus();}catch(e){$("pw").focus();}

/* ===========================================================================
   LINEAGE & VERSIONING — Source Series ▸ Versions ▸ Decks
   A persistent ledger (localStorage) tracks which deck came from which
   workbook version. When a new version is ingested, decks diff against the
   version they reflect; updates are surfaced (notify), applied (auto) or
   held (pinned). Edits live in the deck-spec, so a refresh re-renders the
   numbers everywhere and highlights what changed.
   =========================================================================== */
let _memLedger=null,_ingestTarget=null;
function loadLedger(){try{const s=localStorage.getItem("x2s_ledger");if(s)return JSON.parse(s);}catch(e){}return _memLedger||{series:{},versions:{},decks:{},seq:0};}
function saveLedger(L){_memLedger=L;try{localStorage.setItem("x2s_ledger",JSON.stringify(L));}catch(e){try{trimLedger(L);localStorage.setItem("x2s_ledger",JSON.stringify(L));}catch(e2){}}}
function trimLedger(L){const vs=Object.values(L.versions).sort((a,b)=>b.capturedAt-a.capturedAt);vs.slice(2).forEach(v=>{if(v.catalog)v.catalog=null;});}
function hashStr(s){let h=0x811c9dc5;for(let i=0;i<s.length;i++){h^=s.charCodeAt(i);h=(h*0x01000193)>>>0;}return h.toString(16);}

function snapFromCatalog(cat){
  if(!cat)return{fy:[],kfi:[],roeByFy:{},headlineRoe:null};
  const fy=(cat.fy||SAMPLE.fy).slice();
  const kfi=(cat.kfi||[]).map(r=>({metric:String(r.metric),values:(r.values||[]).slice(),unit:r.unit}));
  const roeRow=kfi.find(r=>/^roe$/i.test(String(r.metric).trim()));
  const roeByFy={};let headlineRoe=null;
  if(roeRow)roeRow.values.forEach((v,i)=>{if(num(v)){roeByFy[fy[i]||("c"+i)]=v;if(headlineRoe==null)headlineRoe=v;}});
  return{fy,kfi,roeByFy,headlineRoe};
}
function fingerprint(cat){const snap=snapFromCatalog(cat);return snap.kfi.map(r=>r.metric).sort().join("|").toLowerCase();}
function deriveSeriesName(name,isSample){if(isSample)return "Group P&L (sample)";
  let b=String(name||"Workbook").replace(/\.(xlsx|xlsm|xls)$/i,"");b=b.replace(/[ _-]*v?\d+$/i,"").trim();return b||"Workbook";}
function seriesVersions(L,sid){return Object.values(L.versions).filter(v=>v.seriesId===sid).sort((a,b)=>a.n-b.n);}
function latestVersion(L,sid){const vs=seriesVersions(L,sid);return vs[vs.length-1];}

function ingestWorkbook(name,catalog,isSample,forceSeriesId){
  const L=loadLedger();const snap=snapFromCatalog(catalog);const fp=fingerprint(catalog);
  const h=hashStr(JSON.stringify(snap.kfi)+"|"+snap.fy.join(","));
  let sid=forceSeriesId||null;
  if(!sid){const m=Object.values(L.series).find(s=>s.schema===fp&&!!s.sample===!!isSample);sid=m?m.id:null;}
  if(!sid){sid="ser"+(++L.seq);L.series[sid]={id:sid,name:deriveSeriesName(name,isSample),schema:fp,sample:!!isSample,createdAt:Date.now()};}
  let ver=Object.values(L.versions).find(v=>v.seriesId===sid&&v.hash===h);
  if(!ver){const n=seriesVersions(L,sid).length+1;const vid="ver"+(++L.seq);
    ver={id:vid,seriesId:sid,n,label:"v"+n,fileName:name,hash:h,capturedAt:Date.now(),snapshot:snap,catalog:catalog};
    L.versions[vid]=ver;saveLedger(L);
    const savedWb=STATE.workbook,savedName=STATE.wbName;
    Object.values(L.decks).filter(d=>d.seriesId===sid&&d.policy==="auto"&&d.boundVersionId!==vid).forEach(d=>applyVersionToDeck(d.id,true));
    STATE.workbook=savedWb;STATE.wbName=savedName;
  } else saveLedger(L);
  return{seriesId:sid,versionId:ver.id};
}

/* ---------- diff ---------- */
function diffSnap(oldS,newS){const out=[],map={};const oldBy={};(oldS.kfi||[]).forEach(r=>oldBy[r.metric]=r);
  (newS.kfi||[]).forEach(r=>{const o=oldBy[r.metric];
    r.values.forEach((nv,ci)=>{const ov=o?o.values[ci]:undefined;const fy=(newS.fy[ci])||("Col"+(ci+1));
      if(num(nv)&&num(ov)){if(Math.abs(nv-ov)>Math.max(1e-9,Math.abs(ov)*0.0005)){const dir=nv>ov?"up":"down";const pct=ov!==0?((nv-ov)/Math.abs(ov)*100):null;
        const rec={role:"kfi:"+r.metric+"@"+ci,label:r.metric+" · "+fy,old:ov,new:nv,dir,pct,unit:r.unit,oldText:fmt(ov,r.unit),newText:fmt(nv,r.unit),pctText:(pct==null?"n/a":Math.abs(pct).toFixed(1)+"%")};
        out.push(rec);map[rec.role]={dir,old:ov,new:nv};}}
      else if(num(nv)&&!num(ov)){const rec={role:"kfi:"+r.metric+"@"+ci,label:r.metric+" · "+fy,old:ov,new:nv,dir:"up",pct:null,unit:r.unit,oldText:"–",newText:fmt(nv,r.unit),pctText:"new"};
        out.push(rec);map[rec.role]={dir:"up",old:null,new:nv};}});});
  out.sort((a,b)=>Math.abs(b.pct||0)-Math.abs(a.pct||0));return{list:out,map};}

function changesSlide(list,oldLabel,newLabel){return{id:uid(),type:"changes",title:"What changed",note:`Material movements: ${oldLabel} → ${newLabel}`,changes:list.slice(0,11)};}

/* ---------- propagation (refresh a deck to the latest version) ---------- */
function recomputeDeckSpec(spec,catalog,snap){
  if(Object.keys(snap.roeByFy).length)spec.keyFigures.roeByFy=snap.roeByFy;
  if(num(snap.headlineRoe))spec.keyFigures.headlineRoe=snap.headlineRoe;
  spec.slides.forEach(sl=>{
    if(sl.type==="table"){sl.fy=snap.fy.slice();sl.rows=JSON.parse(JSON.stringify(snap.kfi));}
    else if(sl.type==="kpi"){sl.cards=kpiSlide(spec).cards;}
    else if(sl.type==="chart"&&sl.source&&sl.source.kind==="sheet"){sl.data=extractFromSheet(sl.source.sheet,sl.source.labelCol,sl.source.valueCol,sl.source.row0,sl.source.row1);}
  });return spec;}
function applyVersionToDeck(deckId,addSummary){
  const L=loadLedger();const d=L.decks[deckId];if(!d)return null;const latest=latestVersion(L,d.seriesId);if(!latest)return null;
  const oldV=L.versions[d.boundVersionId];const {list,map}=diffSnap(oldV?oldV.snapshot:{fy:[],kfi:[]},latest.snapshot);
  const spec=JSON.parse(JSON.stringify(d.spec));
  STATE.workbook=latest.catalog||STATE.workbook;
  recomputeDeckSpec(spec,latest.catalog,latest.snapshot);spec._changeMap=map;
  if(addSummary!==false&&list.length){spec.slides=spec.slides.filter(s=>s.type!=="changes");
    const cs=changesSlide(list,oldV?oldV.label:"prior",latest.label);const at=spec.slides.findIndex(s=>s.type==="title");spec.slides.splice(at>=0?at+1:0,0,cs);}
  d.spec=JSON.parse(JSON.stringify(spec));d.boundVersionId=latest.id;d.updatedAt=Date.now();saveLedger(L);refreshLibBadge();
  return{spec,changes:list};}

/* ---------- deck save / open ---------- */
function saveCurrentDeck(silent){if(!STATE.spec)return;const L=loadLedger();let id=STATE.deckId;
  if(!id||!L.decks[id]){id="deck"+(++L.seq);STATE.deckId=id;}
  const sName=L.series[STATE.seriesId]?L.series[STATE.seriesId].name:"Workbook";const existing=L.decks[id];
  const count=Object.values(L.decks).filter(d=>d.seriesId===STATE.seriesId).length;
  const name=existing?existing.name:(sName+" deck "+(existing?count:count+1));
  const clean=JSON.parse(JSON.stringify(STATE.spec));delete clean._changeMap;
  L.decks[id]={id,name,seriesId:STATE.seriesId,boundVersionId:STATE.versionId,policy:existing?existing.policy:"notify",spec:clean,createdAt:existing?existing.createdAt:Date.now(),updatedAt:Date.now()};
  saveLedger(L);const dn=$("deckName");if(dn)dn.textContent=name;
  if(!silent){const v=L.versions[STATE.versionId];logMsg("a",`Saved “${name}” to the workspace, linked to ${sName} ${v?v.label:""}. It will track future versions of this source.`);}}
function openDeckIntoStudio(id){const L=loadLedger();const d=L.decks[id];if(!d)return;const v=L.versions[d.boundVersionId];const s=L.series[d.seriesId];
  STATE.workbook=v&&v.catalog?v.catalog:STATE.workbook;STATE.isSample=!!(s&&s.sample);STATE.wbName=s?s.name:d.name;
  STATE.seriesId=d.seriesId;STATE.versionId=d.boundVersionId;STATE.deckId=id;
  STATE.spec=JSON.parse(JSON.stringify(d.spec));STATE.cur=0;STATE.history=[];
  show($("review"),false);show($("gate"),false);show($("workspace"),false);show($("intake"),false);show($("studio"),true);
  initSettingsUI({cutoff:STATE.spec.meta.cutoff,roe:STATE.spec.keyFigures.headlineRoe,hideZero:STATE.spec.hideZero,lockKpi:STATE.spec.lockKpi});
  syncBrandBar();renderChips();updateSrcLabel();
  requestAnimationFrame(()=>{buildSlides();$("log").innerHTML="";logMsg("a",`Opened “${d.name}” (${v?v.label:""}). Edit it, or go back to the Workspace to manage versions.`);});}

/* ---------- status + labels ---------- */
function deckStatus(L,d){const latest=latestVersion(L,d.seriesId);const bv=L.versions[d.boundVersionId];
  if(d.policy==="pinned")return{k:"pin",t:"Pinned · "+(bv?bv.label:"")};
  if(latest&&d.boundVersionId!==latest.id)return{k:"upd",t:"Update available → "+latest.label};
  return{k:"cur",t:"Current · "+(latest?latest.label:(bv?bv.label:""))};}
function updateSrcLabel(){try{const L=loadLedger();const s=L.series[STATE.seriesId];const v=L.versions[STATE.versionId];const d=L.decks[STATE.deckId];
  $("srcName").textContent=(s?s.name:STATE.wbName)+(v?" · "+v.label:"");const dn=$("deckName");if(dn)dn.textContent=d?d.name:(STATE.wbName||"Deck");}catch(e){$("srcName").textContent=STATE.wbName;}}
function refreshLibBadge(){try{const L=loadLedger();const any=Object.values(L.decks).some(d=>{const lv=latestVersion(L,d.seriesId);return d.policy!=="pinned"&&lv&&d.boundVersionId!==lv.id;});const dot=$("libDot");if(dot)show(dot,any);}catch(e){}}
function notifyBehindDecks(){try{const L=loadLedger();const lv=latestVersion(L,STATE.seriesId);if(!lv)return;
  const behind=Object.values(L.decks).filter(d=>d.seriesId===STATE.seriesId&&d.id!==STATE.deckId&&d.policy!=="pinned"&&d.boundVersionId!==lv.id);
  if(behind.length)logMsg("a",`Note: ${behind.length} saved deck(s) from an earlier version of ${L.series[STATE.seriesId].name} now have updates available (${lv.label}). Open the Library to review or apply them.`);}catch(e){}}

/* ---------- demo: a modified sample = a new version ---------- */
function makeModifiedSample(step){step=step||1;const cat=sampleCatalog();cat.notes=["Modified illustrative figures (demo new version)."];cat.kfi=JSON.parse(JSON.stringify(SAMPLE.kfi));
  const bump=(name,fn)=>{const r=cat.kfi.find(x=>x.metric===name);if(r)r.values=r.values.map(v=>num(v)?fn(v):v);};
  bump("ROE",v=>Math.round((v+0.6*step)*10)/10);bump("ROIC",v=>Math.round((v+0.3*step)*10)/10);
  bump("Gearing",v=>Math.round((v-1.2*step)*10)/10);bump("FFO Interest Cover",v=>Math.round((v+0.2*step)*10)/10);return cat;}

/* ---------- Workspace (home) ---------- */
function showWorkspace(){show($("gate"),false);show($("studio"),false);show($("intake"),false);show($("review"),false);show($("workspace"),true);renderWorkspace();}
function openIntake(){show($("intake"),true);}
function closeIntake(){show($("intake"),false);}
function closeReview(){show($("review"),false);}

function lineageDiagramSVG(){return `<svg viewBox="0 0 720 175" width="100%" style="max-width:660px">
  <defs><marker id="ar" markerWidth="9" markerHeight="9" refX="6" refY="3" orient="auto"><path d="M0 0l6 3-6 3z" fill="#B7791F"/></marker></defs>
  <text x="90" y="22" text-anchor="middle" fill="#64748B" font-size="12" font-family="Arial">Excel</text>
  <text x="245" y="22" text-anchor="middle" fill="#64748B" font-size="12" font-family="Arial">new version</text>
  <circle cx="90" cy="56" r="14" fill="#00838A"/><text x="90" y="61" text-anchor="middle" fill="#fff" font-size="12" font-weight="700" font-family="Arial">v1</text>
  <line x1="104" y1="56" x2="231" y2="56" stroke="#D4DBE5" stroke-width="3"/>
  <circle cx="245" cy="56" r="14" fill="#00838A"/><text x="245" y="61" text-anchor="middle" fill="#fff" font-size="12" font-weight="700" font-family="Arial">v2</text>
  <line x1="90" y1="70" x2="90" y2="108" stroke="#D4DBE5" stroke-width="2"/>
  <rect x="36" y="108" width="108" height="36" rx="9" fill="#FDF4E3" stroke="#F0D9A8"/><text x="90" y="131" text-anchor="middle" fill="#B7791F" font-size="11" font-weight="600" font-family="Arial">Deck A</text>
  <line x1="146" y1="126" x2="232" y2="80" stroke="#B7791F" stroke-width="2" stroke-dasharray="4 3" marker-end="url(#ar)"/>
  <text x="205" y="120" fill="#B7791F" font-size="10" font-family="Arial">update available</text>
  <line x1="245" y1="70" x2="245" y2="108" stroke="#D4DBE5" stroke-width="2"/>
  <rect x="300" y="108" width="150" height="36" rx="9" fill="#E6F4F5" stroke="#D1ECEE"/><text x="375" y="131" text-anchor="middle" fill="#006B71" font-size="11" font-weight="600" font-family="Arial">Deck B · current v2</text>
  <rect x="478" y="108" width="150" height="36" rx="9" fill="#E7EEFD" stroke="#cfe0fb"/><text x="553" y="131" text-anchor="middle" fill="#2563EB" font-size="11" font-weight="600" font-family="Arial">Deck C · pinned v1</text>
</svg>`;}

function emptyStateHTML(){return `<div class="empty">
  <h2>Turn your Excel into a board deck — and keep it in sync</h2>
  <p>Generate a presentation from your budget workbook. Every deck remembers which version of the Excel it came from. Load a newer version and the decks that should change light up — review the differences and update in one click.</p>
  <div class="diagram">${lineageDiagramSVG()}</div>
  <div class="cta"><button class="btn primary" data-act="new"><svg viewBox="0 0 24 24"><path d="M12 5v14M5 12h14"/></svg>New deck from Excel</button><button class="btn ghost" data-act="sample">Load sample data</button></div>
  <div class="steps">
    <div class="step"><div class="n">1</div><b>Generate</b><span>Upload the consolidation workbook. A branded, editable deck is built automatically.</span></div>
    <div class="step"><div class="n">2</div><b>A new version arrives</b><span>Load next period's Excel. It is tracked as v2, v3 … of the same source.</span></div>
    <div class="step"><div class="n">3</div><b>Review &amp; update</b><span>Decks on an older version flag “update available”. See what changed and refresh the numbers in one click.</span></div>
  </div></div>`;}

function renderWorkspace(){
  const L=loadLedger();const series=Object.values(L.series).sort((a,b)=>a.createdAt-b.createdAt);
  const side=$("srcList");
  if(!series.length){side.innerHTML=`<div class="muted small" style="padding:8px 6px">No sources yet</div>`;$("wsMain").innerHTML=emptyStateHTML();return;}
  if(!STATE.selSeries||!L.series[STATE.selSeries])STATE.selSeries=series[0].id;
  side.innerHTML=series.map(s=>{const vs=seriesVersions(L,s.id);const lv=latestVersion(L,s.id);
    const upd=Object.values(L.decks).some(d=>d.seriesId===s.id&&d.policy!=="pinned"&&lv&&d.boundVersionId!==lv.id);
    return `<button class="srcitem ${s.id===STATE.selSeries?"sel":""}" data-act="selsrc" data-id="${s.id}"><span class="srcname">${esc(s.name)}</span><span class="srcmeta">${vs.length}v</span>${upd?'<span class="dot-upd" title="updates available"></span>':''}</button>`;}).join("");
  const s=L.series[STATE.selSeries];const vs=seriesVersions(L,s.id);const latest=vs[vs.length-1];
  const decks=Object.values(L.decks).filter(d=>d.seriesId===s.id);
  const behind=decks.filter(d=>d.policy!=="pinned"&&latest&&d.boundVersionId!==latest.id);
  let h=`<div class="ws-head"><div><h2>${esc(s.name)}</h2><div class="muted">Latest ${esc(latest?latest.label:"")} · ${vs.length} version(s) · ${decks.length} deck(s)</div></div>
    <div class="ws-actions"><button class="btn ghost" data-act="ingest" data-id="${s.id}"><svg viewBox="0 0 24 24"><path d="M12 16V4M8 8l4-4 4 4M4 20h16"/></svg>Ingest new version</button>${s.sample?`<button class="btn ghost" data-act="demo" data-id="${s.id}">⚡ Demo: modified version</button>`:""}<button class="btn primary" data-act="newfrom" data-id="${s.id}"><svg viewBox="0 0 24 24"><path d="M12 5v14M5 12h14"/></svg>New deck</button></div></div>`;
  h+=`<div class="timeline">`+vs.map((v,i)=>`<div class="vnode ${latest&&v.id===latest.id?"latest":""}"><div class="vdot"></div><div class="vlabel">${esc(v.label)}</div><div class="vdate">${new Date(v.capturedAt).toLocaleDateString()}</div><div class="vcount">${decks.filter(d=>d.boundVersionId===v.id).length} deck(s)</div></div>${i<vs.length-1?'<div class="vconn"></div>':''}`).join("")+`</div>`;
  if(behind.length)h+=`<div class="banner"><span><b>${esc(latest.label)} is available.</b> ${behind.length} deck(s) can be updated to the latest figures.</span><button class="btn primary sm" data-act="updateall" data-id="${s.id}">Review &amp; update all</button></div>`;
  h+=`<div class="deckgrid">`;
  if(!decks.length)h+=`<div class="muted">No decks yet for this source. Click “New deck”.</div>`;
  decks.forEach(d=>{const st=deckStatus(L,d);const bv=L.versions[d.boundVersionId];
    const track=vs.map(v=>`<span class="tdot ${v.id===d.boundVersionId?"on":""} ${latest&&v.id===latest.id?"lt":""}" title="${esc(v.label)}"></span>`).join('<span class="tbar"></span>');
    h+=`<div class="deckcard"><div class="dc-top"><div class="dc-name">${esc(d.name)}</div><span class="stp ${st.k}">${esc(st.t)}</span></div>
      <div class="dc-track">${track}<span class="dc-on">on ${esc(bv?bv.label:"?")}</span></div>
      <div class="dc-foot"><select class="sel" data-act="policy" data-id="${d.id}"><option value="notify" ${d.policy==="notify"?"selected":""}>Notify</option><option value="auto" ${d.policy==="auto"?"selected":""}>Auto</option><option value="pinned" ${d.policy==="pinned"?"selected":""}>Pinned</option></select>
      <div class="dc-btns">${st.k==="upd"?`<button class="btn primary sm" data-act="review" data-id="${d.id}">Review &amp; update</button>`:""}<button class="btn sm" data-act="open" data-id="${d.id}">Open</button><button class="linkbtn rn" data-act="rename" data-id="${d.id}">Rename</button><button class="linkbtn" data-act="del" data-id="${d.id}">Delete</button></div></div></div>`;});
  h+=`</div>`;
  $("wsMain").innerHTML=h;}

/* ---------- review modal ---------- */
function reviewDeck(id){const L=loadLedger();const d=L.decks[id];const latest=latestVersion(L,d.seriesId);const oldV=L.versions[d.boundVersionId];
  const {list}=diffSnap(oldV?oldV.snapshot:{fy:[],kfi:[]},latest.snapshot);
  $("reviewTitle").textContent=`${d.name} — ${oldV?oldV.label:"prior"} → ${latest.label}`;
  let h=`<div class="diffsum">${list.length} material change(s) on the figures this deck uses. Review, then update or keep it pinned.</div>`;
  if(!list.length)h+=`<div class="muted">No material changes versus the prior version.</div>`;
  else list.forEach(c=>{h+=`<div class="diffrow"><span class="lab">${esc(c.label)}</span><span class="ov">${esc(c.oldText)}</span><span style="color:var(--dim)">→</span><span class="nv">${esc(c.newText)}</span><span class="pc ${c.dir}">${c.dir==="up"?"▲":"▼"} ${esc(c.pctText)}</span></div>`;});
  h+=`<div class="modal-foot"><button class="btn primary" data-act="apply" data-id="${id}">Apply update &amp; open</button><button class="btn ghost" data-act="pin" data-id="${id}">Keep pinned on ${esc(oldV?oldV.label:"current")}</button></div>`;
  $("reviewBody").innerHTML=h;show($("review"),true);
  const ap=$("reviewBody").querySelector("[data-act='apply']");if(ap)ap.onclick=()=>{applyVersionToDeck(id,true);openDeckIntoStudio(id);};
  const pn=$("reviewBody").querySelector("[data-act='pin']");if(pn)pn.onclick=()=>{const L2=loadLedger();L2.decks[id].policy="pinned";saveLedger(L2);closeReview();renderWorkspace();};}
function reviewAll(sid){const L=loadLedger();const latest=latestVersion(L,sid);
  const behind=Object.values(L.decks).filter(d=>d.seriesId===sid&&d.policy!=="pinned"&&latest&&d.boundVersionId!==latest.id);
  const vs=seriesVersions(L,sid);const prev=vs[vs.length-2]||{snapshot:{fy:[],kfi:[]}};
  const {list}=diffSnap(prev.snapshot,latest.snapshot);
  $("reviewTitle").textContent=`Update ${behind.length} deck(s) to ${latest.label}`;
  let h=`<div class="diffsum">${behind.length} deck(s) will be refreshed to ${esc(latest.label)} with these figures. Pinned decks are left untouched.</div>`;
  list.forEach(c=>{h+=`<div class="diffrow"><span class="lab">${esc(c.label)}</span><span class="ov">${esc(c.oldText)}</span><span style="color:var(--dim)">→</span><span class="nv">${esc(c.newText)}</span><span class="pc ${c.dir}">${c.dir==="up"?"▲":"▼"} ${esc(c.pctText)}</span></div>`;});
  h+=`<div class="modal-foot"><button class="btn primary" data-act="applyall">Apply to all ${behind.length} deck(s)</button><button class="btn ghost" data-act="cancel">Cancel</button></div>`;
  $("reviewBody").innerHTML=h;show($("review"),true);
  $("reviewBody").querySelector("[data-act='applyall']").onclick=()=>{behind.forEach(d=>applyVersionToDeck(d.id,true));closeReview();renderWorkspace();};
  $("reviewBody").querySelector("[data-act='cancel']").onclick=closeReview;}

/* ---------- workspace wiring ---------- */
$("newDeckBtn").addEventListener("click",openIntake);
$("howBtn").addEventListener("click",()=>{$("wsMain").innerHTML=emptyStateHTML();});
$("backWs").addEventListener("click",showWorkspace);
$("saveBtn").addEventListener("click",()=>saveCurrentDeck(false));
$("intakeClose").addEventListener("click",closeIntake);
$("intake").addEventListener("click",e=>{if(e.target===$("intake"))closeIntake();});
$("reviewClose").addEventListener("click",closeReview);
$("review").addEventListener("click",e=>{if(e.target===$("review"))closeReview();});
$("workspace").addEventListener("click",e=>{const t=e.target.closest("[data-act]");if(!t)return;const act=t.getAttribute("data-act"),id=t.getAttribute("data-id");
  if(act==="selsrc"){STATE.selSeries=id;renderWorkspace();}
  else if(act==="open")openDeckIntoStudio(id);
  else if(act==="review")reviewDeck(id);
  else if(act==="updateall")reviewAll(id);
  else if(act==="newfrom"||act==="new")openIntake();
  else if(act==="ingest"){_ingestTarget=id;$("wsFile").click();}
  else if(act==="demo"){const L0=loadLedger();const stp=seriesVersions(L0,id).length;ingestWorkbook("Sample data (rev "+stp+")",makeModifiedSample(stp),true,id);renderWorkspace();}
  else if(act==="sample"){closeIntake();STATE.isSample=true;STATE.wbName="Sample data";STATE.workbook=sampleCatalog();launch();}
  else if(act==="rename"){const L=loadLedger();const d=L.decks[id];const nn=prompt("Rename deck",d.name);if(nn&&nn.trim()){d.name=nn.trim();saveLedger(L);renderWorkspace();}}
  else if(act==="del"){if(confirm("Delete this deck from the workspace?")){const L=loadLedger();delete L.decks[id];saveLedger(L);renderWorkspace();}}});
$("workspace").addEventListener("change",e=>{const t=e.target.closest("[data-act='policy']");if(!t)return;const id=t.getAttribute("data-id");const L=loadLedger();const d=L.decks[id];if(!d)return;
  d.policy=t.value;saveLedger(L);if(d.policy==="auto"){const lv=latestVersion(L,d.seriesId);if(lv&&d.boundVersionId!==lv.id)applyVersionToDeck(id,true);}renderWorkspace();});
$("wsFile").addEventListener("change",()=>{const f=$("wsFile").files[0];if(!f)return;overlay(true,"Reading workbook…");const r=new FileReader();
  r.onload=e=>{try{const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});const cat=parseWorkbook(wb);ingestWorkbook(f.name,cat,false,_ingestTarget);overlay(false);renderWorkspace();}catch(err){overlay(false);alert("Could not read file: "+err.message);}};
  r.readAsArrayBuffer(f);$("wsFile").value="";});

wireBrandBar(); wireSettings();
