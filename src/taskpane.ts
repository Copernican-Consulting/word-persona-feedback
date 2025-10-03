/* eslint-disable @typescript-eslint/no-explicit-any */

/**
 * Persona Feedback – Word Add-in task pane
 * - Color badge = derived from the exact persona HEX (pane + comments match)
 * - Summaries only in task pane (no body insertion)
 * - Single “Clear comments” that wipes all via document.comments
 * - Thick traffic-light bars (green ≥80, yellow 50–79, red <50)
 * - Export opens print-ready report so you can Save as PDF
 * - Run lock to prevent duplicate runs
 */

type Provider = "openrouter" | "ollama";
type Persona = { id:string; enabled:boolean; name:string; system:string; instruction:string; color:string; };
type PersonaSet = { id:string; name:string; personas:Persona[] };
type ProviderSettings = { provider:Provider; openrouterKey?:string; model:string };
type AppSettings = { provider:ProviderSettings; personaSetId:string; personaSets:PersonaSet[] };

type PersonaRunStatus = "idle" | "running" | "done" | "error";
type PersonaRunResult = {
  personaId:string; personaName:string; status:PersonaRunStatus;
  scores?:{ clarity:number; tone:number; alignment:number };
  global_feedback?:string;
  comments?:{ quote:string; spanStart:number; spanEnd:number; comment:string }[];
  unmatched?:{ quote:string; comment:string }[];
  raw?:any; error?:string;
};

// ---------- Globals ----------
const LS_KEY = "pf.settings.v1";
let SETTINGS: AppSettings;
let LAST_RESULTS: PersonaRunResult[] = [];
let RUN_LOCK = false;

// ---------- DOM helpers ----------
function byId<T extends HTMLElement>(id:string){ return document.getElementById(id) as T | null; }
function req<T extends HTMLElement>(id:string){ const el=byId<T>(id); if(!el) throw new Error(`#${id} missing`); return el; }
function toast(t:string){ const b=byId<HTMLDivElement>("toast"); if(!b) return; const s=byId<HTMLSpanElement>("toastMsg"); if(s) s.textContent=t; b.style.display="block"; setTimeout(()=>b.style.display="none",2000); }
function log(msg:string, data?:any){ const p=byId<HTMLDivElement>("debugLog"); if(data!==undefined) console.log(msg,data); else console.log(msg); if(!p) return; const d=document.createElement("div"); d.style.whiteSpace="pre-wrap"; d.textContent=data?`${msg} ${safeJson(data)}`:msg; p.appendChild(d); p.scrollTop=p.scrollHeight; }
function safeJson(x:any){ try{return JSON.stringify(x,null,2);}catch{return String(x);} }
function showView(id:"view-review"|"view-settings"){ const r=byId<HTMLDivElement>("view-review"); const s=byId<HTMLDivElement>("view-settings"); const b=byId<HTMLButtonElement>("btnBack"); if(id==="view-review"){ r?.classList.remove("hidden"); s?.classList.add("hidden"); b?.classList.add("hidden"); } else { r?.classList.add("hidden"); s?.classList.remove("hidden"); b?.classList.remove("hidden"); } }
function confirmAsync(title:string,message:string){ return new Promise<boolean>(res=>{ const o=req<HTMLDivElement>("confirmOverlay"); req<HTMLHeadingElement>("confirmTitle").textContent=title; req<HTMLDivElement>("confirmMessage").textContent=message; o.style.display="flex"; const ok=req<HTMLButtonElement>("confirmOk"); const no=req<HTMLButtonElement>("confirmCancel"); const done=(v:boolean)=>{o.style.display="none"; ok.onclick=null; no.onclick=null; res(v);}; ok.onclick=()=>done(true); no.onclick=()=>done(false);}); }
function escapeHtml(s:string){ return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }

// ---------- Color: derive emoji from HEX hue so pane + comments match ----------
function hexToRgb(hex:string){ const m = hex.trim().replace("#",""); if(m.length!==6) return {r:200,g:200,b:200}; return { r:parseInt(m.slice(0,2),16), g:parseInt(m.slice(2,4),16), b:parseInt(m.slice(4,6),16) }; }
function rgbToHue({r,g,b}:{r:number;g:number;b:number}){ r/=255; g/=255; b/=255; const max=Math.max(r,g,b), min=Math.min(r,g,b); const d=max-min; let h=0; if(d===0) h=0; else if(max===r) h=((g-b)/d)%6; else if(max===g) h=(b-r)/d+2; else h=(r-g)/d+4; h=Math.round(h*60); if(h<0) h+=360; return h; }
function hueToEmoji(h:number){
  const palette = [{h:0,e:"🟥"},{h:30,e:"🟧"},{h:60,e:"🟨"},{h:120,e:"🟩"},{h:210,e:"🟦"},{h:280,e:"🟪"}];
  let best=palette[0], diff=999;
  for(const p of palette){ const d=Math.min(Math.abs(h-p.h),360-Math.abs(h-p.h)); if(d<diff){ diff=d; best=p; } }
  return best.e;
}
function colorEmojiFromHex(hex:string){ return hueToEmoji(rgbToHue(hexToRgb(hex||"#cccccc"))); }
function personaPrefix(p:Persona){ return `${colorEmojiFromHex(p.color)} [${p.name}]`; }

// ---------- Defaults ----------
function P(name:string,system:string,instruction:string,color:string):Persona{
  return { id:name.toLowerCase().replace(/[^a-z0-9]+/g,"-"), enabled:true, name, system, instruction, color };
}
const DEFAULT_SETS: PersonaSet[] = [
  { id:"cross-functional-team", name:"Cross-Functional Team", personas:[
    P("Senior Manager","You are a senior manager prioritizing clarity, risk, outcomes.","Assess clarity of goals, risks, and outcomes.","#fde047"),
    P("Legal","You are corporate counsel focused on compliance and risk.","Flag ambiguous/risky claims; suggest safer wording.","#f9a8d4"),
    P("HR","You are an HR business partner.","Identify exclusionary tone; suggest inclusive language.","#5eead4"),
    P("Technical Lead","You are a pragmatic engineering lead.","Check feasibility, gaps, technical risks.","#93c5fd"),
    P("Junior Analyst","You are a detail-oriented analyst.","Call out unclear logic and missing data.","#86efac"),
  ]},
  { id:"marketing-focus-group", name:"Marketing Focus Group", personas:[
    P("Midwest Parent","Pragmatic parent.","React to clarity, trust, family benefit.","#f59e0b"),
    P("Gen-Z Student","Digital native.","React to tone/authenticity.","#0ea5e9"),
    P("Retired Veteran","Values respect/responsibility.","React to credibility/plain language.","#6d28d9"),
    P("Small Business Owner","Practical ROI.","React to value proposition.","#16a34a"),
    P("Tech-savvy Pro","Precision and specifics.","React to claims needing detail.","#ef4444"),
  ]},
];

function defaultSettings():AppSettings{
  return { provider:{provider:"openrouter",model:"openrouter/auto",openrouterKey:""}, personaSetId:DEFAULT_SETS[0].id, personaSets:DEFAULT_SETS };
}
function loadSettings():AppSettings{ try{ const s=localStorage.getItem(LS_KEY); if(!s) return defaultSettings(); const v=JSON.parse(s) as AppSettings; if(!v.personaSets?.length) v.personaSets=DEFAULT_SETS; if(!v.personaSetId) v.personaSetId=v.personaSets[0].id; return v;}catch{return defaultSettings();} }
function saveSettings(){ localStorage.setItem(LS_KEY, JSON.stringify(SETTINGS)); }
function currentSet(){ const id=SETTINGS.personaSetId; return SETTINGS.personaSets.find(s=>s.id===id)||SETTINGS.personaSets[0]; }

// ---------- UI ----------
function populatePersonaSets(){
  const sets=SETTINGS.personaSets;
  const sel=byId<HTMLSelectElement>("personaSet"); if(sel){ sel.innerHTML=""; sets.forEach(s=>{ const o=document.createElement("option"); o.value=s.id; o.textContent=s.name; sel.appendChild(o); }); sel.value=SETTINGS.personaSetId; }
  const ssel=byId<HTMLSelectElement>("settingsPersonaSet"); if(ssel){ ssel.innerHTML=""; sets.forEach(s=>{ const o=document.createElement("option"); o.value=s.id; o.textContent=s.name; ssel.appendChild(o); }); ssel.value=SETTINGS.personaSetId; }
  renderPersonaNamesAndLegend(); renderPersonaEditor();
}
function renderPersonaNamesAndLegend(){
  const set=currentSet();
  const names=byId<HTMLSpanElement>("personaList"); if(names) names.textContent=set.personas.filter(p=>p.enabled).map(p=>p.name).join(", ");
  const legend=byId<HTMLDivElement>("legend"); if(legend){ legend.innerHTML=""; set.personas.forEach(p=>{ const row=document.createElement("div"); row.style.display="flex"; row.style.alignItems="center"; row.style.gap="6px"; const dot=document.createElement("span"); dot.style.display="inline-block"; dot.style.width="10px"; dot.style.height="10px"; dot.style.borderRadius="50%"; (dot.style as any).background=p.color; row.appendChild(dot); row.appendChild(document.createTextNode(p.name)); legend.appendChild(row); }); }
}
function renderPersonaEditor(){
  const set=currentSet(); const box=byId<HTMLDivElement>("personaEditor"); if(!box) return; box.innerHTML="";
  set.personas.forEach((p,idx)=>{ const card=document.createElement("div"); card.style.border="1px solid #e5e7eb"; card.style.borderRadius="8px"; card.style.padding="8px"; card.style.marginBottom="8px";
    card.innerHTML=`
      <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;">
        <div style="display:flex;gap:8px;align-items:center">
          <input type="checkbox" id="pe-enabled-${idx}" ${p.enabled?"checked":""}/>
          <strong>${p.name}</strong>
        </div>
        <div style="display:flex;gap:6px;align-items:center">
          <label>Color</label><input id="pe-color-${idx}" type="color" value="${p.color}"/>
        </div>
      </div>
      <div style="margin-top:6px;display:flex;gap:8px;align-items:center;"><label style="min-width:90px;">System</label><input style="flex:1" type="text" id="pe-sys-${idx}" value="${escapeHtml(p.system)}"/></div>
      <div style="margin-top:6px;display:flex;gap:8px;align-items:center;"><label style="min-width:90px;">Instruction</label><input style="flex:1" type="text" id="pe-ins-${idx}" value="${escapeHtml(p.instruction)}"/></div>
    `; box.appendChild(card);
  });
}
function hydrateProviderUI(){
  const pr=byId<HTMLSelectElement>("provider"); if(pr) pr.value=SETTINGS.provider.provider;
  const key=byId<HTMLInputElement>("openrouterKey"); if(key) key.value=SETTINGS.provider.openrouterKey||"";
  const mdl=byId<HTMLInputElement>("model"); if(mdl) mdl.value=SETTINGS.provider.model||"";
  const row=byId<HTMLDivElement>("openrouterKeyRow"); if(row) row.classList.toggle("hidden", SETTINGS.provider.provider!=="openrouter");
}

// ---------- Office bootstrap ----------
window.addEventListener("error",e=>log(`[PF] window.error: ${e.message} @ ${e.filename}:${e.lineno}`));
window.addEventListener("unhandledrejection",ev=>log(`[PF] unhandledrejection: ${String(ev.reason)}`));

Office.onReady(async ()=>{
  SETTINGS=loadSettings(); populatePersonaSets(); hydrateProviderUI();
  log("[PF] Office.onReady → UI initialized");

  byId<HTMLButtonElement>("btnSettings")!.onclick=()=>showView("view-settings");
  byId<HTMLButtonElement>("btnBack")!.onclick=()=>showView("view-review");

  const toggleDebug=byId<HTMLButtonElement>("toggleDebug"); if(toggleDebug) toggleDebug.onclick=()=>{ const p=byId<HTMLDivElement>("debugPanel"); if(!p) return; p.classList.toggle("hidden"); toggleDebug.textContent=p.classList.contains("hidden")?"Show Debug":"Hide Debug"; };
  const clearDbg=byId<HTMLButtonElement>("clearDebug"); if(clearDbg) clearDbg.onclick=()=>{ const p=byId<HTMLDivElement>("debugLog"); if(p) p.innerHTML=""; };

  const sel=byId<HTMLSelectElement>("personaSet"); if(sel) sel.onchange=(e)=>{ SETTINGS.personaSetId=(e.target as HTMLSelectElement).value; saveSettings(); populatePersonaSets(); };

  byId<HTMLButtonElement>("runBtn")!.onclick=handleRunReview;
  byId<HTMLButtonElement>("retryBtn")!.onclick=handleRetryFailed;
  byId<HTMLButtonElement>("exportBtn")!.onclick=handleExportPDF;

  const clearBtn=byId<HTMLButtonElement>("clearBtn");
  if(clearBtn) clearBtn.onclick=async()=>{ const ok=await confirmAsync("Clear all comments","Delete ALL comments in this document?"); if(!ok) return; const n=await clearAllComments(); if(n>=0) toast(n>0?`Deleted ${n} comment(s).`:"No comments found."); };

  const prv=byId<HTMLSelectElement>("provider"); if(prv) prv.onchange=(e)=>{ SETTINGS.provider.provider=(e.target as HTMLSelectElement).value as Provider; hydrateProviderUI(); saveSettings(); };
  const key=byId<HTMLInputElement>("openrouterKey"); if(key) key.oninput=(e)=>{ SETTINGS.provider.openrouterKey=(e.target as HTMLInputElement).value; saveSettings(); };
  const mdl=byId<HTMLInputElement>("model"); if(mdl) mdl.oninput=(e)=>{ SETTINGS.provider.model=(e.target as HTMLInputElement).value; saveSettings(); };

  const ssel=byId<HTMLSelectElement>("settingsPersonaSet"); if(ssel) ssel.onchange=(e)=>{ SETTINGS.personaSetId=(e.target as HTMLSelectElement).value; saveSettings(); populatePersonaSets(); };
  const saveSettingsBtn=byId<HTMLButtonElement>("saveSettings"); if(saveSettingsBtn) saveSettingsBtn.onclick=()=>{ const set=currentSet(); set.personas.forEach((p,idx)=>{ const en=byId<HTMLInputElement>(`pe-enabled-${idx}`); const sys=byId<HTMLInputElement>(`pe-sys-${idx}`); const ins=byId<HTMLInputElement>(`pe-ins-${idx}`); const col=byId<HTMLInputElement>(`pe-color-${idx}`); if(en) p.enabled=en.checked; if(sys) p.system=sys.value; if(ins) p.instruction=ins.value; if(col) p.color=col.value||p.color; }); saveSettings(); renderPersonaNamesAndLegend(); toast("Settings saved"); };
  const restoreBtn=byId<HTMLButtonElement>("restoreDefaults"); if(restoreBtn) restoreBtn.onclick=()=>{ const id=currentSet().id; const fresh=DEFAULT_SETS.find(s=>s.id===id); if(fresh){ const i=SETTINGS.personaSets.findIndex(s=>s.id===id); SETTINGS.personaSets[i]=JSON.parse(JSON.stringify(fresh)); saveSettings(); populatePersonaSets(); toast("Default persona set restored"); } };

  showView("view-review");
});

// ---------- Progress badges ----------
function setProgress(p:number){ const bar=byId<HTMLDivElement>("progBar"); if(bar) bar.style.width=`${Math.max(0,Math.min(100,p))}%`; }
function setBadgesHost(personas:Persona[]){ const host=byId<HTMLDivElement>("personaStatus"); if(!host) return; host.innerHTML=""; personas.forEach(p=>{ const row=document.createElement("div"); row.style.display="flex"; row.style.justifyContent="space-between"; row.style.marginBottom="4px"; row.innerHTML=`<span style="display:inline-flex;align-items:center;gap:6px;"><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${p.color};"></span>${p.name}</span><span id="badge-${p.id}" class="badge">queued</span>`; host.appendChild(row);});}
function setBadge(id:string, st:PersonaRunStatus, note?:string){ const b=byId<HTMLSpanElement>(`badge-${id}`); if(!b) return; b.className="badge "+(st==="done"?"badge-done":st==="error"?"badge-failed":""); b.textContent=st+(note?` – ${note}`:""); }

// ---------- Running ----------
async function handleRunReview(){
  if(RUN_LOCK){ toast("Already running…"); return; }
  RUN_LOCK = true;
  try{
    LAST_RESULTS=[]; const res=byId<HTMLDivElement>("results"); if(res) res.innerHTML="";
    const stat=byId<HTMLDivElement>("personaStatus"); if(stat) stat.innerHTML="";
    await runAllEnabledPersonas(false);
  } finally { RUN_LOCK=false; }
}
async function handleRetryFailed(){ if(RUN_LOCK){ toast("Already running…"); return; } RUN_LOCK=true; try{ await runAllEnabledPersonas(true); } finally { RUN_LOCK=false; } }

async function runAllEnabledPersonas(retryOnly:boolean){
  const set=currentSet(); const personas=set.personas.filter(p=>p.enabled);
  if(!personas.length){ toast("No personas enabled in this set."); return; }
  setBadgesHost(personas);
  const docText=await getWholeDocText();
  const total=personas.length; let done=0; setProgress(0);

  for(const p of personas){
    if(retryOnly){ const prev=LAST_RESULTS.find(r=>r.personaId===p.id); if(prev && prev.status==="done"){ done++; setProgress((done/total)*100); continue; } }
    setBadge(p.id,"running");
    try{
      const resp=await callLLMForPersona(p,docText);
      const normalized=normalizeResponse(resp);
      const { matched, unmatched } = await applyCommentsForMatchesOnly(p, normalized);
      addResultCard(p, normalized, unmatched);
      upsertResult({ personaId:p.id, personaName:p.name, status:"done", scores:normalized.scores, global_feedback:normalized.global_feedback, comments:matched, unmatched, raw:resp });
      setBadge(p.id,"done");
    }catch(err:any){
      log(`[PF] Persona ${p.name} error`, err);
      upsertResult({ personaId:p.id, personaName:p.name, status:"error", error:String(err?.message||err) });
      setBadge(p.id,"error", String(err?.message||"LLM call failed"));
    }
    done++; setProgress((done/total)*100);
  }
  toast("Review finished.");
}
function upsertResult(r:PersonaRunResult){ const i=LAST_RESULTS.findIndex(x=>x.personaId===r.personaId); if(i>=0) LAST_RESULTS[i]=r; else LAST_RESULTS.push(r); }

// ---------- Word helpers ----------
async function getWholeDocText():Promise<string>{
  return Word.run(async ctx=>{ const body=ctx.document.body; body.load("text"); await ctx.sync(); return body.text||""; });
}
async function clearAllComments():Promise<number>{
  return Word.run(async ctx=>{
    const coll = (ctx.document as any).comments as Word.CommentCollection | undefined;
    if(!coll || typeof (coll as any).load!=="function"){ toast("This Word build can’t list comments. Use Review → Delete → Delete All Comments."); return -1; }
    coll.load("items"); await ctx.sync(); let n=0; for(const c of coll.items){ c.delete(); n++; } await ctx.sync(); return n;
  });
}

// ---------- Matching-only insertion ----------
function normalizeQuote(s:string){ return (s||"").replace(/[\u2018\u2019\u201A\u201B]/g,"'").replace(/[\u201C\u201D\u201E\u201F]/g,'"').replace(/[\u2013\u2014]/g,"-").replace(/\u00A0/g," ").replace(/\s+/g," ").trim(); }
function middleSlice(s:string,max:number){ if(s.length<=max) return s; const start=Math.max(0,Math.floor((s.length-max)/2)); return s.slice(start,start+max); }
function seedFrom(s:string,which:"first"|"middle"|"last",words:number){ const t=s.split(/\s+/).filter(Boolean); if(t.length<=words) return s; if(which==="first") return t.slice(0,words).join(" "); if(which==="last") return t.slice(-words).join(" "); const mid=Math.floor(t.length/2); const half=Math.floor(words/2); return t.slice(Math.max(0,mid-half), Math.max(0,mid-half)+words).join(" "); }
async function findRangeForQuote(ctx:any, quote:string):Promise<any|null>{
  const body=ctx.document.body; let q=normalizeQuote(quote); if(q.length>260) q=middleSlice(q,180);
  const trySearch = async (needle:string) => { const r=body.search(needle,{matchCase:false,matchWholeWord:false,matchWildcards:false,ignoreSpace:true,ignorePunct:true}); r.load("items"); await ctx.sync(); return r.items.length? r.items[0] : null; };
  let r=await trySearch(q); if(r) return r;
  const dq=q.replace(/^["'“”‘’]+/,"").replace(/["'“”‘’]+$/,"").trim(); if(dq && dq!==q){ r=await trySearch(dq); if(r) return r; }
  for(const w of [8,6,5]) for(const pos of ["first","middle","last"] as const){ const s=seedFrom(q,pos,w); if(!s) continue; r=await trySearch(s); if(r) return r; }
  return null;
}
async function applyCommentsForMatchesOnly(
  persona:Persona,
  data:{scores:{clarity:number;tone:number;alignment:number};global_feedback:string;comments:any[]}
){
  const matched: {quote:string;spanStart:number;spanEnd:number;comment:string}[] = [];
  const unmatched: {quote:string;comment:string}[] = [];
  if(!Array.isArray(data.comments)||!data.comments.length) return { matched, unmatched };

  for(const [i,c] of data.comments.entries()){
    const quote=String(c.quote||"").trim(); const note=String(c.comment||"").trim();
    if(!quote || quote.length<3){ log(`[PF] ${persona.name}: comment #${i+1} empty/short`, c); continue; }

    const placed = await Word.run( async ctx => {
      const r = await findRangeForQuote(ctx, quote);
      if(!r) return false;
      const cm = r.insertComment(`${personaPrefix(persona)} ${note}`);
      cm.load("id"); await ctx.sync(); return true;
    });

    if(placed) matched.push({ quote, spanStart:Number(c.spanStart||0), spanEnd:Number(c.spanEnd||0), comment:note });
    else unmatched.push({ quote, comment:note });
  }
  return { matched, unmatched };
}

// ---------- LLM ----------
function withTimeout<T>(p:Promise<T>, ms=45000){ return new Promise<T>((res,rej)=>{ const t=setTimeout(()=>rej(new Error(`Request timed out after ${ms}ms`)),ms); p.then(v=>{clearTimeout(t);res(v);},e=>{clearTimeout(t);rej(e);}); }); }
async function fetchJson(url:string,init:RequestInit){ const r=await withTimeout(fetch(url,init)); let body:any=null; let text=""; try{ const ct=r.headers.get("content-type")||""; if(ct.includes("application/json")) body=await r.json(); else { text=await r.text(); try{ body=JSON.parse(text);}catch{} } } catch(e){ try{ text=await r.text(); }catch{} } return { ok:r.ok, status:r.status, body, text }; }
async function callLLMForPersona(persona:Persona, docText:string){
  const META_PROMPT = `
Return ONLY valid JSON matching this schema:
{
  "scores":{"clarity":0-100,"tone":0-100,"alignment":0-100},
  "global_feedback":"short paragraph",
  "comments":[{"quote":"verbatim snippet","spanStart":0,"spanEnd":0,"comment":"..."}]
}
Rules: No extra prose; if you output markdown, fence the JSON as \`\`\`json.
`.trim();

  const sys=`${persona.system}\n\n${META_PROMPT}`;
  const user=`You are acting as: ${persona.name}\n\nINSTRUCTION:\n${persona.instruction}\n\nDOCUMENT (plain text):\n${docText}`.trim();
  const pr=SETTINGS.provider; log(`[PF] Calling LLM → ${pr.provider} / ${pr.model} (${persona.name})`);

  if(pr.provider==="openrouter"){
    if(!pr.openrouterKey) throw new Error("Missing OpenRouter API key.");
    const res=await fetchJson("https://openrouter.ai/api/v1/chat/completions",{
      method:"POST",
      headers:{ "Content-Type":"application/json", "Authorization":`Bearer ${pr.openrouterKey}`, "HTTP-Referer": (typeof window!=="undefined"?window.location.origin:"https://word-persona-feedback.vercel.app"), "X-Title":"Persona Feedback Add-in" },
      body:JSON.stringify({ model: pr.model||"openrouter/auto", messages:[{role:"system",content:sys},{role:"user",content:user}], temperature:0.2 })
    });
    if(!res.ok) throw new Error(`OpenRouter HTTP ${res.status}: ${res.text || safeJson(res.body)}`);
    const content=res.body?.choices?.[0]?.message?.content ?? ""; log("[PF] OpenRouter raw", res.body);
    return parseJsonFromText(content);
  }else{
    const res=await fetchJson("http://127.0.0.1:11434/api/chat",{
      method:"POST", headers:{ "Content-Type":"application/json" },
      body:JSON.stringify({ model: pr.model||"llama3", stream:false, messages:[{role:"system",content:sys},{role:"user",content:user}], options:{temperature:0.2} })
    });
    if(!res.ok) throw new Error(`Ollama HTTP ${res.status}: ${res.text || safeJson(res.body)}`);
    const content=res.body?.message?.content ?? ""; log("[PF] Ollama raw", res.body);
    return parseJsonFromText(content);
  }
}
function parseJsonFromText(text:string){ const m=text.match(/```json([\s\S]*?)```/i) || text.match(/```([\s\S]*?)```/); const raw=m?m[1]:text; try{ return JSON.parse(raw.trim()); }catch{ log("[PF] JSON parse error; full text follows",{text}); throw new Error("Model returned non-JSON. See Debug for raw output."); } }
function normalizeResponse(resp:any){ const clamp=(n:number)=>Math.max(0,Math.min(100,Math.round(n))); return {
  scores:{ clarity:clamp(Number(resp?.scores?.clarity??0)), tone:clamp(Number(resp?.scores?.tone??0)), alignment:clamp(Number(resp?.scores?.alignment??0)) },
  comments:Array.isArray(resp?.comments)?resp.comments.slice(0,12):[],
  global_feedback:String(resp?.global_feedback||"")
}; }

// ---------- Results UI ----------
function barColor(v:number){ if(v>=80) return "#16a34a"; if(v<50) return "#dc2626"; return "#eab308"; }
function scoreBar(label:string,val:number){ const pct=Math.max(0,Math.min(100,val|0)); const c=barColor(pct); return `
  <div style="display:flex;justify-content:space-between;font-size:12px;margin-top:6px;"><span>${label}</span><span>${pct}</span></div>
  <div style="width:100%;height:14px;background:#e5e7eb;border-radius:999px;overflow:hidden;"><div style="height:100%;width:${pct}%;background:${c};"></div></div>`; }

function addResultCard(persona:Persona, data:{scores:{clarity:number;tone:number;alignment:number};global_feedback:string}, unmatched?:{quote:string;comment:string}[]){
  const host=byId<HTMLDivElement>("results"); if(!host) return;
  const card=document.createElement("div"); card.style.border="1px solid #e5e7eb"; card.style.borderRadius="10px"; card.style.padding="10px"; card.style.marginBottom="10px";
  const {clarity,tone,alignment}=data.scores;
  card.innerHTML=`
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
      <div style="display:flex;align-items:center;gap:8px;">
        <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${persona.color};"></span>
        <strong>${persona.name}</strong>
      </div>
      <span class="badge badge-done">done</span>
    </div>
    ${scoreBar("Clarity",clarity)}${scoreBar("Tone",tone)}${scoreBar("Alignment",alignment)}
    <div style="margin-top:8px;"><em>${escapeHtml(data.global_feedback)}</em></div>
    ${unmatched&&unmatched.length?`
      <div style="margin-top:8px;">
        <div style="font-weight:600;margin-bottom:4px;">Unmatched quotes (not inserted):</div>
        <ul style="margin:0 0 0 16px;padding:0;list-style:disc;">
          ${unmatched.slice(0,6).map(u=>`<li><span style="color:#6b7280">"${escapeHtml(u.quote.slice(0,160))}${u.quote.length>160?"…":""}"</span><br/><span>${escapeHtml(u.comment)}</span></li>`).join("")}
        </ul>
      </div>`:""}
  `; host.appendChild(card);
}

// ---------- Export (PDF via print) ----------
function buildReportHtml(){
  const setName=currentSet().name;
  const rows = LAST_RESULTS.map(r=>{
    const persona = currentSet().personas.find(p=>p.id===r.personaId);
    const color = persona?.color || "#93c5fd";
    const s = r.scores || {clarity:0,tone:0,alignment:0};
    return `
      <section style="border:1px solid #e5e7eb;border-radius:10px;padding:14px;margin:10px 0;">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
          <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${color};"></span>
          <strong>${escapeHtml(r.personaName)}</strong>
          <span style="margin-left:auto;padding:2px 8px;border-radius:999px;background:#e5ffe8;border:1px solid #a7f3d0;">${r.status}</span>
        </div>
        ${scoreBar("Clarity",s.clarity)}${scoreBar("Tone",s.tone)}${scoreBar("Alignment",s.alignment)}
        ${r.global_feedback?`<div style="margin-top:8px;"><em>${escapeHtml(r.global_feedback)}</em></div>`:""}
      </section>`;
  }).join("");

  return `<!doctype html><html><head><meta charset="utf-8"><title>Persona Feedback Report</title>
  <style>body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif; padding:20px; max-width:800px; margin:0 auto;} h1{font-size:20px;margin:0 0 10px;} .muted{color:#6b7280}</style>
  </head><body>
    <h1>Persona Feedback Report</h1>
    <div class="muted">${new Date().toLocaleString()} • Set: ${escapeHtml(setName)}</div>
    ${rows || "<p class='muted'>No results.</p>"}
    <script>setTimeout(()=>window.print(), 300);</script>
  </body></html>`;
}
function handleExportPDF(){
  const html = buildReportHtml();
  const blob = new Blob([html], {type:"text/html"});
  const url = URL.createObjectURL(blob);
  window.open(url, "_blank");
}

// ---------- Helpers ----------
function upsert<T extends { personaId:string }>(arr:T[], item:T){ const i=arr.findIndex(x=>x.personaId===item.personaId); if(i>=0) arr[i]=item; else arr.push(item); }
function upsertResultHelper(r:PersonaRunResult){ upsert(LAST_RESULTS, r); }
// Keep name consistency used above:
function upsertResult(r:PersonaRunResult){ upsertResultHelper(r); }
