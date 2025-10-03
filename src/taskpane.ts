/* eslint-disable @typescript-eslint/no-explicit-any */

/**
 * Persona Feedback – Word Add-in task pane
 * - Run lock to prevent duplicate runs (fixes double summaries)
 * - Single onclick wiring (no stacked listeners)
 * - Comments ONLY (no highlights/edits)
 * - Inline color badge in each comment bubble (no "PF" marker)
 * - "Clear all comments" works where the Word build exposes comments collection
 * - Robust quote matching (normalize punctuation/whitespace + extra fallbacks)
 * - Bar charts + debug panel
 */

type Provider = "openrouter" | "ollama";
type Persona = {
  id: string; enabled: boolean; name: string;
  system: string; instruction: string; color: string;
};
type PersonaSet = { id: string; name: string; personas: Persona[] };
type ProviderSettings = { provider: Provider; openrouterKey?: string; model: string };
type AppSettings = { provider: ProviderSettings; personaSetId: string; personaSets: PersonaSet[] };

type PersonaRunStatus = "idle" | "running" | "done" | "error";
type PersonaRunResult = {
  personaId: string; personaName: string; status: PersonaRunStatus;
  scores?: { clarity: number; tone: number; alignment: number };
  global_feedback?: string;
  comments?: { quote: string; spanStart: number; spanEnd: number; comment: string }[];
  unmatched?: { quote: string; comment: string }[];
  raw?: any; error?: string;
};

// ---------- Constants / Globals ----------
const DOT_COLORS = ["#fde047","#f9a8d4","#5eead4","#f87171","#93c5fd","#86efac","#c4b5fd","#f59e0b","#0d9488","#b91c1c","#1d4ed8","#166534","#6d28d9"];
const LS_KEY = "pf.settings.v1";

let SETTINGS: AppSettings;
let LAST_RESULTS: PersonaRunResult[] = [];
let SESSION_COMMENT_IDS: string[] = [];
let RUN_LOCK = false; // prevents duplicate runs

// ---------- Utilities ----------
function byId<T extends HTMLElement>(id: string): T | null { return document.getElementById(id) as T | null; }
function req<T extends HTMLElement>(id: string): T { const el = byId<T>(id); if (!el) throw new Error(`#${id} missing`); return el; }
function safeJson(x: any){ try{ return JSON.stringify(x,null,2);}catch{ return String(x);} }
function log(msg: string, data?: any){ const p=byId<HTMLDivElement>("debugLog"); if(data!==undefined) console.log(msg,data); else console.log(msg); if(!p) return; const d=document.createElement("div"); d.style.whiteSpace="pre-wrap"; d.textContent=data?`${msg} ${safeJson(data)}`:msg; p.appendChild(d); p.scrollTop=p.scrollHeight; }
function toast(t: string){ const box=byId<HTMLDivElement>("toast"); if(!box) return; const msg=byId<HTMLSpanElement>("toastMsg"); if(msg) msg.textContent=t; box.style.display="block"; const close=byId<HTMLSpanElement>("toastClose"); if(close) close.onclick=()=>box.style.display="none"; setTimeout(()=>box.style.display="none",2000); }
function showView(id:"view-review"|"view-settings"){ const r=byId<HTMLDivElement>("view-review"); const s=byId<HTMLDivElement>("view-settings"); const b=byId<HTMLButtonElement>("btnBack"); if(id==="view-review"){ r&&r.classList.remove("hidden"); s&&s.classList.add("hidden"); b&&b.classList.add("hidden"); }else{ r&&r.classList.add("hidden"); s&&s.classList.remove("hidden"); b&&b.classList.remove("hidden"); } }
function confirmAsync(title:string,message:string){ return new Promise<boolean>(res=>{ const o=req<HTMLDivElement>("confirmOverlay"); req<HTMLHeadingElement>("confirmTitle").textContent=title; req<HTMLDivElement>("confirmMessage").textContent=message; o.style.display="flex"; const ok=req<HTMLButtonElement>("confirmOk"); const no=req<HTMLButtonElement>("confirmCancel"); const done=(v:boolean)=>{o.style.display="none"; ok.onclick=null; no.onclick=null; res(v);}; ok.onclick=()=>done(true); no.onclick=()=>done(false);}); }
function escapeHtml(s:string){ return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }
function colorAt(i:number){ return DOT_COLORS[i % DOT_COLORS.length]; }
function colorBadgeFor(hex:string){ const h=(hex||"").toLowerCase(); const map:[RegExp,string][]= [
  [/(f59e0b|f97316|fb923c|fdba74)/,"🟧"], [/(fde047|facc15|eab308|fff59d)/,"🟨"], [/(86efac|34d399|10b981|22c55e|16a34a)/,"🟩"],
  [/(93c5fd|60a5fa|3b82f6|2563eb|1d4ed8)/,"🟦"], [/(c4b5fd|a78bfa|8b5cf6|6d28d9)/,"🟪"], [/(f87171|ef4444|dc2626|b91c1c)/,"🟥"]
]; for(const [rx,b] of map) if(rx.test(h)) return b; return "⬛"; }
function personaPrefix(p:{name:string;color:string}){ return `${colorBadgeFor(p.color)} [${p.name}]`; }

// ---------- Defaults ----------
const META_PROMPT = `
Return ONLY valid JSON matching this schema:
{
  "scores":{"clarity":0-100,"tone":0-100,"alignment":0-100},
  "global_feedback":"short paragraph",
  "comments":[{"quote":"verbatim snippet","spanStart":0,"spanEnd":0,"comment":"..."}]
}
Rules: No extra prose; if you output markdown, fence the JSON as \`\`\`json.
`;

function P(name:string,system:string,instruction:string,idx:number):Persona{
  return { id:name.toLowerCase().replace(/[^a-z0-9]+/g,"-"), enabled:true, name, system, instruction, color:colorAt(idx) };
}
const DEFAULT_SETS: PersonaSet[] = [
  { id:"cross-functional-team", name:"Cross-Functional Team", personas:[
    P("Senior Manager","You are a senior manager prioritizing clarity, risk, outcomes.","Assess clarity of goals, risks, and outcomes.",0),
    P("Legal","You are corporate counsel focused on compliance and risk.","Flag ambiguous/risky claims; suggest safer wording.",1),
    P("HR","You are an HR business partner.","Identify exclusionary tone; suggest inclusive language.",2),
    P("Technical Lead","You are a pragmatic engineering lead.","Check feasibility, gaps, technical risks.",3),
    P("Junior Analyst","You are a detail-oriented analyst.","Call out unclear logic and missing data.",4),
  ]},
  { id:"marketing-focus-group", name:"Marketing Focus Group", personas:[
    P("Midwest Parent","Pragmatic parent.","React to clarity, trust, family benefit.",0),
    P("Gen-Z Student","Digital native.","React to tone/authenticity.",1),
    P("Retired Veteran","Values respect/responsibility.","React to credibility/plain language.",2),
    P("Small Business Owner","Practical ROI.","React to value proposition.",3),
    P("Tech-savvy Pro","Precision and specifics.","React to claims needing detail.",4),
  ]},
  { id:"startup-stakeholders", name:"Startup Stakeholders", personas:[
    P("Founder","Vision & cadence.","Assess clarity of focus and milestones.",0),
    P("CTO","Architecture & risk.","Probe feasibility and tradeoffs.",1),
    P("CMO","Messaging & audience.","Probe differentiation and clarity.",2),
    P("VC Investor","Metrics & risks.","Probe runway, milestones, and risk.",3),
    P("Customer","Buyer value.","Probe outcomes and adoption risks.",4),
  ]},
  { id:"political-spectrum", name:"Political Spectrum", personas:[
    P("Democratic Socialist","Equity & ethics.","Assess public benefit and guardrails.",0),
    P("Center Left","Policy realism.","Assess social impact and feasibility.",1),
    P("Centrist/Independent","Balance & fairness.","Assess tradeoffs and balance.",2),
    P("Center Right","Fiscal prudence.","Assess efficiency and discipline.",3),
    P("MAGA","Populist conservative.","Assess national interest and clarity.",4),
    P("Libertarian","Freedom & regulation.","Assess individual liberty burdens.",5),
  ]},
  { id:"product-review-board", name:"Product Review Board", personas:[
    P("PM","Problem/Solution/Metric.","Assess scope and success criteria.",0),
    P("Design Lead","UX & access.","Assess flows, inclusivity.",1),
    P("Data Scientist","Validity & risk.","Assess metrics and data risks.",2),
    P("Security","Privacy & threat model.","Assess handling and mitigations.",3),
    P("Support Lead","Edge cases.","Assess failures and comms.",4),
  ]},
  { id:"scientific-peer-review", name:"Scientific Peer Review", personas:[
    P("Methods Reviewer","Methods & reproducibility.","Check detail and validity.",0),
    P("Stats Reviewer","Statistics.","Check samples, tests, uncertainty.",1),
    P("Domain Expert","Accuracy.","Check assumptions/citations.",2),
    P("Ethics Reviewer","Ethical impact.","Check risk and consent.",3),
  ]},
  { id:"ux-research-panel", name:"UX Research Panel", personas:[
    P("New User","First-time use.","Assess onboarding clarity.",0),
    P("Power User","Expert flow.","Assess efficiency/discoverability.",1),
    P("Accessibility Advocate","A11y.","Assess contrast/semantics/alt text.",2),
  ]},
  { id:"sales-deal-desk", name:"Sales Deal Desk", personas:[
    P("Sales Director","Top-line growth.","Assess messaging and objections.",0),
    P("Solutions Architect","Technical fit.","Assess constraints/integrations.",1),
    P("Legal (Customer)","Counsel.","Assess indemnities, data use, SLAs.",2),
    P("Procurement","Buyer procurement.","Assess pricing clarity/comparables.",3),
  ]},
  { id:"board-of-directors", name:"Board of Directors", personas:[
    P("Chair","Governance.","Assess strategy coherence and oversight.",0),
    P("Audit","Audit committee.","Assess controls and reporting.",1),
    P("Compensation","Comp committee.","Assess incentives/fairness.",2),
  ]},
  { id:"academic-committee", name:"Academic Committee", personas:[
    P("Dean","Leadership.","Assess rigor & mission alignment.",0),
    P("IRB Chair","Research ethics.","Assess consent & risk.",1),
    P("Funding Reviewer","Grant committee.","Assess merit/feasibility/budget.",2),
  ]},
];

function defaultSettings(): AppSettings {
  return { provider:{provider:"openrouter",model:"openrouter/auto",openrouterKey:""}, personaSetId:DEFAULT_SETS[0].id, personaSets:DEFAULT_SETS };
}
function loadSettings(): AppSettings {
  try{ const s=localStorage.getItem(LS_KEY); if(!s) return defaultSettings();
    const parsed = JSON.parse(s) as AppSettings;
    if(!parsed.personaSets?.length) parsed.personaSets=DEFAULT_SETS;
    if(!parsed.personaSetId) parsed.personaSetId=DEFAULT_SETS[0].id;
    return parsed;
  }catch{ return defaultSettings(); }
}
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
  const set=currentSet(); const names=byId<HTMLSpanElement>("personaList"); if(names) names.textContent=set.personas.filter(p=>p.enabled).map(p=>p.name).join(", ");
  const legend=byId<HTMLDivElement>("legend"); if(legend){ legend.innerHTML=""; set.personas.forEach(p=>{ const row=document.createElement("div"); row.style.display="flex"; row.style.alignItems="center"; row.style.gap="6px"; const dot=document.createElement("span"); dot.style.display="inline-block"; dot.style.width="10px"; dot.style.height="10px"; dot.style.borderRadius="50%"; (dot.style as any).background=p.color; row.appendChild(dot); row.appendChild(document.createTextNode(p.name)); legend.appendChild(row); }); }
}
function toHexColor(c:string){ if(c.startsWith("#")) return c; const map:Record<string,string>={yellow:"#fde047",pink:"#f9a8d4",turquoise:"#5eead4",red:"#f87171",blue:"#93c5fd",green:"#86efac",violet:"#c4b5fd",orange:"#f59e0b"}; return map[c]||"#fde047"; }
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
          <label>Color</label><input id="pe-color-${idx}" type="color" value="${toHexColor(p.color)}"/>
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
window.addEventListener("error",(e)=>log(`[PF] window.error: ${e.message} @ ${e.filename}:${e.lineno}`));
window.addEventListener("unhandledrejection",(ev)=>log(`[PF] unhandledrejection: ${String(ev.reason)}`));

Office.onReady(async ()=>{
  log("[PF] Office.onReady → UI initialized");
  SETTINGS=loadSettings(); populatePersonaSets(); hydrateProviderUI();

  // One-time wiring via onclick to avoid stacking
  const btnSettings = byId<HTMLButtonElement>("btnSettings"); if(btnSettings) btnSettings.onclick=()=>showView("view-settings");
  const btnBack = byId<HTMLButtonElement>("btnBack"); if(btnBack) btnBack.onclick=()=>showView("view-review");

  const toggleDebug = byId<HTMLButtonElement>("toggleDebug");
  if (toggleDebug) toggleDebug.onclick=()=>{ const p=byId<HTMLDivElement>("debugPanel"); if(!p) return; p.classList.toggle("hidden"); toggleDebug.textContent=p.classList.contains("hidden")?"Show Debug":"Hide Debug"; };
  const clearDbg = byId<HTMLButtonElement>("clearDebug"); if(clearDbg) clearDbg.onclick=()=>{ const p=byId<HTMLDivElement>("debugLog"); if(p) p.innerHTML=""; };

  const sel=byId<HTMLSelectElement>("personaSet"); if(sel) sel.onchange=(ev)=>{ SETTINGS.personaSetId=(ev.target as HTMLSelectElement).value; saveSettings(); populatePersonaSets(); };
  const runBtn=byId<HTMLButtonElement>("runBtn"); if(runBtn) runBtn.onclick=handleRunReview;
  const retryBtn=byId<HTMLButtonElement>("retryBtn"); if(retryBtn) retryBtn.onclick=handleRetryFailed;
  const exportBtn=byId<HTMLButtonElement>("exportBtn"); if(exportBtn) exportBtn.onclick=handleExportReport;

  const clearBtn=byId<HTMLButtonElement>("clearBtn"); if(clearBtn) clearBtn.onclick=async()=>{ const ok=await confirmAsync("Clear all comments","Delete ALL comments in this document?"); if(!ok) return; const n=await clearAllComments(); if(n>=0) toast(n>0?`Deleted ${n} comment(s).`:"No comments found."); };

  const prv=byId<HTMLSelectElement>("provider"); if(prv) prv.onchange=(e)=>{ SETTINGS.provider.provider=(e.target as HTMLSelectElement).value as Provider; hydrateProviderUI(); saveSettings(); };
  const key=byId<HTMLInputElement>("openrouterKey"); if(key) key.oninput=(e)=>{ SETTINGS.provider.openrouterKey=(e.target as HTMLInputElement).value; saveSettings(); };
  const mdl=byId<HTMLInputElement>("model"); if(mdl) mdl.oninput=(e)=>{ SETTINGS.provider.model=(e.target as HTMLInputElement).value; saveSettings(); };

  const ssel=byId<HTMLSelectElement>("settingsPersonaSet"); if(ssel) ssel.onchange=(e)=>{ SETTINGS.personaSetId=(e.target as HTMLSelectElement).value; saveSettings(); populatePersonaSets(); };
  const saveSettingsBtn=byId<HTMLButtonElement>("saveSettings"); if(saveSettingsBtn) saveSettingsBtn.onclick=()=>{ const set=currentSet(); set.personas.forEach((p,idx)=>{ const en=byId<HTMLInputElement>(`pe-enabled-${idx}`); const sys=byId<HTMLInputElement>(`pe-sys-${idx}`); const ins=byId<HTMLInputElement>(`pe-ins-${idx}`); const col=byId<HTMLInputElement>(`pe-color-${idx}`); if(en) p.enabled=en.checked; if(sys) p.system=sys.value; if(ins) p.instruction=ins.value; if(col) p.color=col.value||p.color; }); saveSettings(); renderPersonaNamesAndLegend(); toast("Settings saved"); };
  const restoreBtn=byId<HTMLButtonElement>("restoreDefaults"); if(restoreBtn) restoreBtn.onclick=()=>{ const id=currentSet().id; const fresh=DEFAULT_SETS.find(s=>s.id===id); if(fresh){ const i=SETTINGS.personaSets.findIndex(s=>s.id===id); SETTINGS.personaSets[i]=JSON.parse(JSON.stringify(fresh)); saveSettings(); populatePersonaSets(); toast("Default persona set restored"); } };

  showView("view-review");
});

// ---------- Actions ----------
async function handleRunReview(){
  if (RUN_LOCK) { toast("Already running…"); return; }
  RUN_LOCK = true;
  try{
    LAST_RESULTS=[]; SESSION_COMMENT_IDS=[];
    const res=byId<HTMLDivElement>("results"); if(res) res.innerHTML="";
    const stat=byId<HTMLDivElement>("personaStatus"); if(stat) stat.innerHTML="";
    await runAllEnabledPersonas(false);
  } finally { RUN_LOCK = false; }
}
async function handleRetryFailed(){ if(RUN_LOCK){ toast("Already running…"); return; } RUN_LOCK=true; try{ await runAllEnabledPersonas(true); } finally{ RUN_LOCK=false; } }

function setProgress(p:number){ const bar=byId<HTMLDivElement>("progBar"); if(bar) bar.style.width=`${Math.max(0,Math.min(100,p))}%`; }
function setBadgesHost(personas:Persona[]){ const host=byId<HTMLDivElement>("personaStatus"); if(!host) return; host.innerHTML=""; personas.forEach(p=>{ const row=document.createElement("div"); row.style.display="flex"; row.style.justifyContent="space-between"; row.style.marginBottom="4px"; row.innerHTML=`<span style="display:inline-flex;align-items:center;gap:6px;"><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${p.color};"></span>${p.name}</span><span id="badge-${p.id}" class="badge">queued</span>`; host.appendChild(row);});}
function setBadge(id:string, st:PersonaRunStatus, note?:string){ const b=byId<HTMLSpanElement>(`badge-${id}`); if(!b) return; b.className="badge "+(st==="done"?"badge-done":st==="error"?"badge-failed":""); b.textContent=st+(note?` – ${note}`:""); }

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

async function addCommentAtStart(persona:Persona, text:string){
  return Word.run(async ctx=>{
    const start=ctx.document.body.getRange("Start");
    const c=start.insertComment(`${personaPrefix(persona)} ${text}`);
    c.load("id"); await ctx.sync(); if(c.id) SESSION_COMMENT_IDS.push(c.id);
  });
}

// Delete ALL comments in document (preferred)
async function clearAllComments():Promise<number>{
  return Word.run(async ctx=>{
    const docAny:any = ctx.document as any;
    const coll = docAny.comments;
    if(!coll || typeof coll.load!=="function"){ toast("This Word build can’t list comments. Use Review → Delete → Delete All Comments."); return -1; }
    coll.load("items"); await ctx.sync(); let n=0; for(const c of coll.items){ c.delete(); n++; } await ctx.sync(); SESSION_COMMENT_IDS=[]; return n;
  });
}

// ---------- Matching ----------
function normalizeQuote(s:string){ return (s||"").replace(/[\u2018\u2019\u201A\u201B]/g,"'").replace(/[\u201C\u201D\u201E\u201F]/g,'"').replace(/[\u2013\u2014]/g,"-").replace(/\u00A0/g," ").replace(/\s+/g," ").trim(); }
function middleSlice(s:string,max:number){ if(s.length<=max) return s; const start=Math.max(0,Math.floor((s.length-max)/2)); return s.slice(start,start+max); }
function seedFrom(s:string, which:"first"|"middle"|"last", words:number){
  const t=s.split(/\s+/).filter(Boolean); if(t.length<=words) return s;
  if(which==="first") return t.slice(0,words).join(" ");
  if(which==="last") return t.slice(-words).join(" ");
  const mid=Math.floor(t.length/2); const half=Math.floor(words/2); return t.slice(Math.max(0,mid-half), Math.max(0,mid-half)+words).join(" ");
}

async function findRangeForQuote(ctx:any, quote:string):Promise<any|null>{
  const body=ctx.document.body; let q=normalizeQuote(quote); if(q.length>260) q=middleSlice(q,180);

  const trySearch = async (needle:string) => {
    let r = body.search(needle, { matchCase:false, matchWholeWord:false, matchWildcards:false, ignoreSpace:true, ignorePunct:true });
    r.load("items"); await ctx.sync(); return r.items.length>0 ? r.items[0] : null;
  };

  // Full quote
  let r = await trySearch(q); if(r) return r;
  // Dequoted
  const dq=q.replace(/^["'“”‘’]+/,"").replace(/["'“”‘’]+$/,"").trim(); if(dq && dq!==q){ r=await trySearch(dq); if(r) return r; }
  // First / middle / last seeds
  for(const w of [8,6,5]){
    for(const pos of ["first","middle","last"] as const){
      const s=seedFrom(q,pos,w); if(!s) continue; r=await trySearch(s); if(r) return r;
    }
  }
  return null;
}

async function applyCommentsForMatchesOnly(
  persona:Persona,
  data:{scores:{clarity:number;tone:number;alignment:number};global_feedback:string;comments:any[]}
){
  await addCommentAtStart(persona, `Summary (${persona.name}): ${data.global_feedback}`);

  const matched: {quote:string;spanStart:number;spanEnd:number;comment:string}[] = [];
  const unmatched: {quote:string;comment:string}[] = [];

  if(!Array.isArray(data.comments)||!data.comments.length) return { matched, unmatched };

  for(const [i,c] of data.comments.entries()){
    const quote=String(c.quote||"").trim(); const note=String(c.comment||"").trim();
    if(!quote || quote.length<3){ log(`[PF] ${persona.name}: comment #${i+1} empty/short quote`, c); continue; }

    const placed = await Word.run( async ctx => {
      const r = await findRangeForQuote(ctx, quote);
      if(!r) return false;
      const cm = r.insertComment(`${personaPrefix(persona)} ${note}`); cm.load("id"); await ctx.sync(); if(cm.id) SESSION_COMMENT_IDS.push(cm.id); return true;
    });

    if(placed) matched.push({ quote, spanStart:Number(c.spanStart||0), spanEnd:Number(c.spanEnd||0), comment:note });
    else unmatched.push({ quote, comment:note });
  }
  return { matched, unmatched };
}

// ---------- Networking / LLM ----------
function withTimeout<T>(p:Promise<T>, ms=45000){ return new Promise<T>((res,rej)=>{ const t=setTimeout(()=>rej(new Error(`Request timed out after ${ms}ms`)),ms); p.then(v=>{clearTimeout(t);res(v);},e=>{clearTimeout(t);rej(e);}); }); }
async function fetchJson(url:string,init:RequestInit){ const r=await withTimeout(fetch(url,init)); let body:any=null; let text=""; try{ const ct=r.headers.get("content-type")||""; if(ct.includes("application/json")) body=await r.json(); else { text=await r.text(); try{ body=JSON.parse(text);}catch{} } } catch(e){ try{ text=await r.text(); }catch{} } return { ok:r.ok, status:r.status, body, text }; }
async function callLLMForPersona(persona:Persona, docText:string){
  const sys=`${persona.system}\n\n${META_PROMPT}`.trim();
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
    const content=res.body?.choices?.[0]?.message?.content ?? ""; log("[PF] OpenRouter raw", res.body); return parseJsonFromText(content);
  }else{
    const res=await fetchJson("http://127.0.0.1:11434/api/chat",{
      method:"POST", headers:{ "Content-Type":"application/json" },
      body:JSON.stringify({ model: pr.model||"llama3", stream:false, messages:[{role:"system",content:sys},{role:"user",content:user}], options:{temperature:0.2} })
    });
    if(!res.ok) throw new Error(`Ollama HTTP ${res.status}: ${res.text || safeJson(res.body)}`);
    const content=res.body?.message?.content ?? ""; log("[PF] Ollama raw", res.body); return parseJsonFromText(content);
  }
}
function parseJsonFromText(text:string){ const m=text.match(/```json([\s\S]*?)```/i) || text.match(/```([\s\S]*?)```/); const raw=m?m[1]:text; try{ return JSON.parse(raw.trim()); }catch{ log("[PF] JSON parse error; full text follows",{text}); throw new Error("Model returned non-JSON. See Debug for raw output."); } }
function normalizeResponse(resp:any){ const clamp=(n:number)=>Math.max(0,Math.min(100,Math.round(n))); return {
  scores:{ clarity:clamp(Number(resp?.scores?.clarity??0)), tone:clamp(Number(resp?.scores?.tone??0)), alignment:clamp(Number(resp?.scores?.alignment??0)) },
  comments:Array.isArray(resp?.comments)?resp.comments.slice(0,12):[],
  global_feedback:String(resp?.global_feedback||"")
}; }

// ---------- Results UI ----------
function scoreBar(label:string,val:number){ const pct=Math.max(0,Math.min(100,val|0)); return `
  <div style="display:flex;justify-content:space-between;font-size:12px;margin-top:4px;"><span>${label}</span><span>${pct}</span></div>
  <div style="width:100%;height:8px;background:#e5e7eb;border-radius:999px;overflow:hidden;"><div style="height:100%;width:${pct}%;background:#3b82f6;"></div></div>`; }
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
async function handleExportReport(){ const payload={ timestamp:new Date().toISOString(), set:currentSet().name, results:LAST_RESULTS, model:SETTINGS.provider }; const blob=new Blob([JSON.stringify(payload,null,2)],{type:"application/json"}); const url=URL.createObjectURL(blob); const a=document.createElement("a"); a.href=url; a.download=`persona-feedback-${Date.now()}.json`; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url); }
