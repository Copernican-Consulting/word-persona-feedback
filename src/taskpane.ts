/* global Office, Word */
import { DEFAULT_SETS, PersonaSet, Persona } from "./personas";

type Settings = {
  provider: "openrouter" | "ollama";
  openrouterKey?: string;
  model: string;
  personaSets: Record<string, PersonaSet>;
  activeSetId: string;
};

type PersonaRunState = "queued" | "running" | "done" | "failed";

// ===== storage version bump to force merge of default sets
const STORAGE_KEY = "pf_settings_v4";

/* ---------- utils ---------- */
const clone = <T,>(o: T): T => JSON.parse(JSON.stringify(o));
const esc = (s: string) => s.replace(/"/g, "&quot;");
function debug(...a: any[]) {
  try {
    console.log("[PF]", ...a);
    const dbg = document.getElementById("debugLog") as HTMLDivElement | null;
    if (dbg) {
      const line = document.createElement("div");
      line.textContent = a.map(x => (typeof x === "string" ? x : JSON.stringify(x))).join(" ");
      dbg.appendChild(line); dbg.scrollTop = dbg.scrollHeight;
    }
  } catch {}
}
function showToast(msg: string) {
  const toast = document.getElementById("toast") as HTMLDivElement | null;
  const msgEl = document.getElementById("toastMsg") as HTMLSpanElement | null;
  const close = document.getElementById("toastClose") as HTMLSpanElement | null;
  if (!toast || !msgEl || !close) return alert(msg);
  msgEl.textContent = msg;
  toast.style.display = "block";
  const hide = () => { toast.style.display = "none"; close.removeEventListener("click", hide); };
  close.addEventListener("click", hide); setTimeout(hide, 2000);
}

/* ---------- color helpers ---------- */
type HL =
  | "NoColor" | "Yellow" | "Pink" | "BrightGreen" | "Turquoise" | "LightGray" | "Violet"
  | "DarkYellow" | "DarkBlue" | "DarkRed" | "Teal" | "Brown" | "DarkGreen" | "DarkTeal"
  | "Indigo" | "Orange" | "Blue" | "Red" | "Green" | "Black" | "Gray25" | "Gray50";

function hexToRgb(hex?: string): [number, number, number] {
  if (!hex) return [0,0,0];
  const s = hex.replace("#",""); return [0,2,4].map(i => parseInt(s.slice(i,i+2),16)) as any;
}
function hexToHighlightColor(hex?: string): HL {
  const [r,g,b] = hexToRgb(hex);
  const max = Math.max(r,g,b), min = Math.min(r,g,b);
  // quick category guesses
  if (r>220 && g>210 && b<140) return "Yellow";
  if (r>230 && g<160 && b<160) return "Pink";
  if (g>210 && r<150 && b<150) return "BrightGreen";
  if (g>180 && b>180 && r<120) return "Turquoise";
  if (r>230 && g>170 && b<120) return "Orange";
  if (b>200 && r<120) return "Blue";
  if (r>200 && g<120 && b<120) return "Red";
  if (max-min<20 && max>200) return "LightGray";
  return "NoColor";
}

function colorToEmoji(hex?: string): string {
  const [r,g,b] = hexToRgb(hex);
  // pick closest of these tags
  const palette = [
    {e:"🟥",c:[220,70,70]},
    {e:"🟧",c:[240,150,50]},
    {e:"🟨",c:[240,220,80]},
    {e:"🟩",c:[60,180,90]},
    {e:"🟦",c:[70,120,230]},
    {e:"🟪",c:[150,80,220]},
    {e:"🟫",c:[140,90,60]},
    {e:"⬛",c:[30,30,30]},
    {e:"⬜",c:[230,230,230]},
  ];
  let best=palette[0], dmin=1e9;
  for (const p of palette) {
    const d=(r-p.c[0])**2+(g-p.c[1])**2+(b-p.c[2])**2;
    if (d<dmin){dmin=d;best=p;}
  }
  return best.e;
}

/* ---------- state ---------- */
let settings: Settings;

/* ---------- persistence with merge ---------- */
function loadSettings(): Settings {
  const fromStorage = localStorage.getItem(STORAGE_KEY);
  let base: Settings;

  if (fromStorage) {
    base = JSON.parse(fromStorage);
  } else {
    const personaSets: Record<string, PersonaSet> = {};
    DEFAULT_SETS.forEach(s => (personaSets[s.id] = clone(s)));
    base = {
      provider: "openrouter",
      openrouterKey: "",
      model: "openrouter/auto",
      personaSets,
      activeSetId: DEFAULT_SETS[0].id,
    };
  }

  // merge any missing default sets
  const existingIds = new Set(Object.keys(base.personaSets || {}));
  for (const def of DEFAULT_SETS) {
    if (!existingIds.has(def.id)) {
      base.personaSets[def.id] = clone(def);
    }
  }

  // ensure activeSetId is valid
  if (!base.personaSets[base.activeSetId]) {
    base.activeSetId = DEFAULT_SETS[0].id;
  }

  return base;
}
function saveSettings() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

/* ---------- UI renderers (unchanged except legend uses colors) ---------- */
function switchView(view:"review"|"settings"){
  document.getElementById("view-review")?.classList.toggle("hidden", view!=="review");
  document.getElementById("view-settings")?.classList.toggle("hidden", view!=="settings");
}
function renderLegend(){
  const set = settings.personaSets[settings.activeSetId];
  const el = document.getElementById("legend")!;
  el.innerHTML = set.personas.filter(p=>p.enabled).map(p =>
    `<span class="swatch"><span class="dot" style="background:${p.color||"#e5e7eb"}"></span>${esc(p.name)}</span>`
  ).join("");
}
function renderPersonaSetSelectors(){
  const sel = document.getElementById("personaSet") as HTMLSelectElement;
  const list = document.getElementById("personaList") as HTMLDivElement;
  const sets = Object.values(settings.personaSets);
  sel.innerHTML = sets.map(s=>`<option value="${s.id}">${esc(s.name)}</option>`).join("");
  sel.value = settings.activeSetId;
  const names = settings.personaSets[settings.activeSetId].personas.filter(p=>p.enabled).map(p=>p.name);
  list.textContent = names.join(" • ");
  renderLegend();
}
function personaRow(p: Persona){
  return `
  <div class="section">
    <label><input type="checkbox" class="pe-enabled" data-id="${p.id}" ${p.enabled?"checked":""}/> Enabled — <strong>${esc(p.name)}</strong></label>
    <div class="row"><label>Name</label><input type="text" class="pe-name" data-id="${p.id}" value="${esc(p.name)}"/></div>
    <div class="row"><label>Color</label><input type="color" class="pe-color" data-id="${p.id}" value="${p.color||"#e5e7eb"}"/></div>
    <div class="row"><label>System Prompt</label><input type="text" class="pe-system" data-id="${p.id}" value="${esc(p.system)}"/></div>
    <div class="row"><label>Instruction Prompt</label><input type="text" class="pe-instruction" data-id="${p.id}" value="${esc(p.instruction)}"/></div>
  </div>`;
}
function renderPersonaEditor(){
  const setId = (document.getElementById("settingsPersonaSet") as HTMLSelectElement).value;
  const set = settings.personaSets[setId];
  const ed = document.getElementById("personaEditor")!;
  ed.innerHTML = set.personas.map(personaRow).join("");

  ed.querySelectorAll<HTMLInputElement>(".pe-enabled").forEach(inp=>{
    inp.onchange = ()=>{ set.personas.find(x=>x.id===inp.dataset.id)!.enabled = inp.checked; saveSettings(); renderPersonaSetSelectors(); };
  });
  ed.querySelectorAll<HTMLInputElement>(".pe-name").forEach(inp=>{
    inp.oninput = ()=>{ set.personas.find(x=>x.id===inp.dataset.id)!.name = inp.value; saveSettings(); renderPersonaSetSelectors(); };
  });
  ed.querySelectorAll<HTMLInputElement>(".pe-color").forEach(inp=>{
    inp.oninput = ()=>{ set.personas.find(x=>x.id===inp.dataset.id)!.color = inp.value; saveSettings(); renderLegend(); };
  });
  ed.querySelectorAll<HTMLInputElement>(".pe-system").forEach(inp=>{
    inp.oninput = ()=>{ set.personas.find(x=>x.id===inp.dataset.id)!.system = inp.value; saveSettings(); };
  });
  ed.querySelectorAll<HTMLInputElement>(".pe-instruction").forEach(inp=>{
    inp.oninput = ()=>{ set.personas.find(x=>x.id===inp.dataset.id)!.instruction = inp.value; saveSettings(); };
  });
}
function renderSettingsForm(){
  const provider = document.getElementById("provider") as HTMLSelectElement;
  const keyRow = document.getElementById("openrouterKeyRow")!;
  const key = document.getElementById("openrouterKey") as HTMLInputElement;
  const model = document.getElementById("model") as HTMLInputElement;
  const setSel = document.getElementById("settingsPersonaSet") as HTMLSelectElement;

  provider.value = settings.provider;
  key.value = settings.openrouterKey||"";
  model.value = settings.model;
  keyRow.style.display = settings.provider==="openrouter" ? "block" : "none";

  const sets = Object.values(settings.personaSets);
  setSel.innerHTML = sets.map(s=>`<option value="${s.id}">${esc(s.name)}</option>`).join("");
  setSel.value = settings.activeSetId;

  renderPersonaEditor();
}
function scoreBar(v?:number){
  const n = Math.max(0, Math.min(100, Number(v||0)));
  return `<div class="scorebar"><div class="scorebar-fill" style="width:${n}%"></div></div><div class="muted" style="font-size:12px">${n}/100</div>`;
}
function renderResultsView(results: Record<string, any>){
  const set = settings.personaSets[settings.activeSetId];
  const el = document.getElementById("results")!;
  el.innerHTML = set.personas.filter(p=>p.enabled).map(p=>{
    const r = results[p.id]; const s=r?.scores||{};
    const gf=(r?.global_feedback||"").toString().replace(/\n/g,"<br/>");
    return `<div class="result-card">
      <div style="display:flex;align-items:center;gap:8px;">
        <span class="dot" style="background:${p.color||"#e5e7eb"}"></span><strong>${esc(p.name)}</strong>
      </div>
      <div style="display:grid;grid-template-columns:110px 1fr;gap:8px;margin-top:6px">
        <div class="muted">Clarity</div><div>${scoreBar(s.clarity)}</div>
        <div class="muted">Tone</div><div>${scoreBar(s.tone)}</div>
        <div class="muted">Alignment</div><div>${scoreBar(s.alignment)}</div>
      </div>
      <div style="margin-top:8px">${gf}</div>
    </div>`;
  }).join("");
}
function renderStatuses(status: Record<string, PersonaRunState>){
  const set = settings.personaSets[settings.activeSetId];
  const personaStatus = document.getElementById("personaStatus")!;
  const progBar = document.getElementById("progBar") as HTMLDivElement;
  personaStatus.innerHTML = set.personas.filter(p=>p.enabled).map(p=>{
    const st = status[p.id]||"queued";
    const bg = st==="running"?"#fff7ed":st==="done"?"#ecfdf5":st==="failed"?"#fef2f2":"#eef2ff";
    const fg = st==="running"?"#9a3412":st==="done"?"#065f46":st==="failed"?"#991b1b":"#3730a3";
    return `<div class="row" style="display:flex;align-items:center;gap:8px;">
      <span class="dot" style="background:${p.color||"#e5e7eb"}"></span>
      <span style="background:${bg};color:${fg};padding:2px 6px;border-radius:10px;font-size:12px">${esc(p.name)}: ${st}</span>
    </div>`;
  }).join("");
  const total = set.personas.filter(p=>p.enabled).length;
  const done = Object.values(status).filter(s=>s==="done").length;
  progBar.style.width = total ? `${Math.round((done/total)*100)}%` : "0%";
}

/* ---------- LLM I/O (unchanged from your working version) ---------- */
function safeParseJSON(s:string){ try{ const t=s.trim().replace(/^```(json)?/i,"").replace(/```$/,"").trim(); return JSON.parse(t);}catch(e:any){ return {_parse_error:String(e),_raw:s};}}
async function callLLM(persona: Persona, docText: string): Promise<any> {
  const meta = `Return STRICT JSON: {"scores":{"clarity":number,"tone":number,"alignment":number},"global_feedback":string,"comments":[{"quote":string,"comment":string}]}`;
  const messages = [
    { role: "system", content: persona.system },
    { role: "user", content: `${persona.instruction}\n\n---\nDOCUMENT:\n${docText}` },
    { role: "user", content: meta },
  ];
  if ((settings.model||"").toLowerCase()==="debug-stub"){
    return { scores:{clarity:82,tone:76,alignment:88}, global_feedback:`Stub feedback for ${persona.name}.`, comments:[{quote:docText.slice(0,60),comment:"Example inline comment from stub."}] };
  }
  if (settings.provider==="openrouter"){
    const resp = await fetch("https://openrouter.ai/api/v1/chat/completions",{
      method:"POST",
      headers:{
        "Content-Type":"application/json",
        Authorization:`Bearer ${settings.openrouterKey}`,
        "HTTP-Referer": location.origin,
        "X-Title": "Persona Feedback Word Add-in",
      },
      body: JSON.stringify({ model: settings.model, messages, temperature: 0 })
    });
    const raw = await resp.text(); if(!resp.ok) throw new Error(`OpenRouter ${resp.status}: ${raw.slice(0,300)}`);
    const data = JSON.parse(raw); const txt = data?.choices?.[0]?.message?.content ?? "";
    return safeParseJSON(txt);
  } else {
    const resp = await fetch("http://localhost:11434/api/chat",{
      method:"POST", headers:{"Content-Type":"application/json"},
      body: JSON.stringify({ model: settings.model, messages })
    });
    const raw = await resp.text(); if(!resp.ok) throw new Error(`Ollama ${resp.status}: ${raw.slice(0,300)}`);
    let json:any; try{json=JSON.parse(raw);}catch{json={_raw:raw};}
    const txt = json?.message?.content ?? json?.choices?.[0]?.message?.content ?? raw;
    return safeParseJSON(txt);
  }
}

/* ---------- Word comment insertion (improved color/emoji tag) ---------- */
async function insertComments(persona: Persona, comments: {quote:string; comment:string}[]){
  if (!comments?.length) return;
  const hl = hexToHighlightColor(persona.color);
  const tag = `${colorToEmoji(persona.color)} ${persona.name}: `;

  await Word.run(async (ctx)=>{
    const body = ctx.document.body; body.load("text"); await ctx.sync();

    for (const c of comments){
      const text = (c.comment||"").trim(); if (!text) continue;
      // Try to anchor to quote
      let target: Word.Range | null = null;
      if (c.quote && c.quote.trim()){
        const search = body.search(c.quote.trim(), { matchCase:false, matchWholeWord:false });
        search.load("items"); await ctx.sync();
        if (search.items.length) target = search.items[0];
      }
      if (target){
        (target as any).font.highlightColor = hl;
        (target as any).insertComment(tag + text);
      } else {
        const end = body.getRange("End");
        (end as any).insertComment(tag + text);
        try { (end as any).font.highlightColor = hl; } catch {}
      }
      await ctx.sync();
    }
  });
}

/* ---------- run flow & export (same behavior) ---------- */
let lastRunResults: Record<string, any> = {};
async function getDocText(){ let t=""; await Word.run(async ctx=>{const b=ctx.document.body; b.load("text"); await ctx.sync(); t=b.text||"";}); return t;}
async function runReview(){
  const runBtn = document.getElementById("runBtn") as HTMLButtonElement | null;
  try{
    if (runBtn) runBtn.disabled = true;
    const doc = await getDocText(); if (!doc) { showToast("Document appears empty."); return; }
    const set = settings.personaSets[settings.activeSetId];
    const personas = set.personas.filter(p=>p.enabled);
    if (!personas.length){ showToast("No personas enabled."); return; }

    const status: Record<string, PersonaRunState> = {};
    personas.forEach(p=>status[p.id]="queued"); renderStatuses(status);

    const results: Record<string, any> = {};
    for (const p of personas){
      try{
        status[p.id]="running"; renderStatuses(status);
        const json = await callLLM(p, doc);
        results[p.id]=json;
        if (json?.comments?.length) await insertComments(p, json.comments);
        status[p.id] = json?._parse_error ? "failed" : "done";
        renderStatuses(status); renderResultsView(results);
      }catch(e:any){
        status[p.id]="failed"; renderStatuses(status); debug("persona failed",p.id,String(e));
      }
    }
    lastRunResults = results; showToast("Review complete.");
  } finally { if (runBtn) runBtn.disabled=false; }
}

// export report
function htmlEscape(s:string){ return s.replace(/[&<>"]/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]!));}
function buildReportHTML(res:Record<string,any>){
  const set=settings.personaSets[settings.activeSetId];
  const parts = set.personas.filter(p=>p.enabled).map(p=>{
    const r=res[p.id]||{}, s=r.scores||{}, gf=htmlEscape((r.global_feedback||"").toString()).replace(/\n/g,"<br/>");
    const comments:(Array<{quote:string;comment:string}>)=(r.comments||[]);
    const list = comments.length ? `<ul>${comments.map(c=>`<li><em>${htmlEscape(c.quote||"")}</em><br/>${htmlEscape(c.comment||"")}</li>`).join("")}</ul>` : `<div class="muted">No inline comments</div>`;
    return `<section class="card">
      <h2><span class="dot" style="background:${p.color||"#e5e7eb"}"></span> ${esc(p.name)}</h2>
      <div class="grid">
        <div>Clarity</div><div><div class="bar"><div style="width:${Number(s.clarity||0)}%"></div></div><small>${Number(s.clarity||0)}/100</small></div>
        <div>Tone</div><div><div class="bar"><div style="width:${Number(s.tone||0)}%"></div></div><small>${Number(s.tone||0)}/100</small></div>
        <div>Alignment</div><div><div class="bar"><div style="width:${Number(s.alignment||0)}%"></div></div><small>${Number(s.alignment||0)}/100</small></div>
      </div>
      <h3>Global Feedback</h3><p>${gf}</p>
      <h3>Inline Comments</h3>${list}
    </section>`;
  }).join("");
  return `<!doctype html><html><head><meta charset="utf-8"/><title>Persona Feedback Report</title>
  <style>
    body{font-family:Segoe UI,Roboto,Arial,sans-serif;max-width:900px;margin:24px auto;padding:0 16px;color:#111827}
    h1{font-size:22px;margin:0 0 12px} h2{font-size:16px;margin:0 0 8px;display:flex;align-items:center;gap:8px}
    .muted{color:#6b7280} .dot{width:10px;height:10px;border-radius:50%;border:1px solid #d1d5db;display:inline-block}
    .card{border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin-bottom:12px}
    .grid{display:grid;grid-template-columns:120px 1fr;gap:8px;margin:8px 0}
    .bar{height:8px;background:#e5e7eb;border-radius:4px;overflow:hidden}.bar>div{height:100%;background:#2563eb}
  </style></head><body><h1>Persona Feedback Report</h1>${parts||"<div class='muted'>No results</div>"}</body></html>`;
}
function exportReport(){
  const html=buildReportHTML(lastRunResults);
  const blob=new Blob([html],{type:"text/html;charset=utf-8"});
  const url=URL.createObjectURL(blob); const a=document.createElement("a");
  a.href=url; a.download=`persona-feedback-${new Date().toISOString().replace(/[:.]/g,"-")}.html`;
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}

/* ---------- wiring & boot ---------- */
function wire(){
  const gear=document.getElementById("gear"); const back=document.getElementById("backToReview");
  const runBtn=document.getElementById("runBtn"); const personaSet=document.getElementById("personaSet") as HTMLSelectElement;
  const toggleDebug=document.getElementById("toggleDebug"); const clearDebug=document.getElementById("clearDebug");
  const exportBtn=document.getElementById("exportBtn");

  gear?.addEventListener("click",()=>{ renderSettingsForm(); switchView("settings"); });
  back?.addEventListener("click",()=>switchView("review"));
  runBtn?.addEventListener("click",runReview);
  exportBtn?.addEventListener("click",exportReport);
  personaSet?.addEventListener("change",()=>{ settings.activeSetId=personaSet.value; saveSettings(); renderPersonaSetSelectors(); });

  toggleDebug?.addEventListener("click",()=>{
    const dp=document.getElementById("debugPanel")!; const btn=toggleDebug as HTMLButtonElement;
    const v=dp.classList.toggle("hidden"); btn.textContent = v ? "Show Debug" : "Hide Debug";
  });
  clearDebug?.addEventListener("click",()=>{ const dbg=document.getElementById("debugLog"); if (dbg) dbg.innerHTML=""; });

  const provider=document.getElementById("provider") as HTMLSelectElement;
  const keyRow=document.getElementById("openrouterKeyRow")!;
  const key=document.getElementById("openrouterKey") as HTMLInputElement;
  const model=document.getElementById("model") as HTMLInputElement;
  const setSel=document.getElementById("settingsPersonaSet") as HTMLSelectElement;
  const saveBtn=document.getElementById("saveSettings"); const resetBtn=document.getElementById("restoreDefaults");

  provider?.addEventListener("change",()=>{ settings.provider=provider.value as any; keyRow.style.display = settings.provider==="openrouter"?"block":"none"; saveSettings(); });
  key?.addEventListener("input",()=>{ settings.openrouterKey=key.value; saveSettings(); });
  model?.addEventListener("input",()=>{ settings.model=model.value; saveSettings(); });
  setSel?.addEventListener("change",()=>{ settings.activeSetId=setSel.value; saveSettings(); renderPersonaEditor(); renderPersonaSetSelectors(); });
  saveBtn?.addEventListener("click",()=>{ saveSettings(); showToast("Settings saved"); });
  resetBtn?.addEventListener("click",()=>{
    const def = DEFAULT_SETS.find(s=>s.id===settings.activeSetId)!; settings.personaSets[def.id]=clone(def);
    saveSettings(); renderPersonaEditor(); renderPersonaSetSelectors(); showToast("Restored defaults");
  });
}

(function boot(){
  if (typeof (window as any).Office === "undefined"){
    const warn=document.createElement("div"); warn.style.background="#fff7ed"; warn.style.color="#9a3412";
    warn.style.padding="8px"; warn.style.border="1px solid #fed7aa"; warn.style.borderRadius="8px"; warn.style.marginTop="8px";
    warn.textContent="Open this add-in from Word (Home → Persona Feedback)."; document.body.prepend(warn); return;
  }
  (window as any).Office.onReady().then(()=>{
    settings = loadSettings();
    wire();
    renderPersonaSetSelectors();
    renderResultsView({});
    switchView("review");
    debug("ready; active set:", settings.activeSetId);
  });
})();
