
import "./ui.css";

// Minimal Office globals for TS
declare const Word: any;

type Provider = "openrouter" | "ollama";
interface ModelConfig { provider: Provider; apiKey?: string; model: string; }
export interface Persona { id: string; enabled: boolean; name: string; system: string; instruction: string; }
export interface PersonaSet { id: string; name: string; personas: Persona[]; }
type PersonaSets = Record<string, PersonaSet>;

// DOM helpers
const EL = <T extends HTMLElement>(sel: string) => document.querySelector(sel) as T;
const qsa = <T extends HTMLElement>(sel: string, root: HTMLElement|Document = document) => Array.from(root.querySelectorAll(sel)) as T[];

function mkPersona(id:string,name:string,system:string,instruction:string):Persona{
  return {id,name,system,instruction,enabled:true};
}

// ---- Defaults (you can add more sets later) ----
const DEFAULT_SETS: PersonaSets = {
  crossFunctional:{ id:"crossFunctional", name:"Cross-Functional Team", personas:[
    mkPersona("seniorManager","Senior Manager","Be concise, outcome-focused.","Score clarity, tone, goal alignment; leave inline comments."),
    mkPersona("legal","Legal","Risk-aware, compliance-first.","Identify legal risks, ambiguous claims, missing disclaimers."),
    mkPersona("hr","HR","Inclusive, respectful, policy-aligned.","Flag phrasing issues; suggest inclusive alternatives."),
    mkPersona("techLead","Technical Lead","Deep technical rigor.","Point out inaccuracies; suggest precise fixes."),
    mkPersona("juniorAnalyst","Junior Analyst","Curious, literal reader.","Flag confusing parts; ask clarifying questions.")
  ]}
};

// ---- Local storage keys ----
const KEY_MODEL="pf:model";
const KEY_SETS="pf:personas";
const KEY_ACTIVE_SET="pf:activeSet";

function loadModel():ModelConfig{
  try{const raw=localStorage.getItem(KEY_MODEL); if(raw) return JSON.parse(raw);}catch{}
  return {provider:"openrouter",model:"openrouter/auto",apiKey:""};
}
function saveModel(m:ModelConfig){ localStorage.setItem(KEY_MODEL, JSON.stringify(m)); }

function loadSets():PersonaSets{
  try{const raw=localStorage.getItem(KEY_SETS); if(raw) return JSON.parse(raw);}catch{}
  localStorage.setItem(KEY_SETS, JSON.stringify(DEFAULT_SETS));
  return structuredClone(DEFAULT_SETS);
}
function saveSets(s:PersonaSets){ localStorage.setItem(KEY_SETS, JSON.stringify(s)); }

function loadActiveSetId(sets:PersonaSets){
  const s=localStorage.getItem(KEY_ACTIVE_SET);
  if(s && (sets as any)[s]) return s;
  const first=Object.keys(sets)[0];
  localStorage.setItem(KEY_ACTIVE_SET, first);
  return first;
}
function setActiveSetId(id:string){ localStorage.setItem(KEY_ACTIVE_SET, id); }

let model=loadModel();
let personaSets=loadSets();
let activeSetId=loadActiveSetId(personaSets);

// ---- UI elements ----
const btnSettings=EL<HTMLButtonElement>("#btnSettings");
const viewReview=EL<HTMLElement>("#view-review");
const viewSettings=EL<HTMLElement>("#view-settings");
const tabReview=EL<HTMLButtonElement>("#tabReview");
const tabSettings=EL<HTMLButtonElement>("#tabSettings");

const reviewPersonaSetSelect=EL<HTMLSelectElement>("#reviewPersonaSetSelect");
const reviewPersonaList=EL<HTMLDivElement>("#reviewPersonaList");
const btnRunReview=EL<HTMLButtonElement>("#btnRunReview");
const scoresContainer=EL<HTMLDivElement>("#scoresContainer");
const globalFeedbackContainer=EL<HTMLDivElement>("#globalFeedbackContainer");

const providerSelect=EL<HTMLSelectElement>("#providerSelect");
const openrouterKey=EL<HTMLInputElement>("#openrouterKey");
const modelInput=EL<HTMLInputElement>("#modelInput");
const openrouterKeyRow=EL<HTMLDivElement>("#openrouterKeyRow");

const settingsPersonaSetSelect=EL<HTMLSelectElement>("#settingsPersonaSetSelect");
const personaEditorContainer=EL<HTMLDivElement>("#personaEditorContainer");
const btnSaveSettings=EL<HTMLButtonElement>("#btnSaveSettings");
const btnRestoreDefaults=EL<HTMLButtonElement>("#btnRestoreDefaults");
const btnBackToReview=EL<HTMLButtonElement>("#btnBackToReview");

// ---- Navigation ----
function goReview(){ tabReview.classList.add("pf-tab--active"); tabSettings.classList.remove("pf-tab--active"); viewSettings.classList.remove("active"); viewReview.classList.add("active"); renderReview(); }
function goSettings(){ tabSettings.classList.add("pf-tab--active"); tabReview.classList.remove("pf-tab--active"); viewReview.classList.remove("active"); viewSettings.classList.add("active"); renderSettings(); }
btnSettings.onclick=goSettings; tabSettings.onclick=goSettings; tabReview.onclick=goReview; btnBackToReview.onclick=goReview;

// ---- Renderers ----
function renderReview(){
  fillPersonaSetSelect(reviewPersonaSetSelect, personaSets, activeSetId, false);
  const set=personaSets[activeSetId];
  reviewPersonaList.innerHTML="";
  set.personas.filter(p=>p.enabled).forEach(p=>{
    const d=document.createElement("div");
    d.className="pf-persona-pill";
    d.innerHTML=`<span class="name">${escapeHtml(p.name)}</span><span class="hint">(read-only)</span>`;
    reviewPersonaList.appendChild(d);
  });
  scoresContainer.innerHTML="";
  globalFeedbackContainer.innerHTML="";
}

function renderSettings(){
  providerSelect.value=model.provider;
  openrouterKey.value=model.apiKey??"";
  modelInput.value=model.model;
  openrouterKeyRow.style.display=(providerSelect.value==="openrouter")?"block":"none";

  fillPersonaSetSelect(settingsPersonaSetSelect, personaSets, activeSetId, true);
  renderPersonaEditors();
}

function fillPersonaSetSelect(sel:HTMLSelectElement, sets:PersonaSets, active:string, includeCreate:boolean){
  sel.innerHTML="";
  Object.values(sets).forEach(s=>{
    const opt=document.createElement("option"); opt.value=s.id; opt.textContent=s.name;
    if(s.id===active) opt.selected=true;
    sel.appendChild(opt);
  });
  if(includeCreate){
    const opt=document.createElement("option"); opt.value="__create__"; opt.textContent="[Create New Persona Set]";
    sel.appendChild(opt);
  }
}

function renderPersonaEditors(){
  personaEditorContainer.innerHTML="";
  const set=personaSets[activeSetId];
  const wrap=document.createElement("div");
  wrap.className="pf-editor";

  set.personas.forEach((p,idx)=>{
    const block=document.createElement("div");
    block.className="pf-card";
    block.innerHTML = `
      <h4>Persona ${idx+1}</h4>
      <div class="pf-field"><label><input type="checkbox" data-k="enabled" ${p.enabled?"checked":""}/> Enabled</label></div>
      <div class="pf-field"><label>Name</label><input type="text" data-k="name" value="${escapeAttr(p.name)}"/></div>
      <div class="pf-field"><label>System Prompt</label><textarea data-k="system">${escapeText(p.system)}</textarea></div>
      <div class="pf-field"><label>Instruction Prompt</label><textarea data-k="instruction">${escapeText(p.instruction)}</textarea></div>
    `;
    // wire inputs
    qsa<HTMLInputElement|HTMLTextAreaElement>('input,textarea', block).forEach(inp=>{
      inp.addEventListener('input',()=>{
        const key=(inp.getAttribute('data-k')||'') as keyof Persona;
        // @ts-ignore
        p[key] = (key==="enabled") ? (inp as HTMLInputElement).checked : (inp as HTMLInputElement).value;
      });
      inp.addEventListener('change',()=>{
        const key=(inp.getAttribute('data-k')||'') as keyof Persona;
        // @ts-ignore
        p[key] = (key==="enabled") ? (inp as HTMLInputElement).checked : (inp as HTMLInputElement).value;
      });
    });
    personaEditorContainer.appendChild(block);
  });

  personaEditorContainer.appendChild(wrap);
}

// ---- Events ----
reviewPersonaSetSelect.onchange=(e)=>{
  const id=(e.target as HTMLSelectElement).value;
  if(personaSets[id]){ activeSetId=id; setActiveSetId(id); renderReview(); }
};

settingsPersonaSetSelect.onchange=(e)=>{
  const v=(e.target as HTMLSelectElement).value;
  if(v==="__create__"){
    const name=prompt("Name for new persona set?");
    if(!name){ (e.target as HTMLSelectElement).value=activeSetId; return; }
    const newId=slugify(name);
    personaSets[newId]={ id:newId, name, personas:[ mkPersona("p1","Persona 1","",""), mkPersona("p2","Persona 2","",""), mkPersona("p3","Persona 3","","") ] };
    saveSets(personaSets);
    activeSetId=newId; setActiveSetId(newId);
  }else{
    activeSetId=v; setActiveSetId(v);
  }
  renderSettings();
};

providerSelect.onchange=()=>{
  model.provider=providerSelect.value as Provider;
  openrouterKeyRow.style.display=(model.provider==="openrouter")?"block":"none";
};
openrouterKey.oninput=()=>{ model.apiKey=openrouterKey.value.trim(); };
modelInput.oninput=()=>{ model.model=modelInput.value.trim(); };

btnSaveSettings.onclick=()=>{ saveModel(model); saveSets(personaSets); alert("Settings saved."); };
btnRestoreDefaults.onclick=()=>{
  const def=(DEFAULT_SETS as any)[activeSetId];
  if(!def){ alert("No default template for this set."); return; }
  personaSets[activeSetId]=structuredClone(def);
  saveSets(personaSets);
  renderPersonaEditors();
  alert("Restored defaults for current set.");
};

// ---- Review flow ----
type ReviewJSON={
  scores:{clarity:number;tone:number;alignment:number;[k:string]:number};
  global_feedback:string;
  comments:Array<{quote:string;comment:string}>;
};

btnRunReview.onclick=async()=>{
  scoresContainer.innerHTML=""; globalFeedbackContainer.innerHTML="";
  const set=personaSets[activeSetId];
  const enabled=set.personas.filter(p=>p.enabled);
  if(!enabled.length){ alert("No enabled personas."); return; }

  const docText=await getDocumentText();
  for(const p of enabled){
    let result:ReviewJSON;
    try{
      result=await runPersonaReview(p, docText, model);
    }catch(e:any){
      result={ scores:{clarity:0,tone:0,alignment:0}, global_feedback:`Error: ${e?.message||e}`, comments:[] };
    }
    // render scores + global
    const sc = document.createElement("div");
    sc.className="pf-score";
    sc.innerHTML = `<strong>${escapeHtml(p.name)}</strong><br/>
      Clarity: ${fmtScore(result.scores.clarity)} / 100<br/>
      Tone: ${fmtScore(result.scores.tone)} / 100<br/>
      Alignment: ${fmtScore(result.scores.alignment)} / 100`;
    scoresContainer.appendChild(sc);

    const gf = document.createElement("div");
    gf.innerHTML = `<h4>${escapeHtml(p.name)} – Global Feedback</h4><p>${escapeHtml(result.global_feedback||"")}</p>`;
    globalFeedbackContainer.appendChild(gf);

    // add Word comments
    await insertCommentsIntoWord(`${p.name} (AI)`, result.comments||[]);
  }
};

function fmtScore(n:number){ return Number.isFinite(n)? Math.round(n):0; }

const META_PROMPT=(persona:Persona,text:string)=>[
  "You are the following persona reviewing a Word document.",
  "",
  `Persona Name: ${persona.name}`,
  `System Persona: ${persona.system}`,
  "",
  "INSTRUCTIONS FOR PERSONA:",
  persona.instruction,
  "",
  "Return a STRICT JSON object with the following shape (and nothing else):",
  "{",
  "  \"scores\": { \"clarity\": 0-100, \"tone\": 0-100, \"alignment\": 0-100 },",
  "  \"global_feedback\": \"short paragraph of overall feedback\",",
  "  \"comments\": [",
  "     { \"quote\": \"exact span from the doc\", \"comment\": \"your brief comment\" }",
  "  ]",
  "}",
  "",
  "Document text:",
  "--------------------",
  text.slice(0, 15000) // bound prompt size a bit
].join("\n");

async function runPersonaReview(persona:Persona, text:string, cfg:ModelConfig):Promise<ReviewJSON>{
  const prompt=META_PROMPT(persona,text);
  let raw:string;
  if(cfg.provider==="openrouter"){
    if(!cfg.apiKey) throw new Error("OpenRouter API key missing. Add it in Settings.");
    raw=await callOpenRouter(cfg.model||"openrouter/auto", cfg.apiKey!, prompt, persona.system);
  }else{
    raw=await callOllama(cfg.model||"llama3.1:8b", prompt, persona.system);
  }
  const json=safeExtractJSON(raw);
  const scores=json?.scores||{};
  return {
    scores:{ clarity:num0(scores.clarity), tone:num0(scores.tone), alignment:num0(scores.alignment) },
    global_feedback: String(json?.global_feedback||""),
    comments: Array.isArray(json?.comments)? json.comments.map((c:any)=>({quote:String(c.quote||""),comment:String(c.comment||"")})): []
  };
}

function num0(x:any){ const n=Number(x); return Number.isFinite(n)? n:0; }

async function callOpenRouter(model:string, apiKey:string, prompt:string, system:string):Promise<string>{
  const res=await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method:"POST",
    headers:{
      "Content-Type":"application/json",
      "Authorization":`Bearer ${apiKey}`,
      "HTTP-Referer":"https://persona-feedback.local",
      "X-Title":"Persona Feedback Dev"
    },
    body: JSON.stringify({
      model,
      temperature:0.2,
      messages:[
        {role:"system",content:system||"You are a precise reviewer."},
        {role:"user",content:prompt}
      ]
    })
  });
  if(!res.ok) throw new Error(`OpenRouter ${res.status}: ${await res.text()}`);
  const data=await res.json();
  const content = data?.choices?.[0]?.message?.content ?? "";
  return String(content);
}

async function callOllama(model:string, prompt:string, system:string):Promise<string>{
  const res=await fetch("http://127.0.0.1:11434/api/chat", {
    method:"POST",
    headers:{ "Content-Type":"application/json" },
    body: JSON.stringify({
      model,
      options:{temperature:0.2},
      messages:[
        {role:"system",content:system||"You are a precise reviewer."},
        {role:"user",content:prompt}
      ],
      stream:false
    })
  });
  if(!res.ok) throw new Error(`Ollama ${res.status}: ${await res.text()}`);
  const data=await res.json();
  const content = data?.message?.content ?? "";
  return String(content);
}

function safeExtractJSON(s:string):any{
  // try raw
  try{ return JSON.parse(s);}catch{}
  // curly block
  const m=s.match(/\{[\s\S]*\}$/);
  if(m){ try{ return JSON.parse(m[0]); }catch{} }
  // fenced code
  const m2=s.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
  if(m2){ try{ return JSON.parse(m2[1]); }catch{} }
  return {};
}

async function getDocumentText():Promise<string>{
  // Uses Word.run to pull full doc text
  return Word.run(async (context:any)=>{
    const range=context.document.body.getRange();
    range.load("text");
    await context.sync();
    return String(range.text||"");
  });
}

async function insertCommentsIntoWord(authorLabel:string,items:Array<{quote:string;comment:string;}>){
  return Word.run(async (context:any)=>{
    const body=context.document.body;
    for(const it of items){
      if(!it.quote || !it.comment) continue;
      let results=body.search(it.quote, { matchCase:true, matchWholeWord:false, matchPrefix:false, ignoreSpace:false });
      results.load("items");
      await context.sync();
      if(results.items.length===0){
        results=body.search(it.quote, { matchCase:false });
        results.load("items");
        await context.sync();
      }
      if(results.items.length===0) continue;
      const range=results.items[0];
      const c = range.insertComment(`${authorLabel}: ${it.comment}`);
      c.author.initials = "AI";
      // Note: Word JS doesn't let us set author name directly; initials helps differentiate.
    }
    await context.sync();
  });
}

// utils
function escapeHtml(s:string){return s.replace(/[&<>]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;"} as any)[c]);}
function escapeAttr(s:string){return s.replace(/"/g,"&quot;");}
function escapeText(s:string){return s.replace(/</g,"&lt;");}
function slugify(s:string){return s.toLowerCase().replace(/[^a-z0-9]+/g,"-").replace(/(^-|-$)/g,"");}

// boot
function boot(){ goReview(); }
document.addEventListener("DOMContentLoaded", boot);
