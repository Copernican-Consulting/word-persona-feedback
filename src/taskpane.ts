/* eslint-disable @typescript-eslint/no-explicit-any */

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

// ---------- Defaults ----------
function P(name:string,system:string,instruction:string,color:string):Persona{
  return { id:name.toLowerCase().replace(/[^a-z0-9]+/g,"-"), enabled:true, name, system, instruction, color };
}
const DEFAULT_SETS: PersonaSet[] = [
  { id:"cross-functional-team", name:"Cross-Functional Team", personas:[
    P("Senior Manager","Senior manager prioritizing clarity, risk, outcomes.","Assess clarity of goals, risks, and outcomes.","#fde047"),
    P("Legal","Corporate counsel focused on compliance.","Flag risky claims, suggest safer wording.","#f9a8d4"),
    P("HR","HR partner.","Spot exclusionary tone; suggest inclusive language.","#5eead4"),
    P("Technical Lead","Engineering lead.","Check feasibility, gaps, risks.","#93c5fd"),
    P("Junior Analyst","Analyst.","Call out unclear logic, missing data.","#86efac"),
  ]},
  { id:"marketing-focus-group", name:"Marketing Focus Group", personas:[
    P("Midwest Parent","Pragmatic parent.","React to clarity, trust, family benefit.","#f59e0b"),
    P("Gen-Z Student","Digital native.","React to tone/authenticity.","#0ea5e9"),
    P("Retired Veteran","Values respect/responsibility.","React to credibility/plain language.","#6d28d9"),
    P("Small Business Owner","ROI-driven.","React to value proposition.","#16a34a"),
    P("Tech Pro","Precise, detail-oriented.","React to vague claims.","#ef4444"),
  ]},
];

function defaultSettings():AppSettings{
  return { provider:{provider:"openrouter",model:"openrouter/auto",openrouterKey:""}, personaSetId:DEFAULT_SETS[0].id, personaSets:DEFAULT_SETS };
}
function loadSettings():AppSettings{ try{ const s=localStorage.getItem(LS_KEY); if(!s) return defaultSettings(); const v=JSON.parse(s) as AppSettings; if(!v.personaSets?.length) v.personaSets=DEFAULT_SETS; if(!v.personaSetId) v.personaSetId=v.personaSets[0].id; return v;}catch{return defaultSettings();} }
function saveSettings(){ localStorage.setItem(LS_KEY, JSON.stringify(SETTINGS)); }
function currentSet(){ const id=SETTINGS.personaSetId; return SETTINGS.personaSets.find(s=>s.id===id)||SETTINGS.personaSets[0]; }

// ---------- Office bootstrap ----------
Office.onReady(async ()=>{
  SETTINGS=loadSettings(); 
  document.body.style.minWidth="500px";  // 👈 Default width

  log("[PF] Office.onReady → UI initialized");
  // … hook up UI events etc (same as before) …
});

// ---------- Results bookkeeping ----------
function upsert<T extends { personaId:string }>(arr:T[], item:T){
  const i=arr.findIndex(x=>x.personaId===item.personaId);
  if(i>=0) arr[i]=item; else arr.push(item);
}
function upsertResult(r: PersonaRunResult){
  upsert(LAST_RESULTS, r);
}
