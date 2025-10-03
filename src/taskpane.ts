/* eslint-disable @typescript-eslint/no-explicit-any */

// ------------------------------
// Types
// ------------------------------

type HighlightColor =
  | "yellow" | "pink" | "turquoise" | "red" | "blue" | "green" | "violet"
  | "darkYellow" | "darkPink" | "darkTurquoise" | "darkRed" | "darkBlue" | "darkGreen" | "darkViolet"
  | "none";

type Provider = "openrouter" | "ollama";

type Persona = {
  id: string;
  enabled: boolean;
  name: string;
  system: string;
  instruction: string;
  color: HighlightColor;
};

type PersonaSet = {
  id: string;
  name: string;
  personas: Persona[];
};

type ProviderSettings = {
  provider: Provider;
  openrouterKey?: string;
  model: string;
};

type AppSettings = {
  provider: ProviderSettings;
  personaSetId: string;
  personaSets: PersonaSet[];
};

type PersonaRunStatus = "idle" | "running" | "done" | "error";

type PersonaRunResult = {
  personaId: string;
  personaName: string;
  status: PersonaRunStatus;
  scores?: { clarity: number; tone: number; alignment: number };
  global_feedback?: string;
  comments?: {
    quote: string;
    spanStart: number;
    spanEnd: number;
    comment: string;
  }[];
  raw?: any;
  error?: string;
};

// ------------------------------
// Constants / Defaults
// ------------------------------

const WORD_COLORS: HighlightColor[] = [
  "yellow", "pink", "turquoise", "red", "blue", "green", "violet",
  "darkYellow", "darkPink", "darkTurquoise", "darkRed", "darkBlue", "darkGreen", "darkViolet",
];

function colorAt(i: number): HighlightColor {
  return WORD_COLORS[i % WORD_COLORS.length];
}

const META_PROMPT = `
You are a reviewer. Return ONLY valid JSON matching this schema:

{
  "scores": { "clarity": 0-100, "tone": 0-100, "alignment": 0-100 },
  "global_feedback": "short paragraph",
  "comments": [
    {
      "quote": "verbatim snippet from the doc (20-160 chars)",
      "spanStart": whole-document character start index (0-based),
      "spanEnd": whole-document character end index (exclusive),
      "comment": "persona-specific feedback about that snippet"
    }
  ]
}

RULES:
- Do not include markdown.
- Make 3–8 targeted "comments" entries with accurate spanStart/spanEnd covering the exact quote.
- Keep "global_feedback" ~2-5 sentences.
`;

function P(name: string, system: string, instruction: string, idx: number): Persona {
  return {
    id: name.toLowerCase().replace(/[^a-z0-9]+/g, "-"),
    enabled: true,
    name,
    system,
    instruction,
    color: colorAt(idx),
  };
}

// ---- Default Persona Sets ----
const DEFAULT_SETS: PersonaSet[] = [
  {
    id: "cross-functional-team",
    name: "Cross-Functional Team",
    personas: [
      P("Senior Manager", "You are a senior manager prioritizing clarity, risks, and outcomes.", "Assess clarity of goals, risks, and expected outcomes.", 0),
      P("Legal", "You are corporate counsel focused on compliance and risk.", "Flag ambiguity, risky claims, missing disclaimers.", 1),
      P("HR", "You are an HR business partner concerned with tone and inclusion.", "Identify exclusionary tone and suggest inclusive language.", 2),
      P("Technical Lead", "You are a pragmatic engineering lead.", "Check feasibility, gaps, and technical risks.", 3),
      P("Junior Analyst", "You are a detail-oriented junior analyst.", "Call out unclear logic, missing data, or inconsistent units.", 4),
    ],
  },
  {
    id: "marketing-focus-group",
    name: "Marketing Focus Group",
    personas: [
      P("Midwest Parent", "You are a pragmatic parent from the US Midwest.", "React to clarity, trustworthiness, and family benefit.", 0),
      P("Gen-Z Student", "You are a digital-native college student.", "React to tone, authenticity, and modern appeal.", 1),
      P("Retired Veteran", "You are a retired veteran valuing respect and responsibility.", "React to credibility and plain-language clarity.", 2),
      P("Small Business Owner", "You are a small business owner.", "React to practical value and ROI.", 3),
      P("Tech-savvy Pro", "You are a tech-savvy professional.", "React to precision and claims that need specifics.", 4),
    ],
  },
  {
    id: "startup-stakeholders",
    name: "Startup Stakeholders",
    personas: [
      P("Founder", "You are a startup founder.", "Push for clarity of vision, focus, and cadence.", 0),
      P("CTO", "You are a CTO.", "Probe technical feasibility, architecture, and risks.", 1),
      P("CMO", "You are a CMO.", "Probe messaging, audience, and differentiation.", 2),
      P("VC Investor", "You are a VC partner.", "Probe metrics, milestones, and risks.", 3),
      P("Customer", "You are a prospective customer.", "Probe concrete value, outcomes, and adoption risks.", 4),
    ],
  },
  {
    id: "political-spectrum",
    name: "Political Spectrum",
    personas: [
      P("Democratic Socialist", "You are a democratic socialist.", "Assess equity, public benefit, and ethical framing.", 0),
      P("Center Left", "You are center-left.", "Assess policy realism and social impact.", 1),
      P("Centrist/Independent", "You are centrist.", "Assess balance, tradeoffs, and fairness.", 2),
      P("Center Right", "You are center-right.", "Assess fiscal prudence and efficiency.", 3),
      P("MAGA", "You are a populist conservative.", "Assess national interest and plain-language clarity.", 4),
      P("Libertarian", "You are libertarian.", "Assess individual freedom and regulatory burden.", 5),
    ],
  },
  {
    id: "product-review-board",
    name: "Product Review Board",
    personas: [
      P("PM", "You are a product manager.", "Assess problem framing, success metrics, and scope.", 0),
      P("Design Lead", "You are a design lead.", "Assess user flows, accessibility, and tone.", 1),
      P("Data Scientist", "You are a data scientist.", "Assess measurability, data risks, and validity.", 2),
      P("Security", "You are a security lead.", "Assess data handling, privacy, and threat modeling.", 3),
      P("Support Lead", "You are a support lead.", "Assess failure modes and user communication.", 4),
    ],
  },
  {
    id: "scientific-peer-review",
    name: "Scientific Peer Review",
    personas: [
      P("Methods Reviewer", "You examine methods and reproducibility.", "Check experimental detail and validity.", 0),
      P("Stats Reviewer", "You examine statistical claims.", "Check sample size, tests, and uncertainty.", 1),
      P("Domain Expert", "You examine domain-specific accuracy.", "Check citations and assumptions.", 2),
      P("Ethics Reviewer", "You examine ethical and societal impact.", "Check risk mitigation and consent.", 3),
    ],
  },
  {
    id: "ux-research-panel",
    name: "UX Research Panel",
    personas: [
      P("New User", "First-time user perspective.", "Assess onboarding clarity and cognitive load.", 0),
      P("Power User", "Expert user perspective.", "Assess efficiency and discoverability.", 1),
      P("Accessibility Advocate", "Accessibility lens.", "Assess contrast, semantics, and alt text.", 2),
    ],
  },
  {
    id: "sales-deal-desk",
    name: "Sales Deal Desk",
    personas: [
      P("Sales Director", "Top-line growth focus.", "Assess messaging and objections.", 0),
      P("Solutions Architect", "Technical fit focus.", "Assess integrations, constraints, and risks.", 1),
      P("Legal (Customer)", "Customer counsel.", "Assess indemnities, data use, and SLAs.", 2),
      P("Procurement", "Buyer procurement.", "Assess pricing clarity and comparables.", 3),
    ],
  },
  {
    id: "board-of-directors",
    name: "Board of Directors",
    personas: [
      P("Chair", "Governance focus.", "Assess strategy coherence and oversight.", 0),
      P("Audit", "Audit committee.", "Assess controls, risks, and reporting.", 1),
      P("Compensation", "Comp committee.", "Assess incentives and fairness.", 2),
    ],
  },
  {
    id: "academic-committee",
    name: "Academic Committee",
    personas: [
      P("Dean", "Academic leadership.", "Assess alignment to mission and rigor.", 0),
      P("IRB Chair", "Research ethics.", "Assess consent, risk, and data handling.", 1),
      P("Funding Reviewer", "Grant committee.", "Assess merit, feasibility, and budget.", 2),
    ],
  },
];

// ------------------------------
// Globals
// ------------------------------
let SETTINGS: AppSettings;
let LAST_RESULTS: PersonaRunResult[] = [];

const el = (id: string) => document.getElementById(id)!;

function log(msg: string, data?: any) {
  if (data !== undefined) console.log(msg, data);
  else console.log(msg);
  const panel = el("debugLog");
  const line = document.createElement("div");
  line.style.whiteSpace = "pre-wrap";
  line.textContent = data ? `${msg} ${safeJson(data)}` : msg;
  panel.appendChild(line);
  panel.scrollTop = panel.scrollHeight;
}
function safeJson(x: any) { try { return JSON.stringify(x, null, 2); } catch { return String(x); } }

function toast(t: string) {
  const box = el("toast");
  el("toastMsg").textContent = t;
  box.style.display = "block";
  (el("toastClose") as HTMLSpanElement).onclick = () => (box.style.display = "none");
  setTimeout(() => (box.style.display = "none"), 3500);
}

function showView(id: "view-review" | "view-settings") {
  const review = el("view-review");
  const settings = el("view-settings");
  const btnBack = el("btnBack");
  if (id === "view-review") {
    review.classList.remove("hidden");
    settings.classList.add("hidden");
    btnBack.classList.add("hidden");
  } else {
    review.classList.add("hidden");
    settings.classList.remove("hidden");
    btnBack.classList.remove("hidden");
  }
}

function confirmAsync(title: string, message: string): Promise<boolean> {
  return new Promise((resolve) => {
    const overlay = el("confirmOverlay");
    el("confirmTitle").textContent = title;
    el("confirmMessage").textContent = message;
    overlay.style.display = "flex";
    const ok = el("confirmOk"); const cancel = el("confirmCancel");
    const done = (v: boolean) => { overlay.style.display = "none"; ok.removeEventListener("click", onOk); cancel.removeEventListener("click", onCancel); resolve(v); };
    const onOk = () => done(true); const onCancel = () => done(false);
    ok.addEventListener("click", onOk); cancel.addEventListener("click", onCancel);
  });
}

const LS_KEY = "pf.settings.v1";
function defaultSettings(): AppSettings {
  return {
    provider: { provider: "openrouter", model: "openrouter/auto", openrouterKey: "" },
    personaSetId: DEFAULT_SETS[0].id,
    personaSets: DEFAULT_SETS,
  };
}
function loadSettings(): AppSettings {
  try {
    const s = localStorage.getItem(LS_KEY);
    if (!s) return defaultSettings();
    const parsed = JSON.parse(s) as AppSettings;
    if (!parsed.personaSets?.length) parsed.personaSets = DEFAULT_SETS;
    if (!parsed.personaSetId) parsed.personaSetId = DEFAULT_SETS[0].id;
    return parsed;
  } catch { return defaultSettings(); }
}
function saveSettings() { localStorage.setItem(LS_KEY, JSON.stringify(SETTINGS)); }

function currentSet(): PersonaSet {
  const id = SETTINGS.personaSetId;
  return SETTINGS.personaSets.find((s) => s.id === id) || SETTINGS.personaSets[0];
}

function highlightToCss(h: HighlightColor): string {
  const map: Record<string, string> = {
    yellow: "#fde047", darkYellow: "#f59e0b",
    pink: "#f9a8d4", darkPink: "#db2777",
    turquoise: "#5eead4", darkTurquoise: "#0d9488",
    red: "#f87171", darkRed: "#b91c1c",
    blue: "#93c5fd", darkBlue: "#1d4ed8",
    green: "#86efac", darkGreen: "#166534",
    violet: "#c4b5fd", darkViolet: "#6d28d9",
  };
  return map[h] || "#fde047";
}

function populatePersonaSets() {
  const sets = SETTINGS.personaSets;

  const reviewSel = el("personaSet") as HTMLSelectElement;
  reviewSel.innerHTML = "";
  sets.forEach((s) => {
    const opt = document.createElement("option"); opt.value = s.id; opt.textContent = s.name; reviewSel.appendChild(opt);
  });
  reviewSel.value = SETTINGS.personaSetId;

  const settingsSel = el("settingsPersonaSet") as HTMLSelectElement;
  settingsSel.innerHTML = "";
  sets.forEach((s) => {
    const opt = document.createElement("option"); opt.value = s.id; opt.textContent = s.name; settingsSel.appendChild(opt);
  });
  settingsSel.value = SETTINGS.personaSetId;

  renderPersonaNamesAndLegend();
  renderPersonaEditor();
}

function renderPersonaNamesAndLegend() {
  const set = currentSet();
  const list = el("personaList");
  list.textContent = set.personas.filter(p => p.enabled).map((p) => p.name).join(", ");

  const legend = el("legend");
  legend.innerHTML = "";
  set.personas.forEach((p) => {
    const item = document.createElement("div");
    item.className = "swatch";
    const dot = document.createElement("span");
    dot.className = "dot";
    (dot.style as any).background = highlightToCss(p.color);
    item.appendChild(dot);
    item.appendChild(document.createTextNode(p.name));
    legend.appendChild(item);
  });
}

function renderPersonaEditor() {
  const set = currentSet();
  const container = el("personaEditor");
  container.innerHTML = "";
  set.personas.forEach((p, idx) => {
    const block = document.createElement("div");
    block.className = "result-card";
    block.innerHTML = `
      <div class="row" style="justify-content:space-between">
        <div style="display:flex;gap:8px;align-items:center">
          <input type="checkbox" id="pe-enabled-${idx}" ${p.enabled ? "checked" : ""}/>
          <strong>${p.name}</strong>
        </div>
        <div style="display:flex;gap:6px;align-items:center">
          <label style="min-width:auto">Color</label>
          <select id="pe-color-${idx}">
            ${WORD_COLORS.map(c => `<option value="${c}" ${c===p.color?"selected":""}>${c}</option>`).join("")}
          </select>
        </div>
      </div>
      <div class="row"><label>System</label><input type="text" id="pe-sys-${idx}" value="${escapeHtml(p.system)}"/></div>
      <div class="row"><label>Instruction</label><input type="text" id="pe-ins-${idx}" value="${escapeHtml(p.instruction)}"/></div>
    `;
    container.appendChild(block);
  });
}

function hydrateProviderUI() {
  (el("provider") as HTMLSelectElement).value = SETTINGS.provider.provider;
  (el("openrouterKey") as HTMLInputElement).value = SETTINGS.provider.openrouterKey || "";
  (el("model") as HTMLInputElement).value = SETTINGS.provider.model || "";
  el("openrouterKeyRow").classList.toggle("hidden", SETTINGS.provider.provider !== "openrouter");
}
function escapeHtml(s: string) { return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }

// Office bootstrap
window.addEventListener("error", (e) => { log(`[PF] window.error: ${e.message} @ ${e.filename}:${e.lineno}`); });
window.addEventListener("unhandledrejection", (ev) => { log(`[PF] unhandledrejection: ${String(ev.reason)}`); });

Office.onReady(async () => {
  log("[PF] Office.onReady → UI initialized");

  SETTINGS = loadSettings();
  populatePersonaSets();
  hydrateProviderUI();

  (el("btnSettings") as HTMLButtonElement).onclick = () => showView("view-settings");
  (el("btnBack") as HTMLButtonElement).onclick = () => showView("view-review");
  (el("toggleDebug") as HTMLButtonElement).onclick = () => {
    el("debugPanel").classList.toggle("hidden");
    (el("toggleDebug") as HTMLButtonElement).textContent =
      el("debugPanel").classList.contains("hidden") ? "Show Debug" : "Hide Debug";
  };
  (el("clearDebug") as HTMLButtonElement).onclick = () => (el("debugLog").innerHTML = "");

  (el("personaSet") as HTMLSelectElement).onchange = (ev) => {
    SETTINGS.personaSetId = (ev.target as HTMLSelectElement).value;
    saveSettings();
    populatePersonaSets();
  };
  (el("runBtn") as HTMLButtonElement).onclick = handleRunReview;
  (el("retryBtn") as HTMLButtonElement).onclick = handleRetryFailed;
  (el("exportBtn") as HTMLButtonElement).onclick = handleExportReport;

  (el("clearBtn") as HTMLButtonElement).onclick = async () => {
    if (!(await confirmAsync("Clear PF", "Remove Persona Feedback comments and highlights created by this add-in?"))) return;
    await clearPersonaFeedbackOnly();
    toast("Persona Feedback comments cleared.");
  };
  (el("clearAllBtn") as HTMLButtonElement).onclick = async () => {
    if (!(await confirmAsync("Clear ALL", "Delete ALL comments in this document? This cannot be undone."))) return;
    await clearAllComments();
    toast("All comments deleted.");
  };

  (el("provider") as HTMLSelectElement).onchange = (ev) => {
    SETTINGS.provider.provider = (ev.target as HTMLSelectElement).value as Provider;
    hydrateProviderUI();
    saveSettings();
  };
  (el("openrouterKey") as HTMLInputElement).oninput = (ev) => {
    SETTINGS.provider.openrouterKey = (ev.target as HTMLInputElement).value;
    saveSettings();
  };
  (el("model") as HTMLInputElement).oninput = (ev) => {
    SETTINGS.provider.model = (ev.target as HTMLInputElement).value;
    saveSettings();
  };
  (el("settingsPersonaSet") as HTMLSelectElement).onchange = (ev) => {
    SETTINGS.personaSetId = (ev.target as HTMLSelectElement).value;
    saveSettings();
    populatePersonaSets();
  };
  (el("saveSettings") as HTMLButtonElement).onclick = () => {
    const set = currentSet();
    set.personas.forEach((p, idx) => {
      p.enabled = (el(`pe-enabled-${idx}`) as HTMLInputElement).checked;
      p.system = (el(`pe-sys-${idx}`) as HTMLInputElement).value;
      p.instruction = (el(`pe-ins-${idx}`) as HTMLInputElement).value;
      p.color = (el(`pe-color-${idx}`) as HTMLSelectElement).value as HighlightColor;
    });
    saveSettings();
    renderPersonaNamesAndLegend();
    toast("Settings saved");
  };
  (el("restoreDefaults") as HTMLButtonElement).onclick = () => {
    const curr = currentSet().id;
    const fresh = DEFAULT_SETS.find((s) => s.id === curr);
    if (fresh) {
      const idx = SETTINGS.personaSets.findIndex((s) => s.id === curr);
      SETTINGS.personaSets[idx] = JSON.parse(JSON.stringify(fresh));
      saveSettings();
      populatePersonaSets();
      toast("Default persona set restored");
    }
  };

  showView("view-review");
});

// Actions
async function handleRunReview() {
  LAST_RESULTS = [];
  el("results").innerHTML = "";
  el("personaStatus").innerHTML = "";
  await runAllEnabledPersonas(false);
}
async function handleRetryFailed() { await runAllEnabledPersonas(true); }

async function runAllEnabledPersonas(retryOnly: boolean) {
  const set = currentSet();
  const personas = set.personas.filter((p) => p.enabled);
  if (!personas.length) { toast("No personas enabled in this set."); return; }

  const statusHost = el("personaStatus"); statusHost.innerHTML = "";
  personas.forEach((p) => {
    const row = document.createElement("div");
    row.id = `status-${p.id}`; row.className = "row";
    row.innerHTML = `
      <span style="display:inline-flex;align-items:center;gap:6px;">
        <span class="dot" style="background:${highlightToCss(p.color)}"></span>
        ${p.name}
      </span>
      <span id="badge-${p.id}" class="badge">queued</span>`;
    statusHost.appendChild(row);
  });

  const docText = await getWholeDocText();
  const total = personas.length; let done = 0;
  setProgress(0);

  for (const p of personas) {
    if (retryOnly) {
      const prev = LAST_RESULTS.find((r) => r.personaId === p.id);
      if (prev && prev.status === "done") { done++; setProgress((done / total) * 100); continue; }
    }

    setBadge(p.id, "running");
    try {
      const resp = await callLLMForPersona(p, docText);
      const normalized = normalizeResponse(resp);
      await applyCommentsAndHighlights(p, normalized, docText);
      addResultCard(p, normalized);
      LAST_RESULTS = upsertResult(LAST_RESULTS, {
        personaId: p.id, personaName: p.name, status: "done",
        scores: normalized.scores, global_feedback: normalized.global_feedback,
        comments: normalized.comments, raw: resp,
      });
      setBadge(p.id, "done");
    } catch (err: any) {
      log(`[PF] Persona ${p.name} error`, err);
      LAST_RESULTS = upsertResult(LAST_RESULTS, {
        personaId: p.id, personaName: p.name, status: "error", error: String(err?.message || err),
      });
      setBadge(p.id, "error", String(err?.message || "LLM call failed"));
    }
    done++; setProgress((done / total) * 100);
  }
  toast("Review finished.");
}
function upsertResult(arr: PersonaRunResult[], r: PersonaRunResult): PersonaRunResult[] {
  const idx = arr.findIndex((x) => x.personaId === r.personaId);
  if (idx >= 0) arr[idx] = r; else arr.push(r);
  return arr;
}
function setProgress(pct: number) { (el("progBar") as HTMLDivElement).style.width = `${Math.max(0, Math.min(100, pct))}%`; }
function setBadge(personaId: string, status: PersonaRunStatus, note?: string) {
  const b = el(`badge-${personaId}`);
  b.className = "badge " + (status === "done" ? "badge-done" : status === "error" ? "badge-failed" : "");
  b.textContent = status + (note ? ` – ${note}` : "");
}

// Word helpers
async function getWholeDocText(): Promise<string> {
  return Word.run(async (ctx) => {
    const body = ctx.document.body; body.load("text"); await ctx.sync(); return body.text || "";
  });
}
async function applyCommentsAndHighlights(
  persona: Persona,
  data: { scores: { clarity: number; tone: number; alignment: number }; global_feedback: string; comments: any[] },
  _docText: string
) {
  if (Array.isArray(data.comments)) {
    for (const c of data.comments) {
      const start = Math.max(0, Number(c.spanStart || 0));
      const end = Math.max(start, Number(c.spanEnd || start + (c.quote?.length || 0)));
      await addCommentAtRange(persona, start, end, `[${persona.name}] ${c.comment}`);
    }
  }
  await addCommentAtStart(persona, `Summary (${persona.name}): ${data.global_feedback}`);
}
async function addCommentAtRange(persona: Persona, start: number, end: number, text: string) {
  await Word.run(async (ctx) => {
    const body = ctx.document.body;
    const range = body.getRange("Start").expandTo(body.getRange("End"));
    // TS typing shim (runtime supports getSubstring on Range; typings in some office-js versions do not)
    const cRange = (range as any).getSubstring(start, end - start);
    (cRange.font as any).highlightColor = persona.color as any;
    cRange.insertComment(text);
    await ctx.sync();
  });
}
async function addCommentAtStart(persona: Persona, text: string) {
  await Word.run(async (ctx) => {
    const start = ctx.document.body.getRange("Start");
    const inserted = start.insertText(" ", Word.InsertLocation.before);
    (inserted.font as any).highlightColor = persona.color as any;
    inserted.insertComment(text);
    await ctx.sync();
  });
}
async function clearAllComments() {
  await Word.run(async (ctx) => {
    const comments = (ctx.document as any).comments; // TS typing shim
    comments.load("items"); await ctx.sync();
    for (const c of comments.items) c.delete();
    const rng = ctx.document.body.getRange("Whole");
    (rng.font as any).highlightColor = "none" as any;
    await ctx.sync();
  });
}
async function clearPersonaFeedbackOnly() {
  await Word.run(async (ctx) => {
    const comments = (ctx.document as any).comments; // TS typing shim
    comments.load("items,items/content"); await ctx.sync();
    for (const c of comments.items) { if (c.content && /^\[[^\]]+\]/.test(c.content)) c.delete(); }
    const rng = ctx.document.body.getRange("Whole");
    (rng.font as any).highlightColor = "none" as any;
    await ctx.sync();
  });
}

// LLM calls
async function callLLMForPersona(persona: Persona, docText: string): Promise<any> {
  const sys = `${persona.system}\n\n${META_PROMPT}`.trim();
  const user = `
You are acting as: ${persona.name}

INSTRUCTION:
${persona.instruction}

DOCUMENT (plain text):
${docText}
`.trim();

  const provider = SETTINGS.provider;
  log(`[PF] Calling LLM → ${provider.provider} / ${provider.model} (${persona.name})`);

  if (provider.provider === "openrouter") {
    if (!provider.openrouterKey) throw new Error("Missing OpenRouter API key.");
    const res = await fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${provider.openrouterKey}` },
      body: JSON.stringify({ model: provider.model || "openrouter/auto", messages: [
        { role: "system", content: sys }, { role: "user", content: user },
      ], temperature: 0.2 }),
    });
    if (!res.ok) throw new Error(`OpenRouter HTTP ${res.status}`);
    const json = await res.json();
    const content = json?.choices?.[0]?.message?.content ?? "";
    log(`[PF] OpenRouter raw`, json);
    return parseJsonFromText(content);
  } else {
    const res = await fetch("http://127.0.0.1:11434/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ model: provider.model || "llama3", stream: false, messages: [
        { role: "system", content: sys }, { role: "user", content: user },
      ], options: { temperature: 0.2 } }),
    });
    if (!res.ok) throw new Error(`Ollama HTTP ${res.status}`);
    const json = await res.json();
    const content = json?.message?.content ?? "";
    log(`[PF] Ollama raw`, json);
    return parseJsonFromText(content);
  }
}
function parseJsonFromText(text: string): any {
  const m = text.match(/```json([\s\S]*?)```/i) || text.match(/```([\s\S]*?)```/);
  const raw = m ? m[1] : text;
  try { return JSON.parse(raw.trim()); }
  catch { log("[PF] JSON parse error; returning raw text", { text }); throw new Error("Model returned non-JSON. Enable Debug to see raw."); }
}
function normalizeResponse(resp: any) {
  const clamp = (n: number) => Math.max(0, Math.min(100, Math.round(n)));
  const scores = {
    clarity: clamp(Number(resp?.scores?.clarity ?? 0)),
    tone: clamp(Number(resp?.scores?.tone ?? 0)),
    alignment: clamp(Number(resp?.scores?.alignment ?? 0)),
  };
  const comments = Array.isArray(resp?.comments) ? resp.comments.slice(0, 12) : [];
  const global_feedback = String(resp?.global_feedback || "");
  return { scores, comments, global_feedback };
}

// Results UI
function addResultCard(persona: Persona, data: { scores: { clarity: number; tone: number; alignment: number }; global_feedback: string }) {
  const host = el("results");
  const card = document.createElement("div");
  card.className = "result-card";
  const { clarity, tone, alignment } = data.scores;
  card.innerHTML = `
    <div class="row" style="justify-content:space-between">
      <div style="display:flex;align-items:center;gap:8px;">
        <span class="dot" style="background:${highlightToCss(persona.color)}"></span>
        <strong>${persona.name}</strong>
      </div>
      <span class="badge badge-done">done</span>
    </div>
    <div class="row"><div style="flex:1">
      <div style="display:flex;justify-content:space-between"><span>Clarity</span><span>${clarity}</span></div>
      <div class="scorebar"><div class="scorebar-fill" style="width:${clarity}%;"></div></div>
    </div></div>
    <div class="row"><div style="flex:1">
      <div style="display:flex;justify-content:space-between"><span>Tone</span><span>${tone}</span></div>
      <div class="scorebar"><div class="scorebar-fill" style="width:${tone}%;"></div></div>
    </div></div>
    <div class="row"><div style="flex:1">
      <div style="display:flex;justify-content:space-between"><span>Alignment</span><span>${alignment}</span></div>
      <div class="scorebar"><div class="scorebar-fill" style="width:${alignment}%;"></div></div>
    </div></div>
    <div style="margin-top:6px;"><em>${escapeHtml(data.global_feedback)}</em></div>`;
  host.appendChild(card);
}

// Export
async function handleExportReport() {
  const payload = { timestamp: new Date().toISOString(), set: currentSet().name, results: LAST_RESULTS, model: SETTINGS.provider };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob); const a = document.createElement("a");
  a.href = url; a.download = `persona-feedback-${Date.now()}.json`; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}
