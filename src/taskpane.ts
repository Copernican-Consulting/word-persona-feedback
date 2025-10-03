/* eslint-disable @typescript-eslint/no-explicit-any */

/**
 * Persona Feedback – Word Add-in task pane
 * - Brings back inline score bars (no external CSS needed)
 * - NO highlighting or text edits; comments only
 * - Unmatched quotes are NOT inserted in the doc; shown in the results card
 * - "Clear PF comments" deletes only comments created this session (tracked by ID)
 */

type Provider = "openrouter" | "ollama";

type Persona = {
  id: string;
  enabled: boolean;
  name: string;
  system: string;
  instruction: string;
  color: string; // UI-only dot color (legend/taskpane). No document highlight.
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
  comments?: { quote: string; spanStart: number; spanEnd: number; comment: string }[];
  unmatched?: { quote: string; comment: string }[];
  raw?: any;
  error?: string;
};

// ---------- Globals ----------
const DOT_COLORS = [
  "#fde047", "#f9a8d4", "#5eead4", "#f87171", "#93c5fd",
  "#86efac", "#c4b5fd", "#f59e0b", "#db2777", "#0d9488",
  "#b91c1c", "#1d4ed8", "#166534", "#6d28d9",
];

const LS_KEY = "pf.settings.v1";
let SETTINGS: AppSettings;
let LAST_RESULTS: PersonaRunResult[] = [];
let SESSION_COMMENT_IDS: string[] = []; // track per-session comments we insert

// ---------- Utils ----------
function byId<T extends HTMLElement>(id: string): T | null {
  const el = document.getElementById(id) as T | null;
  if (!el) console.warn(`[PF] Missing element #${id}`);
  return el;
}
function req<T extends HTMLElement>(id: string): T {
  const el = byId<T>(id);
  if (!el) throw new Error(`Required element missing: #${id}`);
  return el;
}
function log(msg: string, data?: any) {
  if (data !== undefined) console.log(msg, data); else console.log(msg);
  const panel = byId<HTMLDivElement>("debugLog");
  if (!panel) return;
  const line = document.createElement("div");
  line.style.whiteSpace = "pre-wrap";
  line.textContent = data ? `${msg} ${safeJson(data)}` : msg;
  panel.appendChild(line);
  panel.scrollTop = panel.scrollHeight;
}
function safeJson(x: any) { try { return JSON.stringify(x, null, 2); } catch { return String(x); } }
function toast(t: string) {
  const box = byId<HTMLDivElement>("toast"); if (!box) return;
  const msg = byId<HTMLSpanElement>("toastMsg"); if (msg) msg.textContent = t;
  box.style.display = "block";
  const close = byId<HTMLSpanElement>("toastClose");
  if (close) close.onclick = () => (box.style.display = "none");
  setTimeout(() => (box.style.display = "none"), 3000);
}
function showView(id: "view-review" | "view-settings") {
  const review = byId<HTMLDivElement>("view-review");
  const settings = byId<HTMLDivElement>("view-settings");
  const btnBack = byId<HTMLButtonElement>("btnBack");
  if (id === "view-review") {
    review && review.classList.remove("hidden");
    settings && settings.classList.add("hidden");
    btnBack && btnBack.classList.add("hidden");
  } else {
    review && review.classList.add("hidden");
    settings && settings.classList.remove("hidden");
    btnBack && btnBack.classList.remove("hidden");
  }
}
function confirmAsync(title: string, message: string): Promise<boolean> {
  return new Promise((resolve) => {
    const overlay = req<HTMLDivElement>("confirmOverlay");
    req<HTMLHeadingElement>("confirmTitle").textContent = title;
    req<HTMLDivElement>("confirmMessage").textContent = message;
    overlay.style.display = "flex";
    const ok = req<HTMLButtonElement>("confirmOk");
    const cancel = req<HTMLButtonElement>("confirmCancel");
    const done = (v: boolean) => {
      overlay.style.display = "none";
      ok.removeEventListener("click", onOk); cancel.removeEventListener("click", onCancel);
      resolve(v);
    };
    const onOk = () => done(true); const onCancel = () => done(false);
    ok.addEventListener("click", onOk); cancel.addEventListener("click", onCancel);
  });
}
function escapeHtml(s: string) {
  return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

// ---------- Defaults ----------
function colorAt(i: number): string { return DOT_COLORS[i % DOT_COLORS.length]; }

const META_PROMPT = `
You are a reviewer. Return ONLY valid JSON matching this schema:

{
  "scores": { "clarity": 0-100, "tone": 0-100, "alignment": 0-100 },
  "global_feedback": "short paragraph",
  "comments": [
    { "quote": "verbatim snippet", "spanStart": 0, "spanEnd": 0, "comment": "feedback" }
  ]
}

RULES:
- No markdown (unless the JSON is fenced as \`\`\`json).
- 3–8 comments with accurate "quote".
- Global feedback ~2-5 sentences.
`;

function P(name: string, system: string, instruction: string, idx: number): Persona {
  return {
    id: name.toLowerCase().replace(/[^a-z0-9]+/g, "-"),
    enabled: true, name, system, instruction, color: colorAt(idx),
  };
}

const DEFAULT_SETS: PersonaSet[] = [
  {
    id: "cross-functional-team", name: "Cross-Functional Team", personas: [
      P("Senior Manager","You are a senior manager prioritizing clarity, risks, and outcomes.","Assess clarity of goals, risks, and expected outcomes.",0),
      P("Legal","You are corporate counsel focused on compliance and risk.","Flag ambiguity, risky claims, missing disclaimers.",1),
      P("HR","You are an HR business partner concerned with tone and inclusion.","Identify exclusionary tone and suggest inclusive language.",2),
      P("Technical Lead","You are a pragmatic engineering lead.","Check feasibility, gaps, and technical risks.",3),
      P("Junior Analyst","You are a detail-oriented junior analyst.","Call out unclear logic, missing data, or inconsistent units.",4),
    ],
  },
  { id:"marketing-focus-group", name:"Marketing Focus Group", personas:[
    P("Midwest Parent","You are a pragmatic parent from the US Midwest.","React to clarity, trustworthiness, and family benefit.",0),
    P("Gen-Z Student","You are a digital-native college student.","React to tone, authenticity, and modern appeal.",1),
    P("Retired Veteran","You are a retired veteran valuing respect and responsibility.","React to credibility and plain-language clarity.",2),
    P("Small Business Owner","You are a small business owner.","React to practical value and ROI.",3),
    P("Tech-savvy Pro","You are a tech-savvy professional.","React to precision and claims that need specifics.",4),
  ]},
  { id:"startup-stakeholders", name:"Startup Stakeholders", personas:[
    P("Founder","You are a startup founder.","Push for clarity of vision, focus, and cadence.",0),
    P("CTO","You are a CTO.","Probe technical feasibility, architecture, and risks.",1),
    P("CMO","You are a CMO.","Probe messaging, audience, and differentiation.",2),
    P("VC Investor","You are a VC partner.","Probe metrics, milestones, and risks.",3),
    P("Customer","You are a prospective customer.","Probe concrete value, outcomes, and adoption risks.",4),
  ]},
  { id:"political-spectrum", name:"Political Spectrum", personas:[
    P("Democratic Socialist","You are a democratic socialist.","Assess equity, public benefit, and ethical framing.",0),
    P("Center Left","You are center-left.","Assess policy realism and social impact.",1),
    P("Centrist/Independent","You are centrist.","Assess balance, tradeoffs, and fairness.",2),
    P("Center Right","You are center-right.","Assess fiscal prudence and efficiency.",3),
    P("MAGA","You are a populist conservative.","Assess national interest and plain-language clarity.",4),
    P("Libertarian","You are libertarian.","Assess individual freedom and regulatory burden.",5),
  ]},
  { id:"product-review-board", name:"Product Review Board", personas:[
    P("PM","You are a product manager.","Assess problem framing, success metrics, and scope.",0),
    P("Design Lead","You are a design lead.","Assess user flows, accessibility, and tone.",1),
    P("Data Scientist","You are a data scientist.","Assess measurability, data risks, and validity.",2),
    P("Security","You are a security lead.","Assess data handling, privacy, and threat modeling.",3),
    P("Support Lead","You are a support lead.","Assess failure modes and user communication.",4),
  ]},
  { id:"scientific-peer-review", name:"Scientific Peer Review", personas:[
    P("Methods Reviewer","You examine methods and reproducibility.","Check experimental detail and validity.",0),
    P("Stats Reviewer","You examine statistical claims.","Check sample size, tests, and uncertainty.",1),
    P("Domain Expert","You examine domain-specific accuracy.","Check citations and assumptions.",2),
    P("Ethics Reviewer","You examine ethical and societal impact.","Check risk mitigation and consent.",3),
  ]},
  { id:"ux-research-panel", name:"UX Research Panel", personas:[
    P("New User","First-time user perspective.","Assess onboarding clarity and cognitive load.",0),
    P("Power User","Expert user perspective.","Assess efficiency and discoverability.",1),
    P("Accessibility Advocate","Accessibility lens.","Assess contrast, semantics, and alt text.",2),
  ]},
  { id:"sales-deal-desk", name:"Sales Deal Desk", personas:[
    P("Sales Director","Top-line growth focus.","Assess messaging and objections.",0),
    P("Solutions Architect","Technical fit focus.","Assess integrations, constraints, and risks.",1),
    P("Legal (Customer)","Customer counsel.","Assess indemnities, data use, and SLAs.",2),
    P("Procurement","Buyer procurement.","Assess pricing clarity and comparables.",3),
  ]},
  { id:"board-of-directors", name:"Board of Directors", personas:[
    P("Chair","Governance focus.","Assess strategy coherence and oversight.",0),
    P("Audit","Audit committee.","Assess controls, risks, and reporting.",1),
    P("Compensation","Comp committee.","Assess incentives and fairness.",2),
  ]},
  { id:"academic-committee", name:"Academic Committee", personas:[
    P("Dean","Academic leadership.","Assess alignment to mission and rigor.",0),
    P("IRB Chair","Research ethics.","Assess consent, risk, and data handling.",1),
    P("Funding Reviewer","Grant committee.","Assess merit, feasibility, and budget.",2),
  ]},
];

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

// ---------- UI render ----------
function populatePersonaSets() {
  const sets = SETTINGS.personaSets;

  const reviewSel = byId<HTMLSelectElement>("personaSet");
  if (reviewSel) {
    reviewSel.innerHTML = "";
    sets.forEach((s) => {
      const opt = document.createElement("option");
      opt.value = s.id; opt.textContent = s.name; reviewSel.appendChild(opt);
    });
    reviewSel.value = SETTINGS.personaSetId;
  }

  const settingsSel = byId<HTMLSelectElement>("settingsPersonaSet");
  if (settingsSel) {
    settingsSel.innerHTML = "";
    sets.forEach((s) => {
      const opt = document.createElement("option");
      opt.value = s.id; opt.textContent = s.name; settingsSel.appendChild(opt);
    });
    settingsSel.value = SETTINGS.personaSetId;
  }

  renderPersonaNamesAndLegend();
  renderPersonaEditor();
}

function renderPersonaNamesAndLegend() {
  const set = currentSet();
  const list = byId<HTMLSpanElement>("personaList");
  if (list) list.textContent = set.personas.filter(p => p.enabled).map(p => p.name).join(", ");

  const legend = byId<HTMLDivElement>("legend");
  if (legend) {
    legend.innerHTML = "";
    set.personas.forEach((p) => {
      const item = document.createElement("div");
      item.style.display = "flex";
      item.style.alignItems = "center";
      item.style.gap = "6px";
      const dot = document.createElement("span");
      dot.style.display = "inline-block";
      dot.style.width = "10px";
      dot.style.height = "10px";
      dot.style.borderRadius = "50%";
      (dot.style as any).background = p.color;
      item.appendChild(dot);
      item.appendChild(document.createTextNode(p.name));
      legend.appendChild(item);
    });
  }
}

function renderPersonaEditor() {
  const set = currentSet();
  const container = byId<HTMLDivElement>("personaEditor");
  if (!container) return;
  container.innerHTML = "";
  set.personas.forEach((p, idx) => {
    const block = document.createElement("div");
    block.style.border = "1px solid #e5e7eb";
    block.style.borderRadius = "8px";
    block.style.padding = "8px";
    block.style.marginBottom = "8px";
    block.innerHTML = `
      <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;">
        <div style="display:flex;gap:8px;align-items:center">
          <input type="checkbox" id="pe-enabled-${idx}" ${p.enabled ? "checked" : ""}/>
          <strong>${p.name}</strong>
        </div>
        <div style="display:flex;gap:6px;align-items:center">
          <label style="min-width:auto">Color</label>
          <input id="pe-color-${idx}" type="color" value="${toHexColor(p.color)}" />
        </div>
      </div>
      <div style="margin-top:6px;display:flex;gap:8px;align-items:center;"><label style="min-width:90px;">System</label><input style="flex:1" type="text" id="pe-sys-${idx}" value="${escapeHtml(p.system)}"/></div>
      <div style="margin-top:6px;display:flex;gap:8px;align-items:center;"><label style="min-width:90px;">Instruction</label><input style="flex:1" type="text" id="pe-ins-${idx}" value="${escapeHtml(p.instruction)}"/></div>
    `;
    container.appendChild(block);
  });
}
function toHexColor(c: string): string {
  // accept already-hex or fallback palette name strings; we just pass through or map basic names
  if (c.startsWith("#")) return c;
  // crude map for defaults
  const map: Record<string,string> = {
    yellow:"#fde047", pink:"#f9a8d4", turquoise:"#5eead4", red:"#f87171", blue:"#93c5fd",
    green:"#86efac", violet:"#c4b5fd",
  };
  return map[c] || "#fde047";
}

function hydrateProviderUI() {
  const prov = byId<HTMLSelectElement>("provider"); if (prov) prov.value = SETTINGS.provider.provider;
  const key = byId<HTMLInputElement>("openrouterKey"); if (key) key.value = SETTINGS.provider.openrouterKey || "";
  const model = byId<HTMLInputElement>("model"); if (model) model.value = SETTINGS.provider.model || "";
  const keyRow = byId<HTMLDivElement>("openrouterKeyRow");
  if (keyRow) keyRow.classList.toggle("hidden", SETTINGS.provider.provider !== "openrouter");
}

// ---------- Office bootstrap ----------
window.addEventListener("error", (e) => { log(`[PF] window.error: ${e.message} @ ${e.filename}:${e.lineno}`); });
window.addEventListener("unhandledrejection", (ev) => { log(`[PF] unhandledrejection: ${String(ev.reason)}`); });

Office.onReady(async () => {
  log("[PF] Office.onReady → UI initialized");

  SETTINGS = loadSettings();
  populatePersonaSets();
  hydrateProviderUI();

  const btnSettings = byId<HTMLButtonElement>("btnSettings");
  if (btnSettings) btnSettings.onclick = () => showView("view-settings");

  const btnBack = byId<HTMLButtonElement>("btnBack");
  if (btnBack) btnBack.onclick = () => showView("view-review");

  const toggleDebug = byId<HTMLButtonElement>("toggleDebug");
  if (toggleDebug) toggleDebug.onclick = () => {
    const panel = byId<HTMLDivElement>("debugPanel"); if (!panel) return;
    panel.classList.toggle("hidden");
    toggleDebug.textContent = panel.classList.contains("hidden") ? "Show Debug" : "Hide Debug";
  };

  const clearDebug = byId<HTMLButtonElement>("clearDebug");
  if (clearDebug) clearDebug.onclick = () => { const p = byId<HTMLDivElement>("debugLog"); if (p) p.innerHTML = ""; };

  const reviewSel = byId<HTMLSelectElement>("personaSet");
  if (reviewSel) reviewSel.onchange = (ev) => {
    SETTINGS.personaSetId = (ev.target as HTMLSelectElement).value;
    saveSettings();
    populatePersonaSets();
  };

  const runBtn = byId<HTMLButtonElement>("runBtn");
  if (runBtn) runBtn.onclick = handleRunReview;

  const retryBtn = byId<HTMLButtonElement>("retryBtn");
  if (retryBtn) retryBtn.onclick = handleRetryFailed;

  const exportBtn = byId<HTMLButtonElement>("exportBtn");
  if (exportBtn) exportBtn.onclick = handleExportReport;

  const clearBtn = byId<HTMLButtonElement>("clearBtn");
  if (clearBtn) clearBtn.onclick = async () => {
    if (!(await confirmAsync("Clear PF", "Remove Persona Feedback comments created in this session?"))) return;
    const deleted = await clearSessionComments();
    toast(deleted > 0 ? `Deleted ${deleted} comment(s).` : "No session comments to remove.");
  };

  const prov = byId<HTMLSelectElement>("provider");
  if (prov) prov.onchange = (ev) => {
    SETTINGS.provider.provider = (ev.target as HTMLSelectElement).value as Provider;
    hydrateProviderUI(); saveSettings();
  };

  const key = byId<HTMLInputElement>("openrouterKey");
  if (key) key.oninput = (ev) => {
    SETTINGS.provider.openrouterKey = (ev.target as HTMLInputElement).value; saveSettings();
  };

  const model = byId<HTMLInputElement>("model");
  if (model) model.oninput = (ev) => {
    SETTINGS.provider.model = (ev.target as HTMLInputElement).value; saveSettings();
  };

  const settingsSel = byId<HTMLSelectElement>("settingsPersonaSet");
  if (settingsSel) settingsSel.onchange = (ev) => {
    SETTINGS.personaSetId = (ev.target as HTMLSelectElement).value; saveSettings(); populatePersonaSets();
  };

  const saveSettingsBtn = byId<HTMLButtonElement>("saveSettings");
  if (saveSettingsBtn) saveSettingsBtn.onclick = () => {
    const set = currentSet();
    set.personas.forEach((p, idx) => {
      const chk = byId<HTMLInputElement>(`pe-enabled-${idx}`);
      const sys = byId<HTMLInputElement>(`pe-sys-${idx}`);
      const ins = byId<HTMLInputElement>(`pe-ins-${idx}`);
      const col = byId<HTMLInputElement>(`pe-color-${idx}`);
      if (chk) p.enabled = chk.checked;
      if (sys) p.system = sys.value;
      if (ins) p.instruction = ins.value;
      if (col) p.color = (col.value || p.color);
    });
    saveSettings();
    renderPersonaNamesAndLegend();
    toast("Settings saved");
  };

  const restoreDefaultsBtn = byId<HTMLButtonElement>("restoreDefaults");
  if (restoreDefaultsBtn) restoreDefaultsBtn.onclick = () => {
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

// ---------- Actions ----------
async function handleRunReview() {
  LAST_RESULTS = [];
  SESSION_COMMENTIDS_SAFE_CLEAR(); // reset just in case
  const res = byId<HTMLDivElement>("results"); if (res) res.innerHTML = "";
  const stat = byId<HTMLDivElement>("personaStatus"); if (stat) stat.innerHTML = "";
  await runAllEnabledPersonas(false);
}
async function handleRetryFailed() { await runAllEnabledPersonas(true); }

function setProgress(pct: number) {
  const bar = byId<HTMLDivElement>("progBar"); if (bar) bar.style.width = `${Math.max(0, Math.min(100, pct))}%`;
}
function setBadge(personaId: string, status: PersonaRunStatus, note?: string) {
  const b = byId<HTMLSpanElement>(`badge-${personaId}`); if (!b) return;
  b.className = "badge " + (status === "done" ? "badge-done" : status === "error" ? "badge-failed" : "");
  b.textContent = status + (note ? ` – ${note}` : "");
}

async function runAllEnabledPersonas(retryOnly: boolean) {
  const set = currentSet();
  const personas = set.personas.filter((p) => p.enabled);
  if (!personas.length) { toast("No personas enabled in this set."); return; }

  const statusHost = byId<HTMLDivElement>("personaStatus");
  if (statusHost) {
    statusHost.innerHTML = "";
    personas.forEach((p) => {
      const row = document.createElement("div");
      row.id = `status-${p.id}`;
      row.style.display = "flex";
      row.style.justifyContent = "space-between";
      row.style.marginBottom = "4px";
      row.innerHTML = `
        <span style="display:inline-flex;align-items:center;gap:6px;">
          <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${p.color};"></span>
          ${p.name}
        </span>
        <span id="badge-${p.id}" class="badge">queued</span>`;
      statusHost.appendChild(row);
    });
  }

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
      const { matched, unmatched } = await applyCommentsForMatchesOnly(p, normalized);
      addResultCard(p, normalized, unmatched); // show unmatched only in pane
      upsertResult({
        personaId: p.id, personaName: p.name, status: "done",
        scores: normalized.scores, global_feedback: normalized.global_feedback,
        comments: matched, unmatched, raw: resp,
      });
      setBadge(p.id, "done");
    } catch (err: any) {
      log(`[PF] Persona ${p.name} error`, err);
      upsertResult({
        personaId: p.id, personaName: p.name, status: "error", error: String(err?.message || err),
      });
      setBadge(p.id, "error", String(err?.message || "LLM call failed"));
    }
    done++; setProgress((done / total) * 100);
  }
  toast("Review finished.");
}
function upsertResult(r: PersonaRunResult) {
  const idx = LAST_RESULTS.findIndex((x) => x.personaId === r.personaId);
  if (idx >= 0) LAST_RESULTS[idx] = r; else LAST_RESULTS.push(r);
}

// ---------- Word helpers ----------
async function getWholeDocText(): Promise<string> {
  return Word.run(async (ctx) => {
    const body = ctx.document.body; body.load("text"); await ctx.sync(); return body.text || "";
  });
}

/**
 * Insert ONLY comments for quotes that can be found via body.search.
 * Returns matched comments (as inserted) and a list of unmatched {quote, comment}.
 */
async function applyCommentsForMatchesOnly(
  persona: Persona,
  data: { scores: { clarity: number; tone: number; alignment: number }; global_feedback: string; comments: any[] }
): Promise<{
  matched: { quote: string; spanStart: number; spanEnd: number; comment: string }[];
  unmatched: { quote: string; comment: string }[];
}> {
  // Add a summary comment at the start (persona-level) WITHOUT modifying text/highlight
  await addCommentAtStart(persona, `Summary (${persona.name}): ${data.global_feedback}`);

  const matched: { quote: string; spanStart: number; spanEnd: number; comment: string }[] = [];
  const unmatched: { quote: string; comment: string }[] = [];

  if (!Array.isArray(data.comments) || data.comments.length === 0) {
    log(`[PF] ${persona.name}: no inline comments returned`);
    return { matched, unmatched };
  }

  for (const [i, c] of data.comments.entries()) {
    const quote = String(c.quote || "").trim();
    const note = String(c.comment || "").trim();
    if (!quote || quote.length < 3) {
      log(`[PF] ${persona.name}: comment #${i+1} has empty/short quote; skipping`, c);
      continue;
    }

    const placed = await addCommentBySearchingQuote(persona, quote, note);
    if (placed) {
      matched.push({ quote, spanStart: Number(c.spanStart||0), spanEnd: Number(c.spanEnd||0), comment: note });
    } else {
      unmatched.push({ quote, comment: note });
    }
  }

  return { matched, unmatched };
}

/**
 * Searches for the quoted text and attaches a comment on the first matching range.
 * Returns true if a match was found and a comment inserted.
 */
async function addCommentBySearchingQuote(persona: Persona, quote: string, text: string): Promise<boolean> {
  return Word.run(async (ctx) => {
    const body = ctx.document.body;
    const ranges = body.search(quote, {
      matchCase: false, matchWholeWord: false, matchWildcards: false, ignoreSpace: false, ignorePunct: false,
    });
    ranges.load("items"); await ctx.sync();

    const count = ranges.items.length;
    log(`[PF] search "${quote.slice(0, 60)}${quote.length>60?"…":""}" → ${count} match(es)`);

    if (count === 0) return false;

    const r = ranges.items[0];
    const comment = r.insertComment(`[${persona.name}] ${text}`);
    // Track comment ID (for same-session clear)
    comment.load("id"); await ctx.sync();
    if (comment.id) SESSION_COMMENT_IDS.push(comment.id);
    return true;
  });
}

async function addCommentAtStart(persona: Persona, text: string) {
  return Word.run(async (ctx) => {
    const start = ctx.document.body.getRange("Start");
    const comment = start.insertComment(`[${persona.name}] ${text}`);
    comment.load("id"); await ctx.sync();
    if (comment.id) SESSION_COMMENT_IDS.push(comment.id);
  });
}

async function clearSessionComments(): Promise<number> {
  if (!SESSION_COMMENT_IDS.length) return 0;
  return Word.run(async (ctx) => {
    let deleted = 0;
    // Try document.comments (if supported); if not, we still try direct objects by id
    const docAny: any = ctx.document as any;
    const commentsColl = docAny.comments;
    if (commentsColl) commentsColl.load("items"); // may be undefined on older sets; that's ok
    await ctx.sync().catch(() => { /* swallow */ });

    if (commentsColl && commentsColl.items) {
      for (const c of commentsColl.items) {
        c.load("id"); // get ID to compare
      }
      await ctx.sync().catch(() => { /* ignore */ });
      for (const c of commentsColl.items) {
        if (SESSION_COMMENT_IDS.includes(c.id)) { c.delete(); deleted++; }
      }
      await ctx.sync().catch(() => { /* ignore */ });
    } else {
      // Fallback: best-effort – we cannot enumerate comments; session IDs can’t be fetched back.
      // Nothing we can do except inform the user.
      log("[PF] clearSessionComments: document.comments not available on this build; limited cleanup.");
    }
    // reset session list regardless to avoid double-delete attempts
    SESSION_COMMENT_IDS = [];
    return deleted;
  });
}
function SESSION_COMMENTIDS_SAFE_CLEAR() {
  SESSION_COMMENT_IDS = [];
}

// ---------- Networking ----------
function withTimeout<T>(p: Promise<T>, ms = 45000): Promise<T> {
  return new Promise((resolve, reject) => {
    const t = setTimeout(() => reject(new Error(`Request timed out after ${ms}ms`)), ms);
    p.then((v) => { clearTimeout(t); resolve(v); }, (e) => { clearTimeout(t); reject(e); });
  });
}
async function fetchJson(url: string, init: RequestInit): Promise<{ ok: boolean; status: number; body: any; text?: string }> {
  try {
    const res = await withTimeout(fetch(url, init));
    let body: any = null;
    let text = "";
    try {
      const ct = res.headers.get("content-type") || "";
      if (ct.includes("application/json")) body = await res.json();
      else { text = await res.text(); try { body = JSON.parse(text); } catch { /* keep text */ } }
    } catch (e) {
      text = await res.text().catch(() => "");
    }
    return { ok: res.ok, status: res.status, body, text };
  } catch (e: any) {
    log("[PF] fetchJson network error", { message: e?.message || String(e) });
    throw e;
  }
}

// ---------- LLM ----------
async function callLLMForPersona(persona: Persona, docText: string): Promise<any> {
  const sys = `${persona.system}\n\n${META_PROMPT}`.trim();
  const user = `You are acting as: ${persona.name}\n\nINSTRUCTION:\n${persona.instruction}\n\nDOCUMENT (plain text):\n${docText}`.trim();

  const provider = SETTINGS.provider;
  log(`[PF] Calling LLM → ${provider.provider} / ${provider.model} (${persona.name})`);

  if (provider.provider === "openrouter") {
    if (!provider.openrouterKey) throw new Error("Missing OpenRouter API key.");

    const res = await fetchJson("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${provider.openrouterKey}`,
        "HTTP-Referer": (typeof window !== "undefined" ? window.location.origin : "https://word-persona-feedback.vercel.app"),
        "X-Title": "Persona Feedback Add-in",
      },
      body: JSON.stringify({
        model: provider.model || "openrouter/auto",
        messages: [
          { role: "system", content: sys },
          { role: "user", content: user },
        ],
        temperature: 0.2,
      }),
    });

    if (!res.ok) {
      log("[PF] OpenRouter non-OK", { status: res.status, body: res.body, text: res.text });
      throw new Error(`OpenRouter HTTP ${res.status}: ${res.text || safeJson(res.body)}`);
    }
    const content = res.body?.choices?.[0]?.message?.content ?? "";
    log(`[PF] OpenRouter raw`, res.body);
    return parseJsonFromText(content);

  } else {
    const res = await fetchJson("http://127.0.0.1:11434/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: provider.model || "llama3",
        stream: false,
        messages: [
          { role: "system", content: sys },
          { role: "user", content: user },
        ],
        options: { temperature: 0.2 },
      }),
    });

    if (!res.ok) {
      log("[PF] Ollama non-OK", { status: res.status, body: res.body, text: res.text });
      throw new Error(`Ollama HTTP ${res.status}: ${res.text || safeJson(res.body)}`);
    }
    const content = res.body?.message?.content ?? "";
    log(`[PF] Ollama raw`, res.body);
    return parseJsonFromText(content);
  }
}

function parseJsonFromText(text: string): any {
  const m = text.match(/```json([\s\S]*?)```/i) || text.match(/```([\s\S]*?)```/);
  const raw = m ? m[1] : text;
  try { return JSON.parse(raw.trim()); }
  catch {
    log("[PF] JSON parse error; full text follows", { text });
    throw new Error("Model returned non-JSON. See Debug for raw output.");
  }
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

// ---------- Results UI ----------
function scoreBar(label: string, value: number) {
  const pct = Math.max(0, Math.min(100, value|0));
  return `
  <div style="display:flex;justify-content:space-between;font-size:12px;margin-top:4px;"><span>${label}</span><span>${pct}</span></div>
  <div style="width:100%;height:8px;background:#e5e7eb;border-radius:999px;overflow:hidden;">
    <div style="height:100%;width:${pct}%;background:#3b82f6;"></div>
  </div>`;
}

function addResultCard(
  persona: Persona,
  data: { scores: { clarity: number; tone: number; alignment: number }; global_feedback: string },
  unmatched?: { quote: string; comment: string }[]
) {
  const host = byId<HTMLDivElement>("results"); if (!host) return;
  const card = document.createElement("div");
  card.style.border = "1px solid #e5e7eb";
  card.style.borderRadius = "10px";
  card.style.padding = "10px";
  card.style.marginBottom = "10px";

  const { clarity, tone, alignment } = data.scores;
  card.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
      <div style="display:flex;align-items:center;gap:8px;">
        <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${persona.color};"></span>
        <strong>${persona.name}</strong>
      </div>
      <span class="badge badge-done">done</span>
    </div>
    ${scoreBar("Clarity", clarity)}
    ${scoreBar("Tone", tone)}
    ${scoreBar("Alignment", alignment)}
    <div style="margin-top:8px;"><em>${escapeHtml(data.global_feedback)}</em></div>
    ${unmatched && unmatched.length ? `
      <div style="margin-top:8px;">
        <div style="font-weight:600;margin-bottom:4px;">Unmatched quotes (not inserted):</div>
        <ul style="margin:0 0 0 16px;padding:0;list-style:disc;">
          ${unmatched.slice(0,6).map(u => `<li><span style="color:#6b7280">"${escapeHtml(u.quote.slice(0,120))}${u.quote.length>120?"…":""}"</span><br/><span>${escapeHtml(u.comment)}</span></li>`).join("")}
        </ul>
      </div>` : ``}
  `;
  host.appendChild(card);
}

async function handleExportReport() {
  const payload = { timestamp: new Date().toISOString(), set: currentSet().name, results: LAST_RESULTS, model: SETTINGS.provider };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob); const a = document.createElement("a");
  a.href = url; a.download = `persona-feedback-${Date.now()}.json`; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}

// ---------- Progress / badges host ----------
function setBadgesHost(personas: Persona[]) {
  const statusHost = byId<HTMLDivElement>("personaStatus");
  if (!statusHost) return;
  statusHost.innerHTML = "";
  personas.forEach((p) => {
    const row = document.createElement("div");
    row.style.display = "flex";
    row.style.justifyContent = "space-between";
    row.style.marginBottom = "4px";
    row.innerHTML = `
      <span style="display:inline-flex;align-items:center;gap:6px;">
        <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${p.color};"></span>
        ${p.name}
      </span>
      <span id="badge-${p.id}" class="badge">queued</span>`;
    statusHost.appendChild(row);
  });
}
