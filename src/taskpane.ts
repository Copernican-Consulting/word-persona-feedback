// src/taskpane.ts
/* eslint-disable @typescript-eslint/no-explicit-any */

// Load CSS via webpack (prevents 404 if the HTML links a non-existent taskpane.css)
import "./ui.css";

// -----------------------------
// Types
// -----------------------------
type ChatMessage = { role: "system" | "user" | "assistant"; content: string };

type Persona = {
  id: string;
  name: string;
  color?: string; // hsl/hex
  system: string;
  instruction: string;
  enabled?: boolean;
};

type PersonaSet = { id: string; name: string; personas: Persona[]; version?: number };

type FeedbackJSON = {
  scores: Record<string, number>;
  global_feedback: string;
  comments: Array<{ quote: string; comment: string; category?: string }>;
};

type ProviderConfig =
  | {
      kind: "openrouter";
      apiKey: string;
      model: string; // e.g. "gpt-4o-mini"
    }
  | {
      kind: "ollama";
      host?: string; // default http://127.0.0.1:11434
      model: string; // e.g. "llama3.1:8b"
    };

type AppSettings = {
  provider?: ProviderConfig;
  selectedSetId?: string;
  personaEnabled?: Record<string, boolean>; // personaId -> enabled
};

// -----------------------------
// DOM helpers (resilient to varying ids)
// -----------------------------
const $ = <T extends HTMLElement = HTMLElement>(id: string) =>
  document.getElementById(id) as T | null;

function firstById<T extends HTMLElement = HTMLElement>(ids: string[]): T | null {
  for (const id of ids) {
    const el = document.getElementById(id);
    if (el) return el as T;
  }
  return null;
}

function text(el: HTMLElement | null, s: string) {
  if (el) el.textContent = s;
}

function show(el: HTMLElement | null, on: boolean) {
  if (!el) return;
  el.style.display = on ? "" : "none";
}

function toast(msg: string) {
  console.log("[PF]", msg);
  const t = firstById<HTMLDivElement>(["toast", "pfToast"]);
  if (!t) return;
  t.textContent = msg;
  t.classList.add("show");
  setTimeout(() => t.classList.remove("show"), 2500);
}

async function confirmAsync(message: string): Promise<boolean> {
  // Use custom modal if present; else native confirm()
  const modal = firstById<HTMLDialogElement>(["confirmModal", "pfConfirm"]);
  if (!modal) return Promise.resolve(window.confirm(message));
  const prompt = modal.querySelector<HTMLElement>("[data-confirm-message]");
  if (prompt) prompt.textContent = message;
  const okBtn = modal.querySelector<HTMLButtonElement>("[data-confirm-ok]");
  const cancelBtn = modal.querySelector<HTMLButtonElement>("[data-confirm-cancel]");
  return new Promise<boolean>((resolve) => {
    const onOk = () => {
      cleanup();
      resolve(true);
    };
    const onCancel = () => {
      cleanup();
      resolve(false);
    };
    function cleanup() {
      okBtn?.removeEventListener("click", onOk);
      cancelBtn?.removeEventListener("click", onCancel);
      // @ts-ignore
      if (modal.close) modal.close();
    }
    // @ts-ignore
    if (modal.showModal) modal.showModal();
    okBtn?.addEventListener("click", onOk);
    cancelBtn?.addEventListener("click", onCancel);
  });
}

// -----------------------------
// Inline default persona sets (kept minimal; merged with ./personas.ts)
// -----------------------------
const INLINE_SETS: PersonaSet[] = [
  {
    id: "cross-functional",
    name: "Cross-Functional Team",
    personas: [
      p("senior-manager", "Senior Manager", "hsl(0, 70%, 55%)", `You are a senior manager who values clarity, impact, and risk awareness.`, `Score for clarity, completeness, and business impact. Be concise; number your critiques.`),
      p("legal", "Legal", "hsl(210, 70%, 50%)", `You are in-house counsel focused on compliance and risk.`, `Flag risky or unverifiable claims. Ask for sources and qualifiers.`),
      p("tech-lead", "Technical Lead", "hsl(140, 60%, 45%)", `You are a pragmatic technical lead.`, `Call out feasibility risks, missing constraints, and unrealistic timelines.`),
      p("hr", "HR Partner", "hsl(35, 80%, 55%)", `You are an HR partner focused on inclusion and clarity.`, `Flag jargon and unclear asks; push for inclusive phrasing.`),
      p("junior-analyst", "Junior Analyst", "hsl(280, 60%, 55%)", `You are an eager analyst.`, `Ask sharp questions where data or method is unclear.`)
    ]
  },
  {
    id: "marketing-focus",
    name: "Marketing Focus Group",
    personas: [
      p("midwest-parent", "Midwest Parent", "hsl(5, 70%, 55%)", `You are a practical parent from the Midwest.`, `Do you trust this? Is it practical and affordable?`),
      p("genz-student", "Gen-Z Student", "hsl(260, 70%, 60%)", `You are a Gen-Z student.`, `Is this exciting and authentic? Does it feel cringe?`),
      p("retired-veteran", "Retired Veteran", "hsl(210, 30%, 45%)", `You are a retired veteran.`, `Look for respect, directness, and clear benefits.`),
      p("smb-owner", "Small Business Owner", "hsl(30, 80%, 55%)", `You run a small local business.`, `Does this help revenue or save time? What’s the ROI?`),
      p("tech-pro", "Tech Professional", "hsl(160, 50%, 45%)", `You are a tech professional.`, `Precision, evidence, and no fluff.`)
    ]
  }
];

// helper to build personas
function p(id: string, name: string, color: string, system: string, instruction: string): Persona {
  return { id, name, color, system, instruction, enabled: true };
}

// -----------------------------
// Global state
// -----------------------------
let ALL_SETS: PersonaSet[] = [];
let SETTINGS: AppSettings = {};
let ACTIVE_SET: PersonaSet | null = null;
let ACTIVE_PERSONAS: Persona[] = [];

// -----------------------------
// Office bootstrap
// -----------------------------
Office.onReady(async () => {
  try {
    initViewSwitching();
    loadSettings();
    await loadAndMergePersonaSets(); // merges ./personas.ts with inline defaults
    renderPersonaControls();
    wireButtons();
    toast("UI initialized");
  } catch (err) {
    console.error(err);
    toast("Initialization error (see console)");
  }
});

// -----------------------------
// View switching (Fix #3)
// -----------------------------
function initViewSwitching() {
  // Try to discover main/settings containers by multiple common ids
  const main = firstById<HTMLElement>(["mainView", "reviewView", "mainSection", "reviewSection", "main", "view-main", "pfMain"]);
  const settings = firstById<HTMLElement>(["settingsView", "settingsSection", "settings", "view-settings", "pfSettings"]);

  function switchView(which: "main" | "settings") {
    show(main, which === "main");
    show(settings, which === "settings");
  }

  // Store on window for debugging if needed
  (window as any)._pfSwitchView = switchView;

  // Initially show main
  switchView("main");

  // Settings button toggles
  const settingsBtn = firstById<HTMLButtonElement>(["settingsBtn", "btnSettings", "pfSettingsBtn"]);
  if (settingsBtn) {
    settingsBtn.onclick = () => {
      const isSettingsVisible = settings && settings.style.display !== "none";
      switchView(isSettingsVisible ? "main" : "settings");
    };
  }

  // Optional explicit back button if present
  const backBtn = firstById<HTMLButtonElement>(["backBtn", "backToMainBtn", "closeSettingsBtn", "homeBtn", "pfBackBtn"]);
  if (backBtn) backBtn.onclick = () => switchView("main");

  // Esc key returns to main if in settings
  window.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      const isSettingsVisible = settings && settings.style.display !== "none";
      if (isSettingsVisible) switchView("main");
    }
  });
}

// -----------------------------
// Settings load/save
// -----------------------------
const LS_KEY = "pf.settings.v1";

function loadSettings() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    SETTINGS = raw ? (JSON.parse(raw) as AppSettings) : {};
  } catch {
    SETTINGS = {};
  }
}

function saveSettings() {
  localStorage.setItem(LS_KEY, JSON.stringify(SETTINGS));
}

// -----------------------------
// Merge personas (Fix #2)
// -----------------------------
async function loadAndMergePersonaSets() {
  // Start with inline
  let merged: PersonaSet[] = JSON.parse(JSON.stringify(INLINE_SETS));

  // Try dynamic import of ./personas.ts
  try {
    const mod: any = await import(/* webpackChunkName: "personas" */ "./personas");
    const libSets: PersonaSet[] =
      (mod && (mod.PERSONA_SETS as PersonaSet[])) ||
      (mod && (mod.default as PersonaSet[])) ||
      [];

    merged = mergePersonaSets(merged, libSets);
  } catch (e) {
    console.warn("[PF] No external personas found or failed to import ./personas:", e);
  }

  ALL_SETS = merged;

  // Select active set
  const desired = SETTINGS.selectedSetId || merged[0]?.id;
  ACTIVE_SET = merged.find((s) => s.id === desired) || merged[0] || null;

  // Enabled personas (respect user toggles if any)
  const enabledMap = SETTINGS.personaEnabled || {};
  ACTIVE_PERSONAS =
    ACTIVE_SET?.personas.map((pp) => ({ ...pp, enabled: enabledMap[pp.id] ?? true })) ?? [];
}

function mergePersonaSets(a: PersonaSet[], b: PersonaSet[]): PersonaSet[] {
  const byId = new Map<string, PersonaSet>();
  const norm = (s: string) => s.trim().toLowerCase();

  const put = (set: PersonaSet) => {
    const key = set.id || norm(set.name);
    const existing = byId.get(key);
    if (!existing) {
      byId.set(key, { id: key, name: set.name, version: set.version, personas: [...set.personas] });
      return;
    }
    // merge personas (by id or name)
    const have = new Map<string, Persona>();
    for (const p of existing.personas) have.set(p.id || norm(p.name), p);
    for (const p of set.personas) {
      const k = p.id || norm(p.name);
      if (!have.has(k)) {
        existing.personas.push(p);
        have.set(k, p);
      } else {
        // prefer the richer persona (longer system/instruction)
        const cur = have.get(k)!;
        const better =
          (p.system?.length ?? 0) + (p.instruction?.length ?? 0) >
          (cur.system?.length ?? 0) + (cur.instruction?.length ?? 0)
            ? p
            : cur;
        Object.assign(cur, better);
      }
    }
  };

  [...a, ...b].forEach(put);
  return Array.from(byId.values());
}

// -----------------------------
// Render persona UI (dropdown + checkboxes)
// -----------------------------
function renderPersonaControls() {
  const setSelect =
    firstById<HTMLSelectElement>(["personaSetSelect", "personaSet", "pfPersonaSet"]) ||
    createSelectMount();

  // Populate dropdown
  setSelect.innerHTML = "";
  for (const s of ALL_SETS) {
    const opt = document.createElement("option");
    opt.value = s.id;
    opt.textContent = s.name;
    if (ACTIVE_SET?.id === s.id) opt.selected = true;
    setSelect.appendChild(opt);
  }

  setSelect.onchange = () => {
    SETTINGS.selectedSetId = setSelect.value;
    saveSettings();
    ACTIVE_SET = ALL_SETS.find((s) => s.id === setSelect.value) || null;
    const enabledMap = SETTINGS.personaEnabled || {};
    ACTIVE_PERSONAS =
      ACTIVE_SET?.personas.map((pp) => ({ ...pp, enabled: enabledMap[pp.id] ?? true })) ?? [];
    renderPersonaList();
  };

  // Render the checklist of personas
  renderPersonaList();
}

function createSelectMount(): HTMLSelectElement {
  const mount =
    firstById<HTMLDivElement>(["personaControls", "personaControlBar", "pfPersonaBar"]) ||
    document.body;
  const sel = document.createElement("select");
  sel.id = "personaSetSelect";
  sel.className = "pf-select";
  mount.appendChild(sel);
  return sel;
}

function renderPersonaList() {
  const list =
    firstById<HTMLDivElement>(["personaChecklist", "personaList", "pfPersonaList"]) ||
    createPersonaListMount();

  list.innerHTML = "";

  for (const pp of ACTIVE_PERSONAS) {
    const row = document.createElement("label");
    row.className = "pf-row pf-persona-row";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = !!pp.enabled;
    cb.onchange = () => {
      pp.enabled = cb.checked;
      SETTINGS.personaEnabled = SETTINGS.personaEnabled || {};
      SETTINGS.personaEnabled[pp.id] = !!pp.enabled;
      saveSettings();
    };

    const swatch = document.createElement("span");
    swatch.className = "pf-swatch";
    if (pp.color) swatch.style.background = pp.color;

    const name = document.createElement("span");
    name.className = "pf-persona-name";
    name.textContent = pp.name;

    row.appendChild(cb);
    row.appendChild(swatch);
    row.appendChild(name);
    list.appendChild(row);
  }
}

function createPersonaListMount(): HTMLDivElement {
  const mount =
    firstById<HTMLDivElement>(["personaControls", "personaControlBar", "pfPersonaBar"]) ||
    document.body;
  const div = document.createElement("div");
  div.id = "personaChecklist";
  div.className = "pf-list";
  mount.appendChild(div);
  return div;
}

// -----------------------------
// Wire buttons (includes Fix #1 for clearAllBtn)
// -----------------------------
function wireButtons() {
  const runBtn = firstById<HTMLButtonElement>(["runBtn", "btnRun", "pfRun"]);
  if (runBtn) runBtn.onclick = () => void runAllEnabledPersonas();

  const retryBtn = firstById<HTMLButtonElement>(["retryBtn", "btnRetry", "pfRetry"]);
  if (retryBtn) retryBtn.onclick = () => void runAllEnabledPersonas();

  const exportBtn = firstById<HTMLButtonElement>(["exportBtn", "btnExport", "pfExport"]);
  if (exportBtn) exportBtn.onclick = () => void exportReport();

  const clearBtn = firstById<HTMLButtonElement>(["clearBtn", "btnClear", "pfClear"]);
  if (clearBtn) clearBtn.onclick = () => void clearOurCommentsOnly();

  // FIX #1: Wire Clear ALL Comments (hard delete)
  const clearAllBtn = firstById<HTMLButtonElement>(["clearAllBtn", "btnClearAll", "pfClearAll"]);
  if (clearAllBtn) {
    clearAllBtn.onclick = async () => {
      const ok = await confirmAsync("Delete ALL comments in the document?");
      if (!ok) return;
      try {
        const n = await clearAllComments();
        toast(n > 0 ? `Deleted ${n} comment(s).` : "No comments found.");
      } catch (e) {
        console.error(e);
        toast("Failed to clear comments");
      }
    };
  }

  const debugBtn = firstById<HTMLButtonElement>(["debugBtn", "btnDebug", "pfDebug"]);
  const debugPanel = firstById<HTMLDivElement>(["debugPanel", "pfDebugPanel"]);
  if (debugBtn && debugPanel) {
    debugBtn.onclick = () => {
      const vis = debugPanel.style.display !== "none";
      show(debugPanel, !vis);
    };
  }
}

// -----------------------------
// Core run loop
// -----------------------------
async function runAllEnabledPersonas() {
  if (!ACTIVE_SET) {
    toast("No persona set selected");
    return;
  }
  const enabled = ACTIVE_PERSONAS.filter((p) => p.enabled);
  if (enabled.length === 0) {
    toast("Enable at least one persona");
    return;
  }

  const docText = await getDocumentText();
  if (!docText || !docText.trim()) {
    toast("Document appears empty");
    return;
  }

  const progress = firstById<HTMLDivElement>(["progBar", "progressBar", "pfProgress"]);
  let i = 0;
  for (const persona of enabled) {
    text(progress, `Running ${persona.name} (${++i}/${enabled.length})…`);
    try {
      const out = await reviewWithPersona(persona, docText);
      await insertComments(persona, out, docText);
      appendResults(persona, out);
    } catch (err) {
      console.error(err);
      appendError(persona, err as Error);
    }
  }
  text(progress, "Done");
}

// -----------------------------
// Provider calls
// -----------------------------
async function reviewWithPersona(persona: Persona, documentText: string): Promise<FeedbackJSON> {
  const provider = SETTINGS.provider || {
    kind: "openrouter",
    model: "gpt-4o-mini",
    apiKey: ""
  };

  const messages: ChatMessage[] = [
    { role: "system", content: persona.system },
    {
      role: "user",
      content:
        persona.instruction +
        "\n\nEvaluate the following document and return ONLY valid JSON with shape:\n" +
        `{"scores": {"<metric>": 1-5}, "global_feedback": "string", "comments":[{"quote":"string","comment":"string","category":"optional"}]}\n` +
        "Document:\n<<<DOC\n" +
        documentText +
        "\nDOC>>>"
    }
  ];

  let raw = "";
  if (provider.kind === "openrouter") {
    if (!provider.apiKey) {
      // Let users paste their key into a prompt (kept simple)
      const k = window.prompt("Enter your OpenRouter API key (kept in-memory for this session):", "");
      if (!k) throw new Error("No OpenRouter API key provided");
      provider.apiKey = k;
      SETTINGS.provider = provider;
      saveSettings();
    }
    raw = await callOpenRouter(provider.apiKey, provider.model, messages);
  } else {
    raw = await callOllama(provider.host || "http://127.0.0.1:11434", provider.model, messages);
  }

  const json = coerceJSON(raw);
  validateOutput(json);
  return json;
}

async function callOpenRouter(apiKey: string, model: string, messages: ChatMessage[]): Promise<string> {
  const res = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
      "HTTP-Referer": "https://example.com",
      "X-Title": "Word Persona Feedback"
    },
    body: JSON.stringify({
      model,
      messages,
      temperature: 0.2
    })
  });
  if (!res.ok) throw new Error(`OpenRouter error ${res.status}`);
  const data = await res.json();
  return data.choices?.[0]?.message?.content || "";
}

async function callOllama(host: string, model: string, messages: ChatMessage[]): Promise<string> {
  const res = await fetch(`${host.replace(/\/$/, "")}/api/chat`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ model, messages, options: { temperature: 0.2 } })
  });
  if (!res.ok) throw new Error(`Ollama error ${res.status}`);
  const data = await res.json();
  return data.message?.content || "";
}

// -----------------------------
// JSON coercion & validation
// -----------------------------
function coerceJSON(s: string): FeedbackJSON {
  // strip code fences or prose
  const m = s.match(/\{[\s\S]*\}$/);
  const body = m ? m[0] : s;
  try {
    return JSON.parse(body);
  } catch {
    // last-ditch: attempt to fix trailing commas/quotes
    const fixed = body
      .replace(/[\u201C\u201D]/g, '"')
      .replace(/,\s*}/g, "}")
      .replace(/,\s*]/g, "]");
    return JSON.parse(fixed);
  }
}

function validateOutput(o: any): asserts o is FeedbackJSON {
  if (!o || typeof o !== "object") throw new Error("Model returned no JSON");
  if (!o.scores || typeof o.scores !== "object") o.scores = {};
  if (typeof o.global_feedback !== "string") o.global_feedback = "";
  if (!Array.isArray(o.comments)) o.comments = [];
  // Clamp
  for (const k of Object.keys(o.scores)) {
    const v = Number(o.scores[k]);
    if (!Number.isFinite(v)) delete o.scores[k];
    else o.scores[k] = Math.max(1, Math.min(5, Math.round(v)));
  }
  // Trim long quotes
  o.comments = o.comments.map((c: any) => ({
    quote: String(c.quote || "").slice(0, 240),
    comment: String(c.comment || ""),
    category: c.category ? String(c.category) : undefined
  }));
}

// -----------------------------
// Word integration: text, comments
// -----------------------------
async function getDocumentText(): Promise<string> {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    return body.text || "";
  });
}

async function insertComments(persona: Persona, out: FeedbackJSON, docText: string) {
  if (!out.comments?.length) return;

  const prefix = makePersonaPrefix(persona);
  const notes = out.comments;

  await Word.run(async (context) => {
    for (const c of notes) {
      const quote = c.quote?.trim();
      let range: Word.Range | null = null;

      if (quote && quote.length >= 3) {
        const results = context.document.body.search(quote, {
          matchCase: false,
          matchWholeWord: false,
          matchWildcards: false
        });
        results.load("items");
        await context.sync();

        if (results.items.length > 0) {
          range = results.items[0];
        } else {
          // fuzzy: pick a small snippet
          const snippet = fuzzySnippet(quote, docText);
          if (snippet) {
            const fuzzyResults = context.document.body.search(snippet, {
              matchCase: false,
              matchWholeWord: false
            });
            fuzzyResults.load("items");
            await context.sync();
            if (fuzzyResults.items.length > 0) range = fuzzyResults.items[0];
          }
        }
      }

      const payload = `${prefix} ${c.comment}${c.category ? ` [${c.category}]` : ""}`;

      if (range) {
        range.insertComment(payload);
      } else {
        // attach at start to avoid losing feedback
        context.document.body.getRange("start").insertComment(`${prefix} (unmatched) ${payload}`);
      }
      await context.sync();
    }
  });
}

function makePersonaPrefix(p: Persona) {
  // Simple colored square approximation using emoji (visible across Word surfaces)
  const mark = "🟦";
  return `${mark} ${p.name}`;
}

function fuzzySnippet(quote: string, docText: string): string | null {
  // try middle 30 chars
  if (quote.length <= 30) return quote;
  const mid = Math.floor(quote.length / 2);
  const snip = quote.slice(Math.max(0, mid - 15), mid + 15);
  return docText.includes(snip) ? snip : null;
}

// Delete ALL comments (Fix #1)
async function clearAllComments(): Promise<number> {
  return Word.run(async (context) => {
    const comments = context.document.comments;
    comments.load("items");
    await context.sync();
    const n = comments.items.length;
    for (const c of comments.items) c.delete();
    await context.sync();
    return n;
  });
}

// Delete only comments we added (prefix marker)
async function clearOurCommentsOnly(): Promise<number> {
  return Word.run(async (context) => {
    const comments = context.document.comments;
    comments.load("items");
    await context.sync();
    let n = 0;
    for (const c of comments.items) {
      c.content.load("text");
    }
    await context.sync();
    for (const c of comments.items) {
      const t = (c as any).content?.text ?? "";
      if (/^🟦 /.test(t)) {
        c.delete();
        n++;
      }
    }
    await context.sync();
    toast(`Deleted ${n} comment(s) from Feedback Personas`);
    return n;
  });
}

// -----------------------------
// Results panel (simple, non-intrusive)
// -----------------------------
function appendResults(persona: Persona, out: FeedbackJSON) {
  const container = firstById<HTMLDivElement>(["results", "pfResults"]);
  if (!container) return;

  const card = document.createElement("div");
  card.className = "pf-card";

  const h = document.createElement("div");
  h.className = "pf-card-title";
  h.textContent = persona.name;
  card.appendChild(h);

  if (out.scores && Object.keys(out.scores).length) {
    const table = document.createElement("table");
    table.className = "pf-scores";
    const tb = document.createElement("tbody");
    for (const [k, v] of Object.entries(out.scores)) {
      const tr = document.createElement("tr");
      const td1 = document.createElement("td");
      const td2 = document.createElement("td");
      td1.textContent = k;
      td2.textContent = String(v);
      tr.appendChild(td1);
      tr.appendChild(td2);
      tb.appendChild(tr);
    }
    table.appendChild(tb);
    card.appendChild(table);
  }

  if (out.global_feedback) {
    const p = document.createElement("p");
    p.className = "pf-global";
    p.textContent = out.global_feedback;
    card.appendChild(p);
  }

  container.appendChild(card);
}

function appendError(persona: Persona, err: Error) {
  const container = firstById<HTMLDivElement>(["results", "pfResults"]);
  if (!container) return;
  const card = document.createElement("div");
  card.className = "pf-card pf-error";
  card.textContent = `${persona.name}: ${err.message}`;
  container.appendChild(card);
}

// -----------------------------
// Export report (simple HTML print)
// -----------------------------
function exportReport() {
  const container = firstById<HTMLDivElement>(["results", "pfResults"]);
  if (!container) return window.print();

  const html = `
  <html>
    <head>
      <meta charset="utf-8"/>
      <title>Feedback Report</title>
      <style>
        body{font:14px system-ui, -apple-system, Segoe UI, Roboto, Arial}
        h1{font-size:20px;margin:0 0 12px}
        .card{border:1px solid #ddd;border-radius:8px;padding:12px;margin:12px 0}
        .scores td{padding:2px 6px}
      </style>
    </head>
    <body>
      <h1>Feedback Report</h1>
      ${container.innerHTML}
    </body>
  </html>`;

  const w = window.open("", "_blank");
  if (!w) return;
  w.document.write(html);
  w.document.close();
  w.focus();
  w.print();
}

// -----------------------------
// OPTIONAL: expose helpers for debugging
// -----------------------------
(Object.assign(window as any, {
  _pf: {
    loadSettings,
    saveSettings,
    clearAllComments,
    clearOurCommentsOnly,
    runAllEnabledPersonas
  }
}) as any);
