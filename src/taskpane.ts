/* global Office, Word */
import { DEFAULT_SETS, PersonaSet, Persona } from "./personas";

// ---------- Types ----------
type Settings = {
  provider: "openrouter" | "ollama";
  openrouterKey?: string;
  model: string;
  personaSets: Record<string, PersonaSet>;
  activeSetId: string;
};

type PersonaRunState = "queued" | "running" | "done" | "failed";
const STORAGE_KEY = "pf_settings_v2";

// ---------- Utils ----------
const clone = <T,>(o: T): T => {
  try { /* @ts-ignore */ if (typeof structuredClone === "function") return structuredClone(o); } catch {}
  return JSON.parse(JSON.stringify(o));
};

function debug(...args: any[]) {
  try {
    console.log("[PF]", ...args);
    const dbg = document.getElementById("debugLog") as HTMLDivElement | null;
    if (dbg) {
      const line = document.createElement("div");
      line.textContent = args.map(a => (typeof a === "string" ? a : JSON.stringify(a))).join(" ");
      dbg.appendChild(line);
      dbg.scrollTop = dbg.scrollHeight;
    }
  } catch {}
}

function showToast(msg: string) {
  const toast = document.getElementById("toast") as HTMLDivElement | null;
  const toastMsg = document.getElementById("toastMsg") as HTMLSpanElement | null;
  const toastClose = document.getElementById("toastClose") as HTMLSpanElement | null;
  if (!toast || !toastMsg || !toastClose) {
    debug("toast missing", msg);
    alert(msg); // fallback
    return;
  }
  toastMsg.textContent = msg;
  toast.style.display = "block";
  const hide = () => { toast.style.display = "none"; toastClose.removeEventListener("click", hide); };
  toastClose.addEventListener("click", hide);
  setTimeout(hide, 3000);

  if (/error|failed|missing|could not/i.test(msg)) {
    const dp = document.getElementById("debugPanel");
    const td = document.getElementById("toggleDebug");
    if (dp && td) { dp.classList.remove("hidden"); (td as HTMLButtonElement).textContent = "Hide Debug"; }
  }
}

// Mirror unexpected errors into debug
(function attachGlobalErrorHooks() {
  (window as any).addEventListener("error", (e: ErrorEvent) => {
    const msg = `window.error: ${e.message} @ ${e.filename}:${e.lineno}:${e.colno}`;
    console.error(msg, e.error);
    debug(msg);
  });
  (window as any).addEventListener("unhandledrejection", (e: PromiseRejectionEvent) => {
    const msg = `unhandledrejection: ${String(e.reason)}`;
    console.error(msg);
    debug(msg);
  });
})();

// ---------- State ----------
let settings: Settings;

// ---------- Persistence ----------
function loadSettings(): Settings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch {}
  const personaSets: Record<string, PersonaSet> = {};
  DEFAULT_SETS.forEach(s => { personaSets[s.id] = clone(s); });
  return {
    provider: "openrouter",
    openrouterKey: "",
    model: "openrouter/auto",
    personaSets,
    activeSetId: DEFAULT_SETS[0].id,
  };
}
function saveSettings() { localStorage.setItem(STORAGE_KEY, JSON.stringify(settings)); }

// ---------- Render ----------
function switchView(view: "review" | "settings") {
  const review = document.getElementById("view-review");
  const settingsView = document.getElementById("view-settings");
  if (!review || !settingsView) { debug("switchView missing containers"); return; }
  review.classList.toggle("hidden", view !== "review");
  settingsView.classList.toggle("hidden", view !== "settings");
}

function renderPersonaSetSelectors() {
  const personaSet = document.getElementById("personaSet") as HTMLSelectElement | null;
  const personaList = document.getElementById("personaList") as HTMLDivElement | null;
  if (!personaSet || !personaList) { debug("renderPersonaSetSelectors missing elements"); return; }

  const sets = Object.values(settings.personaSets);
  personaSet.innerHTML = sets.map(s => `<option value="${s.id}">${s.name}</option>`).join("");
  personaSet.value = settings.activeSetId;

  const names = sets.find(s => s.id === settings.activeSetId)?.personas
    .filter(p => p.enabled)
    .map(p => p.name) || [];
  personaList.textContent = names.join(" • ");
}

function renderSettingsForm() {
  const provider = document.getElementById("provider") as HTMLSelectElement | null;
  const openrouterKeyRow = document.getElementById("openrouterKeyRow") as HTMLDivElement | null;
  const openrouterKey = document.getElementById("openrouterKey") as HTMLInputElement | null;
  const model = document.getElementById("model") as HTMLInputElement | null;
  const settingsPersonaSet = document.getElementById("settingsPersonaSet") as HTMLSelectElement | null;

  if (!provider || !openrouterKeyRow || !openrouterKey || !model || !settingsPersonaSet) {
    debug("renderSettingsForm missing elements");
    return;
  }

  provider.value = settings.provider;
  openrouterKey.value = settings.openrouterKey || "";
  model.value = settings.model;
  openrouterKeyRow.style.display = settings.provider === "openrouter" ? "block" : "none";

  const sets = Object.values(settings.personaSets);
  settingsPersonaSet.innerHTML = sets.map(s => `<option value="${s.id}">${s.name}</option>`).join("");
  settingsPersonaSet.value = settings.activeSetId;

  renderPersonaEditor();
}

function renderPersonaEditor() {
  const settingsPersonaSet = document.getElementById("settingsPersonaSet") as HTMLSelectElement | null;
  const personaEditor = document.getElementById("personaEditor") as HTMLDivElement | null;
  if (!settingsPersonaSet || !personaEditor) { debug("renderPersonaEditor missing elements"); return; }
  const set = settings.personaSets[settingsPersonaSet.value];
  if (!set) { debug("renderPersonaEditor no set", settingsPersonaSet.value); return; }

  personaEditor.innerHTML = set.personas.map(p => `
    <div class="section">
      <label>
        <input type="checkbox" data-id="${p.id}" class="pe-enabled" ${p.enabled ? "checked" : ""}/>
        Enabled — <strong>${p.name}</strong>
      </label>
      <div class="row">
        <label>Name</label>
        <input type="text" class="pe-name" data-id="${p.id}" value="${p.name.replace(/"/g, "&quot;")}"/>
      </div>
      <div class="row">
        <label>System Prompt</label>
        <input type="text" class="pe-system" data-id="${p.id}" value="${p.system.replace(/"/g, "&quot;")}"/>
      </div>
      <div class="row">
        <label>Instruction Prompt</label>
        <input type="text" class="pe-instruction" data-id="${p.id}" value="${p.instruction.replace(/"/g, "&quot;")}"/>
      </div>
    </div>
  `).join("");

  // Wire inputs
  personaEditor.querySelectorAll<HTMLInputElement>(".pe-enabled").forEach(inp => {
    inp.onchange = () => {
      const p = set.personas.find(x => x.id === inp.dataset.id)!;
      p.enabled = inp.checked; saveSettings(); renderPersonaSetSelectors();
    };
  });
  personaEditor.querySelectorAll<HTMLInputElement>(".pe-name").forEach(inp => {
    inp.oninput = () => { const p = set.personas.find(x => x.id === inp.dataset.id)!; p.name = inp.value; saveSettings(); renderPersonaSetSelectors(); };
  });
  personaEditor.querySelectorAll<HTMLInputElement>(".pe-system").forEach(inp => {
    inp.oninput = () => { const p = set.personas.find(x => x.id === inp.dataset.id)!; p.system = inp.value; saveSettings(); };
  });
  personaEditor.querySelectorAll<HTMLInputElement>(".pe-instruction").forEach(inp => {
    inp.oninput = () => { const p = set.personas.find(x => x.id === inp.dataset.id)!; p.instruction = inp.value; saveSettings(); };
  });
}

function renderStatuses(status: Record<string, PersonaRunState>) {
  const personaStatus = document.getElementById("personaStatus") as HTMLDivElement | null;
  const progBar = document.getElementById("progBar") as HTMLDivElement | null;
  if (!personaStatus || !progBar) { debug("renderStatuses missing elements"); return; }

  const set = settings.personaSets[settings.activeSetId];
  const enabled = set.personas.filter(p => p.enabled);
  personaStatus.innerHTML = enabled.map(p => {
    const st = status[p.id] || "queued";
    const cls = st === "running" ? "running" : st === "done" ? "done" : st === "failed" ? "failed" : "queued";
    return `<div class="row"><span class="chip ${cls}">${p.name}: ${st}</span></div>`;
  }).join("");

  const total = enabled.length;
  const done = Object.values(status).filter(s => s === "done").length;
  progBar.style.width = total ? `${Math.floor((done / total) * 100)}%` : "0%";
}

function renderResultsView(results: Record<string, any>) {
  const resultsEl = document.getElementById("results") as HTMLDivElement | null;
  if (!resultsEl) { debug("renderResultsView missing results element"); return; }
  const set = settings.personaSets[settings.activeSetId];
  resultsEl.innerHTML = set.personas.filter(p => p.enabled).map(p => {
    const r = results[p.id];
    if (!r) return `<div class="row"><strong>${p.name}</strong><div class="muted">No result.</div></div>`;
    const s = r.scores || {};
    const gf = (r.global_feedback || "").toString().replace(/\n/g, "<br/>");
    return `
      <div class="section">
        <strong>${p.name}</strong>
        <div class="muted">Clarity: ${s.clarity ?? "—"} | Tone: ${s.tone ?? "—"} | Alignment: ${s.alignment ?? "—"}</div>
        <div style="margin-top:6px;">${gf}</div>
      </div>
    `;
  }).join("");
}

// ---------- JSON + LLM ----------
function safeParseJSON(s: string) {
  try {
    if (!s || typeof s !== "string") return { _parse_error: "empty", _raw: s };
    const trimmed = s.trim().replace(/^```(json)?/i, "").replace(/```$/, "").trim();
    return JSON.parse(trimmed);
  } catch (e: any) {
    return { _parse_error: String(e), _raw: s };
  }
}

async function callLLM(persona: Persona, docText: string): Promise<any> {
  const mdl = (settings.model || "").trim().toLowerCase();
  if (mdl === "debug-stub") {
    debug("using debug-stub", { persona: persona.id });
    return {
      scores: { clarity: 82, tone: 76, alignment: 88 },
      global_feedback: `Stubbed feedback for ${persona.name}. Your pipeline is working.`,
      comments: [{ quote: docText.slice(0, 30), comment: "Example inline comment from stub." }]
    };
  }

  const metaPrompt = `
You are "${persona.name}". Return STRICT JSON matching this schema:

{
  "scores": { "clarity": number, "tone": number, "alignment": number },
  "global_feedback": string,
  "comments": [ { "quote": string, "comment": string } ]
}

Rules:
- numbers 0..100
- comments array may be empty
- ONLY JSON (no markdown fences, no extra text)`.trim();

  const messages = [
    { role: "system", content: persona.system },
    { role: "user", content: `${persona.instruction}\n\n---\nDOCUMENT:\n${docText}` },
    { role: "user", content: metaPrompt }
  ];

  if (settings.provider === "openrouter") {
    if (!settings.openrouterKey) throw new Error("OpenRouter API key missing. Add it in Settings → Model.");
    const url = "https://openrouter.ai/api/v1/chat/completions";
    const payload = { model: settings.model, messages, temperature: 0 };
    debug("openrouter fetch →", { url, payload });

    let resp: Response;
    try {
      resp = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${settings.openrouterKey}`,
          "HTTP-Referer": window.location.origin,
          "X-Title": "Persona Feedback Word Add-in"
        },
        body: JSON.stringify(payload)
      });
    } catch (e: any) {
      debug("openrouter network error", String(e));
      throw new Error("Network error calling OpenRouter.");
    }

    const raw = await resp.text();
    debug("openrouter raw", raw.slice(0, 500));
    if (!resp.ok) throw new Error(`OpenRouter HTTP ${resp.status}: ${raw.slice(0, 300)}`);

    let data: any; try { data = JSON.parse(raw); } catch { data = { _raw: raw }; }
    const txt = data?.choices?.[0]?.message?.content ?? "";
    return safeParseJSON(txt);
  }

  // Ollama
  const url = "http://localhost:11434/api/chat";
  const payload = { model: settings.model, messages };
  debug("ollama fetch →", { url, payload });

  let resp: Response;
  try {
    resp = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
  } catch (e: any) {
    debug("ollama network error", String(e));
    throw new Error("Network error calling Ollama. Is it running on port 11434?");
  }

  const raw = await resp.text();
  debug("ollama raw", raw.slice(0, 500));
  if (!resp.ok) throw new Error(`Ollama HTTP ${resp.status}: ${raw.slice(0, 300)}`);

  let json: any; try { json = JSON.parse(raw); } catch { json = { _raw: raw }; }
  const txt = json?.message?.content ?? json?.choices?.[0]?.message?.content ?? raw;
  return safeParseJSON(txt);
}

// ---------- Word comments ----------
async function insertComments(personaName: string, comments: { quote: string; comment: string }[]) {
  if (!comments || !comments.length) return;
  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    for (const c of comments) {
      if (!c.comment) continue;

      if (c.quote && c.quote.trim().length > 0) {
        const search = body.search(c.quote, { matchCase: false, matchWholeWord: false });
        search.load("items");
        await context.sync();
        if (search.items.length > 0) {
          (search.items[0] as any).insertComment(`${personaName}: ${c.comment}`);
          await context.sync();
          continue;
        }
      }
      const tail = body.getRange("End");
      (tail as any).insertComment(`${personaName}: ${c.comment}`);
      await context.sync();
    }
  });
}

// ---------- Run flow ----------
async function runReview() {
  try {
    const runBtn = document.getElementById("runBtn") as HTMLButtonElement | null;
    if (runBtn) runBtn.disabled = true;

    // doc text
    let docText = "";
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        docText = body.text || "";
      });
      debug("runReview: doc loaded", { length: docText.length });
    } catch (e: any) {
      debug("Word.run error", String(e));
      showToast("Could not read document text.");
      if (runBtn) runBtn.disabled = false;
      return;
    }

    const set = settings.personaSets[settings.activeSetId];
    const personas = set.personas.filter(p => p.enabled);
    if (!personas.length) { showToast("No personas enabled."); if (runBtn) runBtn.disabled = false; return; }

    const status: Record<string, PersonaRunState> = {};
    personas.forEach(p => status[p.id] = "queued");
    renderStatuses(status);

    const results: Record<string, any> = {};
    for (const p of personas) {
      try {
        status[p.id] = "running"; renderStatuses(status);
        debug("persona start", p.id);
        const json = await callLLM(p, docText);
        results[p.id] = json;
        debug("persona result", { id: p.id, parsed: json && !json._parse_error });

        if (json?._parse_error) debug("parse error", { id: p.id, err: json._parse_error });
        if (json?.comments?.length) { await insertComments(p.name, json.comments); debug("comments inserted", p.id); }

        status[p.id] = json?._parse_error ? "failed" : "done";
        renderStatuses(status);
        renderResultsView(results);
      } catch (err: any) {
        status[p.id] = "failed";
        renderStatuses(status);
        debug("persona failed", { id: p.id, error: String(err) });
      }
    }

    renderResultsView(results);
    showToast("Review complete.");
    if (runBtn) runBtn.disabled = false;
  } catch (e: any) {
    debug("runReview fatal", String(e));
    showToast("Run failed (see Debug).");
    const runBtn = document.getElementById("runBtn") as HTMLButtonElement | null;
    if (runBtn) runBtn.disabled = false;
  }
}

// ---------- Wiring ----------
function wireEvents() {
  const gear = document.getElementById("gear");
  const back = document.getElementById("backToReview");
  const personaSet = document.getElementById("personaSet") as HTMLSelectElement | null;
  const runBtn = document.getElementById("runBtn");
  const toggleDebug = document.getElementById("toggleDebug");
  const clearDebug = document.getElementById("clearDebug");

  if (gear) gear.addEventListener("click", () => { renderSettingsForm(); switchView("settings"); });
  if (back) back.addEventListener("click", () => { switchView("review"); });
  if (personaSet) personaSet.addEventListener("change", () => {
    settings.activeSetId = personaSet.value; saveSettings(); renderPersonaSetSelectors();
  });
  if (toggleDebug) {
    let dv = false;
    toggleDebug.addEventListener("click", () => {
      dv = !dv;
      const dp = document.getElementById("debugPanel");
      if (dp) dp.classList.toggle("hidden", !dv);
      (toggleDebug as HTMLButtonElement).textContent = dv ? "Hide Debug" : "Show Debug";
    });
  }
  if (clearDebug) clearDebug.addEventListener("click", () => {
    const dbg = document.getElementById("debugLog"); if (dbg) dbg.innerHTML = "";
  });
  if (runBtn) runBtn.addEventListener("click", () => { runReview(); });

  // Settings
  const provider = document.getElementById("provider") as HTMLSelectElement | null;
  const openrouterKeyRow = document.getElementById("openrouterKeyRow");
  const openrouterKey = document.getElementById("openrouterKey") as HTMLInputElement | null;
  const model = document.getElementById("model") as HTMLInputElement | null;
  const settingsPersonaSet = document.getElementById("settingsPersonaSet") as HTMLSelectElement | null;
  const saveSettingsBtn = document.getElementById("saveSettings");
  const restoreDefaultsBtn = document.getElementById("restoreDefaults");

  if (provider) provider.addEventListener("change", () => {
    settings.provider = provider.value as Settings["provider"];
    if (openrouterKeyRow) openrouterKeyRow.style.display = settings.provider === "openrouter" ? "block" : "none";
    saveSettings();
  });
  if (openrouterKey) openrouterKey.addEventListener("input", () => { settings.openrouterKey = openrouterKey.value; saveSettings(); });
  if (model) model.addEventListener("input", () => { settings.model = model.value; saveSettings(); });
  if (settingsPersonaSet) settingsPersonaSet.addEventListener("change", () => {
    settings.activeSetId = settingsPersonaSet.value; saveSettings(); renderPersonaEditor(); renderPersonaSetSelectors();
  });
  if (saveSettingsBtn) saveSettingsBtn.addEventListener("click", () => { saveSettings(); showToast("Settings saved"); });
  if (restoreDefaultsBtn) restoreDefaultsBtn.addEventListener("click", () => {
    const def = DEFAULT_SETS.find(s => s.id === settings.activeSetId)!;
    settings.personaSets[def.id] = clone(def);
    saveSettings(); renderPersonaEditor(); renderPersonaSetSelectors(); showToast("Restored defaults");
  });
}

// ---------- Boot (guarded) ----------
(function boot() {
  // If Office.js isn’t present (previewing in a normal browser), show a friendly message.
  if (typeof (window as any).Office === "undefined") {
    debug("Office.js not available — are you opening taskpane.html directly in a browser?");
    const warn = document.createElement("div");
    warn.style.background = "#fff7ed";
    warn.style.color = "#9a3412";
    warn.style.padding = "8px";
    warn.style.border = "1px solid #fed7aa";
    warn.style.borderRadius = "8px";
    warn.style.marginTop = "8px";
    warn.textContent = "This page is intended to run inside Microsoft Word as an Office Add-in. Install the manifest and open from Word → Home → Persona Feedback.";
    document.body.prepend(warn);
    return;
  }

  (window as any).Office.onReady()
    .then(() => {
      try {
        settings = loadSettings();

        // DOM sanity check (helps diagnose mismatched HTML)
        const requiredIds = [
          "view-review","view-settings","gear","backToReview","personaSet","personaList","runBtn",
          "results","personaStatus","progBar","toggleDebug","debugPanel","debugLog","clearDebug",
          "provider","openrouterKeyRow","openrouterKey","model","settingsPersonaSet","personaEditor",
          "saveSettings","restoreDefaults","toast","toastMsg","toastClose"
        ];
        const missing = requiredIds.filter(id => !document.getElementById(id));
        if (missing.length) {
          debug("Missing DOM ids:", missing);
          const warn = document.createElement("div");
          warn.style.background = "#fff7ed";
          warn.style.color = "#9a3412";
          warn.style.padding = "8px";
          warn.style.border = "1px solid #fed7aa";
          warn.style.borderRadius = "8px";
          warn.style.marginTop = "8px";
          warn.textContent = "UI mismatch detected. Please replace public/taskpane.html with the provided version and reload.";
          document.body.prepend(warn);
        }

        wireEvents();
        renderPersonaSetSelectors();
        renderResultsView({});
        switchView("review");
        debug("Office.onReady → UI initialized");
      } catch (e: any) {
        debug("init error", String(e));
        showToast("Init failed (see Debug).");
      }
    })
    .catch((e: any) => {
      debug("Office.onReady failed", String(e));
      showToast("Office not ready (see Debug).");
    });
})();
