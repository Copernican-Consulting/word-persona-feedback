/* global Office, Word */
import { DEFAULT_SETS, PersonaSet, Persona } from "./personas";

// ---------- Types & constants ----------
type Settings = {
  provider: "openrouter" | "ollama";
  openrouterKey?: string;
  model: string;
  personaSets: Record<string, PersonaSet>;
  activeSetId: string;
};

type PersonaRunState = "queued" | "running" | "done" | "failed";

const STORAGE_KEY = "pf_settings_v2";

// ---------- Small utils ----------
const clone = <T,>(o: T): T => {
  try { /* @ts-ignore */ if (typeof structuredClone === "function") return structuredClone(o); } catch {}
  return JSON.parse(JSON.stringify(o));
};

function showToast(msg: string) {
  els.toastMsg.textContent = msg;
  els.toast.style.display = "block";
  const hide = () => { els.toast.style.display = "none"; els.toastClose.removeEventListener("click", hide); };
  els.toastClose.addEventListener("click", hide);
  setTimeout(hide, 3000);

  // If it sounds like an error, auto-open debug panel
  if (/error|failed|missing|could not/i.test(msg)) {
    els.debugPanel.classList.remove("hidden");
    els.toggleDebug.textContent = "Hide Debug";
  }
}

function debug(...args: any[]) {
  try {
    console.log("[PF]", ...args);
    const line = document.createElement("div");
    line.textContent = args.map(a => (typeof a === "string" ? a : JSON.stringify(a))).join(" ");
    els.debugLog.appendChild(line);
    els.debugLog.scrollTop = els.debugLog.scrollHeight;
  } catch {}
}

// Mirror unexpected errors into the debug panel too
(function attachGlobalErrorHooks() {
  (window as any).addEventListener("error", (e: ErrorEvent) => {
    try {
      const msg = `window.error: ${e.message} @ ${e.filename}:${e.lineno}:${e.colno}`;
      console.error(msg, e.error);
      const d = document.createElement("div");
      d.textContent = msg;
      (document.getElementById("debugLog") as HTMLDivElement)?.appendChild(d);
    } catch {}
  });
  (window as any).addEventListener("unhandledrejection", (e: PromiseRejectionEvent) => {
    try {
      const msg = `unhandledrejection: ${String(e.reason)}`;
      console.error(msg);
      const d = document.createElement("div");
      d.textContent = msg;
      (document.getElementById("debugLog") as HTMLDivElement)?.appendChild(d);
    } catch {}
  });
})();

// ---------- DOM refs ----------
const els = {
  // Views
  review: document.getElementById("view-review") as HTMLDivElement,
  settings: document.getElementById("view-settings") as HTMLDivElement,
  gear: document.getElementById("gear") as HTMLSpanElement,
  back: document.getElementById("backToReview") as HTMLSpanElement,

  // Review
  personaSet: document.getElementById("personaSet") as HTMLSelectElement,
  personaList: document.getElementById("personaList") as HTMLDivElement,
  runBtn: document.getElementById("runBtn") as HTMLButtonElement,
  results: document.getElementById("results") as HTMLDivElement,
  personaStatus: document.getElementById("personaStatus") as HTMLDivElement,
  progBar: document.getElementById("progBar") as HTMLDivElement,

  // Debug
  toggleDebug: document.getElementById("toggleDebug") as HTMLButtonElement,
  debugPanel: document.getElementById("debugPanel") as HTMLDivElement,
  debugLog: document.getElementById("debugLog") as HTMLDivElement,
  clearDebug: document.getElementById("clearDebug") as HTMLButtonElement,

  // Settings
  provider: document.getElementById("provider") as HTMLSelectElement,
  openrouterKeyRow: document.getElementById("openrouterKeyRow") as HTMLDivElement,
  openrouterKey: document.getElementById("openrouterKey") as HTMLInputElement,
  model: document.getElementById("model") as HTMLInputElement,
  settingsPersonaSet: document.getElementById("settingsPersonaSet") as HTMLSelectElement,
  personaEditor: document.getElementById("personaEditor") as HTMLDivElement,
  saveSettings: document.getElementById("saveSettings") as HTMLButtonElement,
  restoreDefaults: document.getElementById("restoreDefaults") as HTMLButtonElement,

  // Toast
  toast: document.getElementById("toast") as HTMLDivElement,
  toastMsg: document.getElementById("toastMsg") as HTMLSpanElement,
  toastClose: document.getElementById("toastClose") as HTMLSpanElement,
};

let settings: Settings;

// ---------- Persistence ----------
function loadSettings(): Settings {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (raw) {
    try { return JSON.parse(raw); } catch { /* fallthrough */ }
  }
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

function saveSettings() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

// ---------- UI render ----------
function switchView(view: "review" | "settings") {
  els.review.classList.toggle("hidden", view !== "review");
  els.settings.classList.toggle("hidden", view !== "settings");
}

function renderPersonaSetSelectors() {
  const sets = Object.values(settings.personaSets);
  // Review selector
  els.personaSet.innerHTML = sets.map(s => `<option value="${s.id}">${s.name}</option>`).join("");
  els.personaSet.value = settings.activeSetId;

  // Names only on review page
  const names = sets.find(s => s.id === settings.activeSetId)?.personas
    .filter(p => p.enabled)
    .map(p => p.name) || [];
  els.personaList.textContent = names.join(" • ");
}

function renderSettingsForm() {
  // Provider/Model
  els.provider.value = settings.provider;
  els.openrouterKey.value = settings.openrouterKey || "";
  els.model.value = settings.model;
  els.openrouterKeyRow.style.display = settings.provider === "openrouter" ? "block" : "none";

  // Persona sets dropdown (settings)
  const sets = Object.values(settings.personaSets);
  els.settingsPersonaSet.innerHTML = sets.map(s => `<option value="${s.id}">${s.name}</option>`).join("");
  els.settingsPersonaSet.value = settings.activeSetId;

  renderPersonaEditor();
}

function renderPersonaEditor() {
  const set = settings.personaSets[els.settingsPersonaSet.value];
  if (!set) return;

  els.personaEditor.innerHTML = set.personas.map(p => `
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
  els.personaEditor.querySelectorAll<HTMLInputElement>(".pe-enabled").forEach(inp => {
    inp.onchange = () => {
      const p = set.personas.find(x => x.id === inp.dataset.id)!;
      p.enabled = inp.checked;
      saveSettings();
      renderPersonaSetSelectors();
    };
  });
  els.personaEditor.querySelectorAll<HTMLInputElement>(".pe-name").forEach(inp => {
    inp.oninput = () => { const p = set.personas.find(x => x.id === inp.dataset.id)!; p.name = inp.value; saveSettings(); renderPersonaSetSelectors(); };
  });
  els.personaEditor.querySelectorAll<HTMLInputElement>(".pe-system").forEach(inp => {
    inp.oninput = () => { const p = set.personas.find(x => x.id === inp.dataset.id)!; p.system = inp.value; saveSettings(); };
  });
  els.personaEditor.querySelectorAll<HTMLInputElement>(".pe-instruction").forEach(inp => {
    inp.oninput = () => { const p = set.personas.find(x => x.id === inp.dataset.id)!; p.instruction = inp.value; saveSettings(); };
  });
}

function renderStatuses(status: Record<string, PersonaRunState>) {
  const set = settings.personaSets[settings.activeSetId];
  const enabled = set.personas.filter(p => p.enabled);
  els.personaStatus.innerHTML = enabled.map(p => {
    const st = status[p.id] || "queued";
    const cls = st === "running" ? "running" : st === "done" ? "done" : st === "failed" ? "failed" : "queued";
    return `<div class="row"><span class="chip ${cls}">${p.name}: ${st}</span></div>`;
  }).join("");

  const total = enabled.length;
  const done = Object.values(status).filter(s => s === "done").length;
  els.progBar.style.width = total ? `${Math.floor((done / total) * 100)}%` : "0%";
}

function renderResultsView(results: Record<string, any>) {
  const set = settings.personaSets[settings.activeSetId];
  els.results.innerHTML = set.personas.filter(p => p.enabled).map(p => {
    const r = results[p.id];
    if (!r) {
      return `<div class="row"><strong>${p.name}</strong><div class="muted">No result.</div></div>`;
    }
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

// ---------- JSON helpers ----------
function safeParseJSON(s: string) {
  try {
    if (!s || typeof s !== "string") return { _parse_error: "empty", _raw: s };
    const trimmed = s.trim().replace(/^```(json)?/i, "").replace(/```$/, "").trim();
    return JSON.parse(trimmed);
  } catch (e: any) {
    return { _parse_error: String(e), _raw: s };
  }
}

// ---------- LLM calls ----------
async function callLLM(persona: Persona, docText: string): Promise<any> {
  // Debug stub model: end-to-end test without network
  if ((settings.model || "").trim().toLowerCase() === "debug-stub") {
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
    if (!settings.openrouterKey) {
      const err = "OpenRouter API key missing. Add it in Settings → Model.";
      debug("openrouter error", err);
      throw new Error(err);
    }

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

    let data: any;
    try { data = JSON.parse(raw); } catch { data = { _raw: raw }; }
    const txt = data?.choices?.[0]?.message?.content ?? "";
    return safeParseJSON(txt);
  }

  // Ollama
  const url = "http://localhost:11434/api/chat";
  const payload = { model: settings.model, messages };
  debug("ollama fetch →", { url, payload });

  let resp: Response;
  try {
    resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
  } catch (e: any) {
    debug("ollama network error", String(e));
    throw new Error("Network error calling Ollama. Is it running on port 11434?");
  }

  const raw = await resp.text();
  debug("ollama raw", raw.slice(0, 500));
  if (!resp.ok) throw new Error(`Ollama HTTP ${resp.status}: ${raw.slice(0, 300)}`);

  // Try to be liberal about response shapes
  let json: any;
  try { json = JSON.parse(raw); } catch { json = { _raw: raw }; }
  const txt = json?.message?.content ?? json?.choices?.[0]?.message?.content ?? raw;
  return safeParseJSON(txt);
}

// ---------- Word comment insertion ----------
async function insertComments(personaName: string, comments: { quote: string; comment: string }[]) {
  if (!comments || !comments.length) return;
  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    for (const c of comments) {
      if (!c.comment) continue;

      // If we have a quote, try to find it
      if (c.quote && c.quote.trim().length > 0) {
        const search = body.search(c.quote, { matchCase: false, matchWholeWord: false });
        search.load("items");
        await context.sync();

        if (search.items.length > 0) {
          const targetRange = search.items[0];
          (targetRange as any).insertComment(`${personaName}: ${c.comment}`);
          await context.sync();
          continue;
        }
      }

      // Fallback: comment at the end of the document
      const tail = body.getRange("End");
      (tail as any).insertComment(`${personaName}: ${c.comment}`);
      await context.sync();
    }
  });
}

// ---------- Run flow ----------
async function runReview() {
  try {
    els.results.innerHTML = "";
    els.runBtn.disabled = true;

    const set = settings.personaSets[settings.activeSetId];
    const personas = set.personas.filter(p => p.enabled);

    debug("runReview: start", { activeSetId: settings.activeSetId, personas: personas.map(p => p.id) });

    if (!personas.length) {
      showToast("No personas enabled.");
      els.runBtn.disabled = false;
      return;
    }

    // Read document text
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
      els.runBtn.disabled = false;
      return;
    }

    // Init status
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

        if (json?._parse_error) {
          debug("parse error", { id: p.id, err: json._parse_error });
        }

        if (json?.comments?.length) {
          await insertComments(p.name, json.comments);
          debug("comments inserted", p.id);
        }

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
  } catch (e: any) {
    debug("runReview fatal", String(e));
    showToast("Run failed (see Debug).");
  } finally {
    els.runBtn.disabled = false;
  }
}

// ---------- Events ----------
function wireEvents() {
  els.gear.onclick = () => { renderSettingsForm(); switchView("settings"); };
  els.back.onclick = () => { switchView("review"); };

  // Review selectors
  els.personaSet.onchange = () => {
    settings.activeSetId = els.personaSet.value;
    saveSettings();
    renderPersonaSetSelectors();
  };

  // Debug panel
  let debugVisible = false;
  els.toggleDebug.onclick = () => {
    debugVisible = !debugVisible;
    els.debugPanel.classList.toggle("hidden", !debugVisible);
    els.toggleDebug.textContent = debugVisible ? "Hide Debug" : "Show Debug";
  };
  els.clearDebug.onclick = () => { els.debugLog.innerHTML = ""; };

  els.runBtn.onclick = () => { runReview(); };

  // Settings panel
  els.provider.onchange = () => {
    settings.provider = els.provider.value as Settings["provider"];
    els.openrouterKeyRow.style.display = settings.provider === "openrouter" ? "block" : "none";
    saveSettings();
  };
  els.openrouterKey.oninput = () => { settings.openrouterKey = els.openrouterKey.value; saveSettings(); };
  els.model.oninput = () => { settings.model = els.model.value; saveSettings(); };

  els.settingsPersonaSet.onchange = () => {
    settings.activeSetId = els.settingsPersonaSet.value;
    saveSettings();
    renderPersonaEditor();
    renderPersonaSetSelectors();
  };

  els.saveSettings.onclick = () => { saveSettings(); showToast("Settings saved"); };
  els.restoreDefaults.onclick = () => {
    const def = DEFAULT_SETS.find(s => s.id === settings.activeSetId)!;
    settings.personaSets[def.id] = clone(def);
    saveSettings();
    renderPersonaEditor();
    renderPersonaSetSelectors();
    showToast("Restored defaults");
  };
}

// ---------- Boot after Office is ready ----------
Office.onReady()
  .then(() => {
    try {
      settings = loadSettings();
      wireEvents();
      renderPersonaSetSelectors();
      renderResultsView({});
      switchView("review");
      debug("Office.onReady → UI initialized");
    } catch (e) {
      debug("init error", String(e));
      showToast("Init failed (see Debug).");
    }
  })
  .catch((e) => {
    debug("Office.onReady failed", String(e));
    showToast("Office not ready (see Debug).");
  });
