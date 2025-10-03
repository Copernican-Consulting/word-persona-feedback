/* global Office */
import { DEFAULT_SETS, PersonaSet, Persona } from "./personas";

type Settings = {
  provider: "openrouter" | "ollama";
  openrouterKey?: string;
  model: string;
  personaSets: Record<string, PersonaSet>; // user-editable copies
  activeSetId: string;
};

const STORAGE_KEY = "pf_settings_v2";

const els = {
  // views
  review: document.getElementById("view-review") as HTMLDivElement,
  settings: document.getElementById("view-settings") as HTMLDivElement,
  gear: document.getElementById("gear") as HTMLSpanElement,
  back: document.getElementById("backToReview") as HTMLSpanElement,

  // review
  personaSet: document.getElementById("personaSet") as HTMLSelectElement,
  personaList: document.getElementById("personaList") as HTMLDivElement,
  runBtn: document.getElementById("runBtn") as HTMLButtonElement,
  results: document.getElementById("results") as HTMLDivElement,
  personaStatus: document.getElementById("personaStatus") as HTMLDivElement,
  progBar: document.getElementById("progBar") as HTMLDivElement,

  // debug
  toggleDebug: document.getElementById("toggleDebug") as HTMLButtonElement,
  debugPanel: document.getElementById("debugPanel") as HTMLDivElement,
  debugLog: document.getElementById("debugLog") as HTMLDivElement,
  clearDebug: document.getElementById("clearDebug") as HTMLButtonElement,

  // settings
  provider: document.getElementById("provider") as HTMLSelectElement,
  openrouterKeyRow: document.getElementById("openrouterKeyRow") as HTMLDivElement,
  openrouterKey: document.getElementById("openrouterKey") as HTMLInputElement,
  model: document.getElementById("model") as HTMLInputElement,
  settingsPersonaSet: document.getElementById("settingsPersonaSet") as HTMLSelectElement,
  personaEditor: document.getElementById("personaEditor") as HTMLDivElement,
  saveSettings: document.getElementById("saveSettings") as HTMLButtonElement,
  restoreDefaults: document.getElementById("restoreDefaults") as HTMLButtonElement,

  // toast
  toast: document.getElementById("toast") as HTMLDivElement,
  toastMsg: document.getElementById("toastMsg") as HTMLSpanElement,
  toastClose: document.getElementById("toastClose") as HTMLSpanElement,
};

let settings: Settings;

// Utilities
function ok<T>(v: T) { return v; }
function showToast(msg: string) {
  els.toastMsg.textContent = msg;
  els.toast.style.display = "block";
  const hide = () => { els.toast.style.display = "none"; els.toastClose.removeEventListener("click", hide); };
  els.toastClose.addEventListener("click", hide);
  setTimeout(hide, 3000);
}
function debug(...args: any[]) {
  const line = document.createElement("div");
  line.textContent = args.map(a => (typeof a === "string" ? a : JSON.stringify(a))).join(" ");
  els.debugLog.appendChild(line);
  els.debugLog.scrollTop = els.debugLog.scrollHeight;
}

// Persist
function loadSettings(): Settings {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (raw) {
    try { return JSON.parse(raw); } catch {}
  }
  // seed defaults
  const personaSets: Record<string, PersonaSet> = {};
  DEFAULT_SETS.forEach(s => personaSets[s.id] = structuredClone(s));
  return {
    provider: "openrouter",
    model: "openrouter/auto",
    openrouterKey: "",
    personaSets,
    activeSetId: DEFAULT_SETS[0].id
  };
}

function saveSettings() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

// UI wiring
function switchView(view: "review" | "settings") {
  els.review.classList.toggle("hidden", view !== "review");
  els.settings.classList.toggle("hidden", view !== "settings");
}

function renderPersonaSetSelectors() {
  const sets = Object.values(settings.personaSets);
  // review selector (names only)
  els.personaSet.innerHTML = sets.map(s => `<option value="${s.id}">${s.name}</option>`).join("");
  els.personaSet.value = settings.activeSetId;

  // list names
  const names = sets.find(s => s.id === settings.activeSetId)?.personas.filter(p => p.enabled).map(p => p.name) || [];
  els.personaList.textContent = names.join(" • ");
}

function renderSettingsForm() {
  // provider/key/model
  els.provider.value = settings.provider;
  els.openrouterKey.value = settings.openrouterKey || "";
  els.model.value = settings.model;
  els.openrouterKeyRow.style.display = settings.provider === "openrouter" ? "block" : "none";

  // persona sets dropdown
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
      <label><input type="checkbox" data-id="${p.id}" class="pe-enabled" ${p.enabled ? "checked" : ""}/> Enabled — <strong>${p.name}</strong></label>
      <div class="row"><label>Name</label><input type="text" class="pe-name" data-id="${p.id}" value="${p.name}"/></div>
      <div class="row"><label>System Prompt</label><input type="text" class="pe-system" data-id="${p.id}" value="${p.system.replace(/"/g, "&quot;")}"/></div>
      <div class="row"><label>Instruction Prompt</label><input type="text" class="pe-instruction" data-id="${p.id}" value="${p.instruction.replace(/"/g, "&quot;")}"/></div>
    </div>
  `).join("");

  // wire inputs
  els.personaEditor.querySelectorAll<HTMLInputElement>(".pe-enabled").forEach(inp => {
    inp.onchange = () => {
      const p = set.personas.find(x => x.id === inp.dataset.id)!;
      p.enabled = inp.checked; saveSettings();
    };
  });
  els.personaEditor.querySelectorAll<HTMLInputElement>(".pe-name").forEach(inp => {
    inp.oninput = () => { const p = set.personas.find(x => x.id === inp.dataset.id)!; p.name = inp.value; saveSettings(); };
  });
  els.personaEditor.querySelectorAll<HTMLInputElement>(".pe-system").forEach(inp => {
    inp.oninput = () => { const p = set.personas.find(x => x.id === inp.dataset.id)!; p.system = inp.value; saveSettings(); };
  });
  els.personaEditor.querySelectorAll<HTMLInputElement>(".pe-instruction").forEach(inp => {
    inp.oninput = () => { const p = set.personas.find(x => x.id === inp.dataset.id)!; p.instruction = inp.value; saveSettings(); };
  });
}

// Status & results
type PersonaRunState = "queued" | "running" | "done" | "failed";
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
    if (!r) return `<div class="row"><strong>${p.name}</strong><div class="muted">No result.</div></div>`;
    const s = r.scores || {};
    return `
      <div class="section">
        <strong>${p.name}</strong>
        <div class="muted">Clarity: ${s.clarity ?? "—"} | Tone: ${s.tone ?? "—"} | Alignment: ${s.alignment ?? "—"}</div>
        <div style="margin-top:6px;">${(r.global_feedback || "").replace(/\n/g, "<br/>")}</div>
      </div>
    `;
  }).join("");
}

// LLM calls
async function callLLM(persona: Persona, docText: string): Promise<any> {
  const metaPrompt = `
You are "${persona.name}". Return STRICT JSON matching this schema:

{
  "scores": { "clarity": number, "tone": number, "alignment": number },
  "global_feedback": string,
  "comments": [
    { "quote": string, "comment": string }
  ]
}

Rules:
- numbers 0..100
- comments array may be empty
- DO NOT include markdown fences or prose, ONLY JSON`;

  const messages = [
    { role: "system", content: persona.system },
    { role: "user", content: `${persona.instruction}\n\n---\nDOCUMENT:\n${docText}` },
    { role: "user", content: metaPrompt }
  ];

  if (settings.provider === "openrouter") {
    const resp = await fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${settings.openrouterKey ?? ""}`,
        "HTTP-Referer": window.location.origin,
        "X-Title": "Persona Feedback Word Add-in"
      },
      body: JSON.stringify({ model: settings.model, messages, temperature: 0 })
    });
    const data = await resp.json();
    debug("openrouter", { persona: persona.id, status: resp.status, data });
    const txt = data?.choices?.[0]?.message?.content ?? "";
    return safeParseJSON(txt);
  } else {
    // Ollama
    const resp = await fetch("http://localhost:11434/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ model: settings.model, messages })
    });
    const data = await resp.json();
    debug("ollama", { persona: persona.id, status: resp.status, data });
    const txt = data?.message?.content ?? data?.choices?.[0]?.message?.content ?? "";
    return safeParseJSON(txt);
  }
}

function safeParseJSON(s: string) {
  try {
    // strip fences if any
    const trimmed = s.trim().replace(/^```(json)?/i, "").replace(/```$/, "").trim();
    return JSON.parse(trimmed);
  } catch (e) {
    return { _parse_error: String(e), _raw: s };
  }
}

// Word helpers: insert comments (best-effort)
async function insertComments(personaName: string, comments: { quote: string; comment: string }[]) {
  if (!comments || !comments.length) return;
  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    for (const c of comments) {
      if (!c.comment) continue;

      // If we have a quote, try to find it first.
      if (c.quote && c.quote.trim().length > 0) {
        const search = body.search(c.quote, { matchCase: false, matchWholeWord: false });
        search.load("items");
        await context.sync();

        if (search.items.length > 0) {
          // Insert the comment directly on the matched range.
          const targetRange = search.items[0];
          // Range has insertComment; for older typings, cast to any to avoid TS complaints.
          (targetRange as any).insertComment(`${personaName}: ${c.comment}`);
          await context.sync();
          continue;
        }
      }

      // Fallback: attach a comment at the end of the document
      const tail = body.getRange("End"); // Range
      (tail as any).insertComment(`${personaName}: ${c.comment}`);
      await context.sync();
    }
  });
}


// Run flow
async function runReview() {
  els.results.innerHTML = "";
  const set = settings.personaSets[settings.activeSetId];
  const personas = set.personas.filter(p => p.enabled);
  if (!personas.length) { showToast("No personas enabled."); return; }

  // doc text
  let docText = "";
  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    docText = body.text || "";
  });

  const status: Record<string, PersonaRunState> = {};
  personas.forEach(p => status[p.id] = "queued");
  renderStatuses(status);

  const results: Record<string, any> = {};

  for (const p of personas) {
    try {
      status[p.id] = "running"; renderStatuses(status);

      const json = await callLLM(p, docText);
      results[p.id] = json;

      if (json?.comments?.length) {
        await insertComments(p.name, json.comments);
      }

      status[p.id] = json?._parse_error ? "failed" : "done";
      renderStatuses(status);
    } catch (err: any) {
      status[p.id] = "failed";
      renderStatuses(status);
      debug("error", { persona: p.id, error: String(err) });
    }
    renderResultsView(results);
  }

  showToast("Review complete.");
}

// Wire events
function wireEvents() {
  els.gear.onclick = () => { renderSettingsForm(); switchView("settings"); };
  els.back.onclick = () => { switchView("review"); };

  els.provider.onchange = () => {
    settings.provider = els.provider.value as any;
    els.openrouterKeyRow.style.display = settings.provider === "openrouter" ? "block" : "none";
    saveSettings();
  };
  els.openrouterKey.oninput = () => { settings.openrouterKey = els.openrouterKey.value; saveSettings(); };
  els.model.oninput = () => { settings.model = els.model.value; saveSettings(); };

  els.settingsPersonaSet.onchange = () => { settings.activeSetId = els.settingsPersonaSet.value; saveSettings(); renderPersonaEditor(); renderPersonaSetSelectors(); };

  els.saveSettings.onclick = () => { saveSettings(); showToast("Settings saved"); };
  els.restoreDefaults.onclick = () => {
    const def = DEFAULT_SETS.find(s => s.id === settings.activeSetId)!;
    settings.personaSets[def.id] = structuredClone(def);
    saveSettings();
    renderPersonaEditor(); renderPersonaSetSelectors();
    showToast("Restored defaults");
  };

  els.personaSet.onchange = () => { settings.activeSetId = els.personaSet.value; saveSettings(); renderPersonaSetSelectors(); };

  els.runBtn.onclick = () => { runReview(); };

  // debug panel
  let debugVisible = false;
  els.toggleDebug.onclick = () => {
    debugVisible = !debugVisible;
    els.debugPanel.classList.toggle("hidden", !debugVisible);
    els.toggleDebug.textContent = debugVisible ? "Hide Debug" : "Show Debug";
  };
  els.clearDebug.onclick = () => { els.debugLog.innerHTML = ""; };
}

// Boot
(function init() {
  settings = loadSettings();
  wireEvents();
  renderPersonaSetSelectors();
  renderResultsView({});
  switchView("review");
})();
