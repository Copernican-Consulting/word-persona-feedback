/* global Office, Word */
import { DEFAULT_SETS, PersonaSet, Persona } from "./personas";

/* =========================
      TYPES / CONSTANTS
========================= */
type Settings = {
  provider: "openrouter" | "ollama";
  openrouterKey?: string;
  model: string;
  personaSets: Record<string, PersonaSet>;
  activeSetId: string;
};

type PersonaRunState = "queued" | "running" | "done" | "failed";
const STORAGE_KEY = "pf_settings_v3";

/* =========================
            UTILS
========================= */
const clone = <T,>(o: T): T => {
  try {
    // @ts-ignore
    if (typeof structuredClone === "function") return structuredClone(o);
  } catch {}
  return JSON.parse(JSON.stringify(o));
};

function debug(...args: any[]) {
  try {
    console.log("[PF]", ...args);
    const dbg = document.getElementById("debugLog") as HTMLDivElement | null;
    if (dbg) {
      const line = document.createElement("div");
      line.textContent = args
        .map((a) => (typeof a === "string" ? a : JSON.stringify(a)))
        .join(" ");
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
    alert(msg);
    return;
  }
  toastMsg.textContent = msg;
  toast.style.display = "block";
  const hide = () => {
    toast.style.display = "none";
    toastClose.removeEventListener("click", hide);
  };
  toastClose.addEventListener("click", hide);
  setTimeout(hide, 2200);
}

/**
 * Word typings don’t export a RangeHighlightColor type.
 * highlightColor accepts string literals like "Yellow", "Pink", etc.
 */
type HL =
  | "NoColor"
  | "Yellow"
  | "Pink"
  | "BrightGreen"
  | "Turquoise"
  | "LightGray"
  | "Violet"
  | "DarkYellow"
  | "DarkBlue"
  | "DarkRed"
  | "Teal"
  | "Brown"
  | "DarkGreen"
  | "DarkTeal"
  | "Indigo"
  | "Orange"
  | "Blue"
  | "Red"
  | "Green"
  | "Black"
  | "Gray25"
  | "Gray50";

/** crude hex → nearest highlight color guess */
function hexToHighlightColor(hex?: string): HL {
  if (!hex) return "NoColor";
  const rgb = (h: string) => {
    const s = h.replace("#", "");
    return [0, 2, 4].map((i) => parseInt(s.substring(i, i + 2), 16));
  };
  const [r, g, b] = rgb(hex);
  // quick / forgiving buckets
  if (r > 230 && g > 200 && b < 120) return "Yellow";
  if (r > 235 && g < 170 && b > 200) return "Violet";
  if (r > 235 && g < 170 && b < 170) return "Pink";
  if (g > 220 && r < 160 && b < 160) return "BrightGreen";
  if (g > 180 && b > 180 && r < 120) return "Turquoise";
  if (r > 240 && g > 170 && b < 120) return "Orange";
  if (r < 140 && g > 170 && b > 220) return "Blue";
  if (r > 200 && g > 200 && b > 200) return "LightGray";
  return "NoColor";
}

/* =========================
           STATE
========================= */
let settings: Settings;

/* =========================
       PERSISTENCE
========================= */
function loadSettings(): Settings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch {}
  const personaSets: Record<string, PersonaSet> = {};
  DEFAULT_SETS.forEach((s) => {
    personaSets[s.id] = clone(s);
  });
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

/* =========================
           RENDER
========================= */
function switchView(view: "review" | "settings") {
  const review = document.getElementById("view-review");
  const settingsView = document.getElementById("view-settings");
  if (!review || !settingsView) return;
  review.classList.toggle("hidden", view !== "review");
  settingsView.classList.toggle("hidden", view !== "settings");
}

function renderLegend() {
  const legend = document.getElementById("legend")!;
  const set = settings.personaSets[settings.activeSetId];
  const enabled = set.personas.filter((p) => p.enabled);
  legend.innerHTML = enabled
    .map(
      (p) =>
        `<span class="swatch"><span class="dot" style="background:${
          p.color || "#e5e7eb"
        }"></span>${p.name}</span>`
    )
    .join("");
}

function renderPersonaSetSelectors() {
  const personaSet = document.getElementById("personaSet") as HTMLSelectElement | null;
  const personaList = document.getElementById("personaList") as HTMLDivElement | null;
  if (!personaSet || !personaList) return;

  const sets = Object.values(settings.personaSets);
  personaSet.innerHTML = sets.map((s) => `<option value="${s.id}">${s.name}</option>`).join("");
  personaSet.value = settings.activeSetId;

  const names =
    sets
      .find((s) => s.id === settings.activeSetId)
      ?.personas.filter((p) => p.enabled)
      .map((p) => p.name) || [];
  personaList.textContent = names.join(" • ");

  renderLegend();
}

function renderSettingsForm() {
  const provider = document.getElementById("provider") as HTMLSelectElement | null;
  const openrouterKeyRow = document.getElementById("openrouterKeyRow") as HTMLDivElement | null;
  const openrouterKey = document.getElementById("openrouterKey") as HTMLInputElement | null;
  const model = document.getElementById("model") as HTMLInputElement | null;
  const settingsPersonaSet = document.getElementById("settingsPersonaSet") as HTMLSelectElement | null;

  if (!provider || !openrouterKeyRow || !openrouterKey || !model || !settingsPersonaSet) return;

  provider.value = settings.provider;
  openrouterKey.value = settings.openrouterKey || "";
  model.value = settings.model;
  openrouterKeyRow.style.display = settings.provider === "openrouter" ? "block" : "none";

  const sets = Object.values(settings.personaSets);
  settingsPersonaSet.innerHTML = sets.map((s) => `<option value="${s.id}">${s.name}</option>`).join("");
  settingsPersonaSet.value = settings.activeSetId;

  renderPersonaEditor();
}

function personaRow(p: Persona) {
  const esc = (s: string) => s.replace(/"/g, "&quot;");
  return `
    <div class="section">
      <label>
        <input type="checkbox" data-id="${p.id}" class="pe-enabled" ${p.enabled ? "checked" : ""}/>
        Enabled — <strong>${p.name}</strong>
      </label>
      <div class="row">
        <label>Name</label>
        <input type="text" class="pe-name" data-id="${p.id}" value="${esc(p.name)}"/>
      </div>
      <div class="row">
        <label>Color</label>
        <input type="color" class="pe-color" data-id="${p.id}" value="${p.color || "#e5e7eb"}" />
      </div>
      <div class="row">
        <label>System Prompt</label>
        <input type="text" class="pe-system" data-id="${p.id}" value="${esc(p.system)}"/>
      </div>
      <div class="row">
        <label>Instruction Prompt</label>
        <input type="text" class="pe-instruction" data-id="${p.id}" value="${esc(p.instruction)}"/>
      </div>
    </div>
  `;
}

function renderPersonaEditor() {
  const settingsPersonaSet = document.getElementById("settingsPersonaSet") as HTMLSelectElement | null;
  const personaEditor = document.getElementById("personaEditor") as HTMLDivElement | null;
  if (!settingsPersonaSet || !personaEditor) return;

  const set = settings.personaSets[settingsPersonaSet.value];
  personaEditor.innerHTML = set.personas.map(personaRow).join("");

  // Wire inputs
  personaEditor.querySelectorAll<HTMLInputElement>(".pe-enabled").forEach((inp) => {
    inp.onchange = () => {
      const p = set.personas.find((x) => x.id === inp.dataset.id)!;
      p.enabled = inp.checked;
      saveSettings();
      renderPersonaSetSelectors();
    };
  });
  personaEditor.querySelectorAll<HTMLInputElement>(".pe-name").forEach((inp) => {
    inp.oninput = () => {
      const p = set.personas.find((x) => x.id === inp.dataset.id)!;
      p.name = inp.value;
      saveSettings();
      renderPersonaSetSelectors();
    };
  });
  personaEditor.querySelectorAll<HTMLInputElement>(".pe-color").forEach((inp) => {
    inp.oninput = () => {
      const p = set.personas.find((x) => x.id === inp.dataset.id)!;
      p.color = inp.value;
      saveSettings();
      renderLegend();
    };
  });
  personaEditor.querySelectorAll<HTMLInputElement>(".pe-system").forEach((inp) => {
    inp.oninput = () => {
      const p = set.personas.find((x) => x.id === inp.dataset.id)!;
      p.system = inp.value;
      saveSettings();
    };
  });
  personaEditor.querySelectorAll<HTMLInputElement>(".pe-instruction").forEach((inp) => {
    inp.oninput = () => {
      const p = set.personas.find((x) => x.id === inp.dataset.id)!;
      p.instruction = inp.value;
      saveSettings();
    };
  });
}

function renderStatuses(status: Record<string, PersonaRunState>) {
  const personaStatus = document.getElementById("personaStatus") as HTMLDivElement | null;
  const progBar = document.getElementById("progBar") as HTMLDivElement | null;
  if (!personaStatus || !progBar) return;

  const set = settings.personaSets[settings.activeSetId];
  const enabled = set.personas.filter((p) => p.enabled);
  personaStatus.innerHTML = enabled
    .map((p) => {
      const st = status[p.id] || "queued";
      const color = p.color || "#e5e7eb";
      const bg =
        st === "running" ? "#fff7ed" : st === "done" ? "#ecfdf5" : st === "failed" ? "#fef2f2" : "#eef2ff";
      const fg =
        st === "running" ? "#9a3412" : st === "done" ? "#065f46" : st === "failed" ? "#991b1b" : "#3730a3";
      return `
      <div class="row" style="display:flex;align-items:center;gap:8px;">
        <span class="dot" style="background:${color}"></span>
        <span style="background:${bg};color:${fg};padding:2px 6px;border-radius:10px;font-size:12px;">
          ${p.name}: ${st}
        </span>
      </div>`;
    })
    .join("");

  const total = enabled.length;
  const done = Object.values(status).filter((s) => s === "done").length;
  progBar.style.width = total ? `${Math.floor((done / total) * 100)}%` : "0%";
}

function scoreBarHTML(val?: number) {
  const v = Math.max(0, Math.min(100, Number(val ?? 0)));
  return `
    <div class="scorebar"><div class="scorebar-fill" style="width:${v}%"></div></div>
    <div class="muted" style="font-size:12px;margin-top:2px;">${v}/100</div>
  `;
}

function renderResultsView(results: Record<string, any>) {
  const resultsEl = document.getElementById("results") as HTMLDivElement | null;
  if (!resultsEl) return;
  const set = settings.personaSets[settings.activeSetId];
  resultsEl.innerHTML = set.personas
    .filter((p) => p.enabled)
    .map((p) => {
      const r = results[p.id];
      if (!r) {
        return `<div class="result-card"><strong>${p.name}</strong><div class="muted">No result.</div></div>`;
      }
      const s = r.scores || {};
      const gf = (r.global_feedback || "").toString().replace(/\n/g, "<br/>");
      return `
      <div class="result-card">
        <div style="display:flex;align-items:center;gap:8px;">
          <span class="dot" style="background:${p.color || "#e5e7eb"}"></span>
          <strong>${p.name}</strong>
        </div>
        <div style="display:grid;grid-template-columns:110px 1fr;gap:8px; margin-top:6px;">
          <div class="muted">Clarity</div><div>${scoreBarHTML(s.clarity)}</div>
          <div class="muted">Tone</div><div>${scoreBarHTML(s.tone)}</div>
          <div class="muted">Alignment</div><div>${scoreBarHTML(s.alignment)}</div>
        </div>
        <div style="margin-top:8px;">${gf}</div>
      </div>
    `;
    })
    .join("");
}

/* =========================
       JSON + LLM I/O
========================= */
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
      global_feedback: `Stub feedback for ${persona.name}.`,
      comments: [{ quote: docText.slice(0, 60), comment: "Example inline comment from stub." }],
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
    { role: "user", content: metaPrompt },
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
          Authorization: `Bearer ${settings.openrouterKey}`,
          "HTTP-Referer": window.location.origin,
          "X-Title": "Persona Feedback Word Add-in",
        },
        body: JSON.stringify(payload),
      });
    } catch (e: any) {
      debug("openrouter network error", String(e));
      throw new Error("Network error calling OpenRouter.");
    }

    const raw = await resp.text();
    debug("openrouter raw", raw.slice(0, 500));
    if (!resp.ok) throw new Error(`OpenRouter HTTP ${resp.status}: ${raw.slice(0, 300)}`);

    let data: any;
    try {
      data = JSON.parse(raw);
    } catch {
      data = { _raw: raw };
    }
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
      body: JSON.stringify(payload),
    });
  } catch (e: any) {
    debug("ollama network error", String(e));
    throw new Error("Network error calling Ollama. Is it running on port 11434?");
  }

  const raw = await resp.text();
  debug("ollama raw", raw.slice(0, 500));
  if (!resp.ok) throw new Error(`Ollama HTTP ${resp.status}: ${raw.slice(0, 300)}`);

  let json: any;
  try {
    json = JSON.parse(raw);
  } catch {
    json = { _raw: raw };
  }
  const txt = json?.message?.content ?? json?.choices?.[0]?.message?.content ?? raw;
  return safeParseJSON(txt);
}

/* =========================
       WORD COMMENTS
========================= */
async function insertComments(
  persona: Persona,
  comments: { quote: string; comment: string }[]
) {
  if (!comments || !comments.length) return;
  const hl: HL = hexToHighlightColor(persona.color);

  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    for (const c of comments) {
      if (!c.comment) continue;

      // Try to anchor to the quoted text if provided
      if (c.quote && c.quote.trim().length > 0) {
        const search = body.search(c.quote, { matchCase: false, matchWholeWord: false });
        search.load("items");
        await context.sync();

        if (search.items.length > 0) {
          const target = search.items[0];
          try {
            // highlight color expects a specific string literal at runtime
            (target as any).font.highlightColor = hl;
          } catch {}
          (target as any).insertComment(`${persona.name}: ${c.comment}`);
          await context.sync();
          continue;
        }
      }

      // Fallback: append at end of document
      const tail = body.getRange("End");
      (tail as any).insertComment(`${persona.name}: ${c.comment}`);
      try {
        (tail as any).font.highlightColor = hl;
      } catch {}
      await context.sync();
    }
  });
}

/* =========================
          RUN FLOW
========================= */
async function getDocText(): Promise<string> {
  let docText = "";
  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    docText = body.text || "";
  });
  return docText;
}

async function runReview() {
  const runBtn = document.getElementById("runBtn") as HTMLButtonElement | null;
  try {
    if (runBtn) runBtn.disabled = true;

    let docText = "";
    try {
      docText = await getDocText();
      debug("runReview: doc loaded", { length: docText.length });
    } catch (e: any) {
      debug("Word.run error", String(e));
      showToast("Could not read document text.");
      return;
    }

    const set = settings.personaSets[settings.activeSetId];
    const personas = set.personas.filter((p) => p.enabled);
    if (!personas.length) {
      showToast("No personas enabled.");
      return;
    }

    const status: Record<string, PersonaRunState> = {};
    personas.forEach((p) => (status[p.id] = "queued"));
    renderStatuses(status);

    const results: Record<string, any> = {};
    for (const p of personas) {
      try {
        status[p.id] = "running";
        renderStatuses(status);
        const json = await callLLM(p, docText);
        results[p.id] = json;
        debug("persona result", { id: p.id, parsed: json && !json._parse_error });

        if (json?.comments?.length) {
          await insertComments(p, json.comments);
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
    lastRunResults = results; // for export
    showToast("Review complete.");
  } finally {
    if (runBtn) runBtn.disabled = false;
  }
}

/* =========================
          EXPORT
========================= */
let lastRunResults: Record<string, any> = {};

function htmlEscape(s: string) {
  return s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]!));
}

function buildReportHTML(results: Record<string, any>): string {
  const set = settings.personaSets[settings.activeSetId];
  const personas = set.personas.filter((p) => p.enabled);

  const sections = personas
    .map((p) => {
      const r = results[p.id];
      const s = r?.scores || {};
      const gf = htmlEscape((r?.global_feedback || "").toString());
      const comments = (r?.comments || []) as Array<{ quote: string; comment: string }>;
      const commentHtml = comments.length
        ? `<ul>${comments
            .map(
              (c) =>
                `<li><em>${htmlEscape(c.quote || "")}</em><br/>${htmlEscape(c.comment || "")}</li>`
            )
            .join("")}</ul>`
        : `<div class="muted">No inline comments</div>`;

      return `
      <section class="card">
        <h2><span class="dot" style="background:${p.color || "#e5e7eb"}"></span> ${htmlEscape(
        p.name
      )}</h2>
        <div class="grid">
          <div>Clarity</div><div><div class="bar"><div style="width:${Number(
            s.clarity || 0
          )}%"></div></div><small>${Number(s.clarity || 0)}/100</small></div>
          <div>Tone</div><div><div class="bar"><div style="width:${Number(
            s.tone || 0
          )}%"></div></div><small>${Number(s.tone || 0)}/100</small></div>
          <div>Alignment</div><div><div class="bar"><div style="width:${Number(
            s.alignment || 0
          )}%"></div></div><small>${Number(s.alignment || 0)}/100</small></div>
        </div>
        <h3>Global Feedback</h3>
        <p>${gf.replace(/\n/g, "<br/>")}</p>
        <h3>Inline Comments</h3>
        ${commentHtml}
      </section>`;
    })
    .join("");

  return `<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<title>Persona Feedback Report</title>
<style>
  body{font-family:Segoe UI,Roboto,Arial,sans-serif;max-width:900px;margin:24px auto;padding:0 16px;color:#111827}
  h1{font-size:22px;margin:0 0 12px}
  h2{font-size:16px;margin:0 0 8px;display:flex;align-items:center;gap:8px}
  h3{font-size:14px;margin:12px 0 6px}
  .muted{color:#6b7280}
  .legend{display:flex;gap:8px;flex-wrap:wrap;margin:8px 0 16px}
  .swatch{display:inline-flex;align-items:center;gap:6px;padding:2px 6px;border:1px solid #e5e7eb;border-radius:999px;font-size:12px}
  .dot{width:10px;height:10px;border-radius:50%;border:1px solid #d1d5db;display:inline-block}
  .card{border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin-bottom:12px}
  .grid{display:grid;grid-template-columns:120px 1fr;gap:8px;margin:8px 0}
  .bar{height:8px;background:#e5e7eb;border-radius:4px;overflow:hidden}
  .bar>div{height:100%;background:#2563eb}
</style>
</head>
<body>
  <h1>Persona Feedback Report</h1>
  <div class="legend">
    ${personas
      .map(
        (p) =>
          `<span class="swatch"><span class="dot" style="background:${p.color ||
            "#e5e7eb"}"></span>${htmlEscape(p.name)}</span>`
      )
      .join("")}
  </div>
  ${sections || `<div class="muted">No results yet.</div>`}
</body>
</html>`;
}

function exportReport() {
  const html = buildReportHTML(lastRunResults);
  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  a.download = `persona-feedback-${stamp}.html`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* =========================
          WIRING
========================= */
function wireEvents() {
  const gear = document.getElementById("gear");
  const back = document.getElementById("backToReview");
  const personaSet = document.getElementById("personaSet") as HTMLSelectElement | null;
  const runBtn = document.getElementById("runBtn");
  const toggleDebug = document.getElementById("toggleDebug");
  const clearDebug = document.getElementById("clearDebug");
  const exportBtn = document.getElementById("exportBtn");

  if (gear) gear.addEventListener("click", () => {
    renderSettingsForm();
    switchView("settings");
  });
  if (back) back.addEventListener("click", () => {
    switchView("review");
  });
  if (personaSet)
    personaSet.addEventListener("change", () => {
      settings.activeSetId = personaSet.value;
      saveSettings();
      renderPersonaSetSelectors();
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
  if (clearDebug)
    clearDebug.addEventListener("click", () => {
      const dbg = document.getElementById("debugLog");
      if (dbg) dbg.innerHTML = "";
    });
  if (runBtn) runBtn.addEventListener("click", () => {
    runReview();
  });
  if (exportBtn) exportBtn.addEventListener("click", exportReport);

  // Settings
  const provider = document.getElementById("provider") as HTMLSelectElement | null;
  const openrouterKeyRow = document.getElementById("openrouterKeyRow");
  const openrouterKey = document.getElementById("openrouterKey") as HTMLInputElement | null;
  const model = document.getElementById("model") as HTMLInputElement | null;
  const settingsPersonaSet = document.getElementById("settingsPersonaSet") as HTMLSelectElement | null;
  const saveSettingsBtn = document.getElementById("saveSettings");
  const restoreDefaultsBtn = document.getElementById("restoreDefaults");

  if (provider)
    provider.addEventListener("change", () => {
      settings.provider = provider.value as Settings["provider"];
      if (openrouterKeyRow) openrouterKeyRow.style.display = settings.provider === "openrouter" ? "block" : "none";
      saveSettings();
    });
  if (openrouterKey)
    openrouterKey.addEventListener("input", () => {
      settings.openrouterKey = openrouterKey.value;
      saveSettings();
    });
  if (model)
    model.addEventListener("input", () => {
      settings.model = model.value;
      saveSettings();
    });
  if (settingsPersonaSet)
    settingsPersonaSet.addEventListener("change", () => {
      settings.activeSetId = settingsPersonaSet.value;
      saveSettings();
      renderPersonaEditor();
      renderPersonaSetSelectors();
    });
  if (saveSettingsBtn)
    saveSettingsBtn.addEventListener("click", () => {
      saveSettings();
      showToast("Settings saved");
    });
  if (restoreDefaultsBtn)
    restoreDefaultsBtn.addEventListener("click", () => {
      const def = DEFAULT_SETS.find((s) => s.id === settings.activeSetId)!;
      settings.personaSets[def.id] = clone(def);
      saveSettings();
      renderPersonaEditor();
      renderPersonaSetSelectors();
      showToast("Restored defaults");
    });
}

/* =========================
           BOOT
========================= */
(function boot() {
  if (typeof (window as any).Office === "undefined") {
    const warn = document.createElement("div");
    warn.style.background = "#fff7ed";
    warn.style.color = "#9a3412";
    warn.style.padding = "8px";
    warn.style.border = "1px solid #fed7aa";
    warn.style.borderRadius = "8px";
    warn.style.marginTop = "8px";
    warn.textContent = "Open this add-in from Word (Home → Persona Feedback).";
    document.body.prepend(warn);
    return;
  }

  (window as any).Office.onReady()
    .then(() => {
      settings = loadSettings();
      wireEvents();
      renderPersonaSetSelectors();
      renderResultsView({});
      switchView("review");
      debug("Office.onReady → UI initialized");
    })
    .catch((e: any) => {
      debug("Office.onReady failed", String(e));
      showToast("Office not ready (see Debug).");
    });
})();
