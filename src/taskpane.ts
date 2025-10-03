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
const STORAGE_KEY = "pf_settings_v5"; // bump to force merge of new default sets

/* =========================
            UTILS
========================= */
const clone = <T,>(o: T): T => JSON.parse(JSON.stringify(o));

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

/* =========================
        COLOR / TAGGING
========================= */
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

function hexToRgb(hex?: string): [number, number, number] {
  if (!hex) return [0, 0, 0];
  const s = hex.replace("#", "");
  return [0, 2, 4].map((i) => parseInt(s.substring(i, i + 2), 16)) as any;
}
function hexToHighlightColor(hex?: string): HL {
  const [r, g, b] = hexToRgb(hex);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  if (r > 220 && g > 210 && b < 140) return "Yellow";
  if (r > 230 && g < 160 && b < 160) return "Pink";
  if (g > 210 && r < 150 && b < 150) return "BrightGreen";
  if (g > 180 && b > 180 && r < 120) return "Turquoise";
  if (r > 230 && g > 170 && b < 120) return "Orange";
  if (b > 200 && r < 120) return "Blue";
  if (r > 200 && g < 120 && b < 120) return "Red";
  if (max - min < 20 && max > 200) return "LightGray";
  return "NoColor";
}
function colorToEmoji(hex?: string): string {
  const [r, g, b] = hexToRgb(hex);
  const palette = [
    { e: "🟥", c: [220, 70, 70] },
    { e: "🟧", c: [240, 150, 50] },
    { e: "🟨", c: [240, 220, 80] },
    { e: "🟩", c: [60, 180, 90] },
    { e: "🟦", c: [70, 120, 230] },
    { e: "🟪", c: [150, 80, 220] },
    { e: "🟫", c: [140, 90, 60] },
    { e: "⬛", c: [30, 30, 30] },
    { e: "⬜", c: [230, 230, 230] },
  ];
  let best = palette[0],
    dmin = 1e9;
  for (const p of palette) {
    const d = (r - p.c[0]) ** 2 + (g - p.c[1]) ** 2 + (b - p.c[2]) ** 2;
    if (d < dmin) {
      dmin = d;
      best = p;
    }
  }
  return best.e;
}
const PF_TAG_PREFIX = (p: Persona) => `${colorToEmoji(p.color)} ${p.name}: `;

/* =========================
           STATE
========================= */
let settings: Settings;
let lastRunResults: Record<string, any> = {};
let lastStatus: Record<string, PersonaRunState> = {};
let lastDebugPerPersona: Record<string, { raw?: string; error?: string }> = {};

/* =========================
       PERSISTENCE
========================= */
function loadSettings(): Settings {
  let base: Settings | null = null;
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) base = JSON.parse(raw);
  } catch {}

  if (!base) {
    base = {
      provider: "openrouter",
      openrouterKey: "",
      model: "openrouter/auto",
      personaSets: {},
      activeSetId: "cross-functional",
    };
  }

  // merge in any missing default sets without clobbering edits
  const existing = new Set(Object.keys(base.personaSets || {}));
  for (const def of DEFAULT_SETS) {
    if (!existing.has(def.id)) {
      base.personaSets[def.id] = clone(def);
    }
  }
  if (!base.personaSets[base.activeSetId]) {
    base.activeSetId = DEFAULT_SETS[0].id;
  }
  return base;
}
function saveSettings() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

/* =========================
     SCORE NORMALIZATION
========================= */
/** Normalize scores to 0..100 if model returns 0..1 or 0..10. */
function normalizeScores(s: any) {
  if (!s) return s;
  const vals = ["clarity", "tone", "alignment"].map((k) => Number(s[k] ?? 0));
  const max = Math.max(...vals);
  const min = Math.min(...vals);
  let factor = 1;
  if (max <= 1 && min >= 0) factor = 100; // 0..1 -> percentage
  else if (max <= 10 && min >= 0) factor = 10; // 0..10 -> 0..100
  const out: any = {};
  ["clarity", "tone", "alignment"].forEach((k, i) => {
    const v = Math.round(Math.max(0, Math.min(100, vals[i] * factor)));
    out[k] = v;
  });
  return out;
}

/* =========================
           RENDER
========================= */
function switchView(view: "review" | "settings") {
  document.getElementById("view-review")?.classList.toggle("hidden", view !== "review");
  document.getElementById("view-settings")?.classList.toggle("hidden", view !== "settings");
}
function renderLegend() {
  const set = settings.personaSets[settings.activeSetId];
  const el = document.getElementById("legend")!;
  el.innerHTML = set.personas
    .filter((p) => p.enabled)
    .map(
      (p) =>
        `<span class="swatch"><span class="dot" style="background:${p.color ||
          "#e5e7eb"}"></span>${p.name}</span>`
    )
    .join("");
}
function renderPersonaSetSelectors() {
  const sel = document.getElementById("personaSet") as HTMLSelectElement;
  const list = document.getElementById("personaList") as HTMLDivElement;
  const sets = Object.values(settings.personaSets);
  sel.innerHTML = sets.map((s) => `<option value="${s.id}">${s.name}</option>`).join("");
  sel.value = settings.activeSetId;
  const names = settings.personaSets[settings.activeSetId].personas
    .filter((p) => p.enabled)
    .map((p) => p.name);
  list.textContent = names.join(" • ");
  renderLegend();
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
      const s = normalizeScores(r.scores || {});
      const gf = (r.global_feedback || "").toString().replace(/\n/g, "<br/>");
      const statusBadge =
        lastStatus[p.id] === "failed"
          ? `<span class="badge badge-failed">failed</span>`
          : `<span class="badge badge-done">done</span>`;
      const err =
        lastStatus[p.id] === "failed" && lastDebugPerPersona[p.id]?.error
          ? `<div class="muted" style="margin-top:6px"><strong>Error:</strong> ${lastDebugPerPersona[
              p.id
            ].error!.slice(0, 400)}</div>`
          : "";
      return `
      <div class="result-card">
        <div style="display:flex;align-items:center;gap:8px;justify-content:space-between;">
          <div style="display:flex;align-items:center;gap:8px;">
            <span class="dot" style="background:${p.color || "#e5e7eb"}"></span>
            <strong>${p.name}</strong>
          </div>
          ${statusBadge}
        </div>
        <div style="display:grid;grid-template-columns:110px 1fr;gap:8px; margin-top:6px;">
          <div class="muted">Clarity</div><div>${scoreBarHTML(s.clarity)}</div>
          <div class="muted">Tone</div><div>${scoreBarHTML(s.tone)}</div>
          <div class="muted">Alignment</div><div>${scoreBarHTML(s.alignment)}</div>
        </div>
        <div style="margin-top:8px;">${gf}</div>
        ${err}
        <details style="margin-top:6px;"><summary>Debug JSON</summary><pre style="white-space:pre-wrap">${JSON.stringify(
          lastRunResults[p.id] || {},
          null,
          2
        )}</pre></details>
      </div>
    `;
    })
    .join("");
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

async function callLLM(persona: Persona, docText: string): Promise<{ parsed: any; raw: string }> {
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

  // DEBUG stub
  if ((settings.model || "").toLowerCase() === "debug-stub") {
    const raw = JSON.stringify({
      scores: { clarity: 8.2, tone: 7.6, alignment: 8.8 }, // will normalize to 82/76/88
      global_feedback: `Stub feedback for ${persona.name}.`,
      comments: [{ quote: docText.slice(0, 60), comment: "Example inline comment from stub." }],
    });
    return { parsed: safeParseJSON(raw), raw };
  }

  // Providers
  if (settings.provider === "openrouter") {
    if (!settings.openrouterKey) throw new Error("OpenRouter API key missing. Add it in Settings → Model.");
    const url = "https://openrouter.ai/api/v1/chat/completions";
    const payload = { model: settings.model, messages, temperature: 0 };

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
      const raw = String(e);
      return { parsed: { _parse_error: "network", _raw: raw }, raw };
    }

    const raw = await resp.text();
    if (!resp.ok) return { parsed: { _parse_error: `HTTP ${resp.status}`, _raw: raw }, raw };

    let data: any;
    try {
      data = JSON.parse(raw);
    } catch {
      data = { _raw: raw };
    }
    const txt = data?.choices?.[0]?.message?.content ?? "";
    return { parsed: safeParseJSON(txt), raw };
  }

  // Ollama
  const url = "http://localhost:11434/api/chat";
  const payload = { model: settings.model, messages };

  let resp: Response;
  try {
    resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
  } catch (e: any) {
    const raw = String(e);
    return { parsed: { _parse_error: "network", _raw: raw }, raw };
  }

  const raw = await resp.text();
  if (!resp.ok) return { parsed: { _parse_error: `HTTP ${resp.status}`, _raw: raw }, raw };

  let json: any;
  try {
    json = JSON.parse(raw);
  } catch {
    json = { _raw: raw };
  }
  const txt = json?.message?.content ?? json?.choices?.[0]?.message?.content ?? raw;
  return { parsed: safeParseJSON(txt), raw };
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
  const tag = PF_TAG_PREFIX(persona);

  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    for (const c of comments) {
      const text = (c.comment || "").trim();
      if (!text) continue;

      let inserted = false;
      if (c.quote && c.quote.trim().length > 0) {
        const search = body.search(c.quote.trim(), { matchCase: false, matchWholeWord: false });
        search.load("items");
        await context.sync();

        if (search.items.length > 0) {
          const target = search.items[0];
          try {
            (target as any).font.highlightColor = hl;
          } catch {}
          (target as any).insertComment(tag + text);
          inserted = true;
        }
      }

      if (!inserted) {
        const tail = body.getRange("End");
        (tail as any).insertComment(tag + text);
        try {
          (tail as any).font.highlightColor = hl;
        } catch {}
      }

      await context.sync();
    }
  });
}

/** Best-effort: remove our highlights & comments (tagged with emoji+name).
 * If Word's typings change, we’re defensive with try/catch.
 */
async function clearPFComments({ all = false } = {}) {
  await Word.run(async (context) => {
    const doc = context.document;

    // Remove highlights we applied (NoColor reset). We search common highlight shades.
    const body = doc.body;

    // We can’t search by highlight color directly, so we scan for key persona quotes again would be heavy.
    // Simpler: remove highlight from entire document (safe visual reset for PF). This won’t alter content.
    const whole = body.getRange();
    (whole as any).font.highlightColor = "NoColor";
    // Try to delete comments
    try {
      // @ts-ignore – runtime may expose comments
      const comments = (doc as any).comments;
      if (comments) {
        comments.load("items");
        await context.sync();
        const pfPrefixes = Object.values(settings.personaSets[settings.activeSetId].personas).map((p) =>
          PF_TAG_PREFIX(p)
        );
        for (const c of comments.items || []) {
          let text = "";
          try {
            // Word comment range text:
            c.contentRange.load("text");
            await context.sync();
            text = c.contentRange.text || "";
          } catch {}
          const matches = pfPrefixes.some((pref) => text.startsWith(pref));
          if (all || matches) {
            c.delete();
          }
        }
      }
    } catch (e) {
      // If delete via comments API fails, we still cleared highlights.
      debug("clearPFComments warn:", String(e));
    }
    await context.sync();
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

async function runReview({ onlyFailed = false } = {}) {
  const runBtn = document.getElementById("runBtn") as HTMLButtonElement | null;
  const retryBtn = document.getElementById("retryBtn") as HTMLButtonElement | null;

  try {
    if (runBtn) runBtn.disabled = true;
    if (retryBtn) retryBtn.disabled = true;

    const docText = await getDocText();
    if (!docText) {
      showToast("Document appears empty.");
      return;
    }

    const set = settings.personaSets[settings.activeSetId];
    const personas = set.personas.filter((p) => p.enabled);
    if (!personas.length) {
      showToast("No personas enabled.");
      return;
    }

    const targets = onlyFailed
      ? personas.filter((p) => lastStatus[p.id] === "failed")
      : personas;

    if (!targets.length && onlyFailed) {
      showToast("No failed personas to retry.");
      return;
    }

    if (!onlyFailed) {
      lastRunResults = {};
      lastStatus = {};
      lastDebugPerPersona = {};
    }

    for (const p of targets) {
      lastStatus[p.id] = "running";
    }
    renderStatuses(lastStatus);

    for (const p of targets) {
      try {
        const { parsed, raw } = await callLLM(p, docText);
        lastDebugPerPersona[p.id] = { raw };

        if (parsed?._parse_error) {
          lastDebugPerPersona[p.id].error = parsed._parse_error;
          lastRunResults[p.id] = parsed; // show in debug JSON
          lastStatus[p.id] = "failed";
          continue;
        }

        // normalize scores to 0..100
        if (parsed?.scores) parsed.scores = normalizeScores(parsed.scores);

        lastRunResults[p.id] = parsed;

        if (parsed?.comments?.length) {
          await insertComments(p, parsed.comments);
        }

        lastStatus[p.id] = "done";
        renderStatuses(lastStatus);
        renderResultsView(lastRunResults);
      } catch (err: any) {
        lastStatus[p.id] = "failed";
        lastDebugPerPersona[p.id] = { error: String(err) };
        renderStatuses(lastStatus);
      }
    }

    renderResultsView(lastRunResults);
    showToast("Review complete.");
  } finally {
    if (runBtn) runBtn.disabled = false;
    if (retryBtn) retryBtn.disabled = false;
  }
}

/* =========================
          EXPORT
========================= */
function htmlEscape(s: string) {
  return s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]!));
}
function buildReportHTML(results: Record<string, any>): string {
  const set = settings.personaSets[settings.activeSetId];
  const personas = set.personas.filter((p) => p.enabled);
  const sections = personas
    .map((p) => {
      const r = results[p.id];
      const s = normalizeScores(r?.scores || {});
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
<html><head><meta charset="utf-8"/>
<title>Persona Feedback Report</title>
<style>
  body{font-family:Segoe UI,Roboto,Arial,sans-serif;max-width:900px;margin:24px auto;padding:0 16px;color:#111827}
  h1{font-size:22px;margin:0 0 12px}
  h2{font-size:16px;margin:0 0 8px;display:flex;align-items:center;gap:8px}
  .muted{color:#6b7280}
  .dot{width:10px;height:10px;border-radius:50%;border:1px solid #d1d5db;display:inline-block}
  .card{border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin-bottom:12px}
  .grid{display:grid;grid-template-columns:120px 1fr;gap:8px;margin:8px 0}
  .bar{height:8px;background:#e5e7eb;border-radius:4px;overflow:hidden}
  .bar>div{height:100%;background:#2563eb}
</style></head>
<body>
  <h1>Persona Feedback Report</h1>
  ${sections || `<div class="muted">No results yet.</div>`}
</body></html>`;
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
  const retryBtn = document.getElementById("retryBtn");
  const clearBtn = document.getElementById("clearBtn");
  const clearAllBtn = document.getElementById("clearAllBtn");
  const toggleDebug = document.getElementById("toggleDebug");
  const clearDebug = document.getElementById("clearDebug");
  const exportBtn = document.getElementById("exportBtn");

  gear?.addEventListener("click", () => {
    renderSettingsForm();
    switchView("settings");
  });
  back?.addEventListener("click", () => switchView("review"));
  personaSet?.addEventListener("change", () => {
    settings.activeSetId = (personaSet as HTMLSelectElement).value;
    saveSettings();
    renderPersonaSetSelectors();
  });
  runBtn?.addEventListener("click", () => runReview({ onlyFailed: false }));
  retryBtn?.addEventListener("click", () => runReview({ onlyFailed: true }));
  clearBtn?.addEventListener("click", async () => {
    await clearPFComments({ all: false });
    showToast("Cleared persona highlights & PF-tagged comments.");
  });
  clearAllBtn?.addEventListener("click", async () => {
    if (confirm("Delete ALL comments in this document? This cannot be undone.")) {
      await clearPFComments({ all: true });
      showToast("Cleared all comments & highlights.");
    }
  });
  toggleDebug?.addEventListener("click", () => {
    const dp = document.getElementById("debugPanel");
    const v = dp?.classList.toggle("hidden");
    (toggleDebug as HTMLButtonElement).textContent = v ? "Show Debug" : "Hide Debug";
  });
  clearDebug?.addEventListener("click", () => {
    const dbg = document.getElementById("debugLog");
    if (dbg) dbg.innerHTML = "";
  });
  exportBtn?.addEventListener("click", exportReport);

  // Settings form
  const provider = document.getElementById("provider") as HTMLSelectElement | null;
  const openrouterKeyRow = document.getElementById("openrouterKeyRow");
  const openrouterKey = document.getElementById("openrouterKey") as HTMLInputElement | null;
  const model = document.getElementById("model") as HTMLInputElement | null;
  const settingsPersonaSet = document.getElementById("settingsPersonaSet") as HTMLSelectElement | null;
  const saveSettingsBtn = document.getElementById("saveSettings");
  const restoreDefaultsBtn = document.getElementById("restoreDefaults");

  provider?.addEventListener("change", () => {
    settings.provider = provider.value as Settings["provider"];
    if (openrouterKeyRow)
      openrouterKeyRow.style.display = settings.provider === "openrouter" ? "block" : "none";
    saveSettings();
  });
  openrouterKey?.addEventListener("input", () => {
    settings.openrouterKey = openrouterKey.value;
    saveSettings();
  });
  model?.addEventListener("input", () => {
    settings.model = model.value;
    saveSettings();
  });
  settingsPersonaSet?.addEventListener("change", () => {
    settings.activeSetId = settingsPersonaSet.value;
    saveSettings();
    renderPersonaEditor();
    renderPersonaSetSelectors();
  });
  saveSettingsBtn?.addEventListener("click", () => {
    saveSettings();
    showToast("Settings saved");
  });
  restoreDefaultsBtn?.addEventListener("click", () => {
    const def = DEFAULT_SETS.find((s) => s.id === settings.activeSetId)!;
    settings.personaSets[def.id] = clone(def);
    saveSettings();
    renderPersonaEditor();
    renderPersonaSetSelectors();
    showToast("Restored defaults");
  });
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
