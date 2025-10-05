/* eslint-disable @typescript-eslint/no-explicit-any */
/* global Office, Word */
//
// Persona Feedback – Taskpane (TypeScript, all personas inlined)
// - Inlines DEFAULT_SETS (no external personas.ts)
// - Keeps proven behavior: matching-only comments, progress/debug, PDF export, settings, color legend
// - Avoids problematic APIs (no Word.HighlightColor, no Range.getSubstring)
//

/* ----------------------------- Utilities ----------------------------- */

type ProviderConfig = {
  provider: "openrouter" | "ollama";
  model: string;
  openrouterKey?: string;
};

type Persona = {
  id: string;
  enabled: boolean;
  name: string;
  system: string;
  instruction: string;
  color: string; // hex
};

type PersonaSet = {
  id: string;
  name: string;
  personas: Persona[];
};

type NormalizedLLM = {
  scores: { clarity: number; tone: number; alignment: number };
  global_feedback: string;
  comments: Array<{ quote: string; spanStart?: number; spanEnd?: number; comment: string }>;
};

type RunResult = {
  personaId: string;
  personaName: string;
  status: "done" | "error";
  scores?: NormalizedLLM["scores"];
  global_feedback?: string;
  comments?: NormalizedLLM["comments"];
  unmatched?: Array<{ quote: string; comment: string }>;
  raw?: any;
  error?: string;
};

const LS_KEY = "pf.settings.v1";

let SETTINGS: {
  provider: ProviderConfig;
  personaSetId: string;
  personaSets: PersonaSet[];
};

let LAST_RESULTS: RunResult[] = [];
let RUN_LOCK = false;

/* ----------------------------- DOM helpers ----------------------------- */

function byId<T extends HTMLElement = HTMLElement>(id: string) {
  return document.getElementById(id) as T | null;
}

function req<T extends HTMLElement = HTMLElement>(id: string) {
  const el = byId<T>(id);
  if (!el) throw new Error(`#${id} not found`);
  return el;
}

function toast(msg: string) {
  const t = byId("toast");
  const m = byId("toastMsg");
  if (!t || !m) return;
  m.textContent = msg;
  t.style.display = "block";
  setTimeout(() => (t.style.display = "none"), 1800);
}

function log(label: string, data?: any) {
  const pane = byId("debugLog");
  if (data !== undefined) console.log(label, data);
  else console.log(label);
  if (!pane) return;
  const div = document.createElement("div");
  div.style.whiteSpace = "pre-wrap";
  div.textContent = data !== undefined ? `${label} ${safeJson(data)}` : label;
  pane.appendChild(div);
  pane.scrollTop = pane.scrollHeight;
}

function safeJson(x: any) {
  try {
    return JSON.stringify(x, null, 2);
  } catch {
    return String(x);
  }
}

function escapeHtml(s: string) {
  return (s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function showView(id: "view-review" | "view-settings") {
  const r = byId("view-review");
  const s = byId("view-settings");
  const back = byId("btnBack") as HTMLButtonElement | null;
  if (id === "view-review") {
    r?.classList.remove("hidden");
    s?.classList.add("hidden");
    if (back) back.style.display = "none";
  } else {
    r?.classList.add("hidden");
    s?.classList.remove("hidden");
    if (back) back.style.display = "";
  }
}

function confirmAsync(title: string, message: string): Promise<boolean> {
  return new Promise((res) => {
    const overlay = req("confirmOverlay");
    req<HTMLDivElement>("confirmTitle").textContent = title;
    req<HTMLDivElement>("confirmMessage").textContent = message;
    overlay.style.display = "flex";
    const ok = req<HTMLButtonElement>("confirmOk");
    const cc = req<HTMLButtonElement>("confirmCancel");
    const done = (v: boolean) => {
      overlay.style.display = "none";
      ok.onclick = null;
      cc.onclick = null;
      res(v);
    };
    ok.onclick = () => done(true);
    cc.onclick = () => done(false);
  });
}

/* ----------------------------- Colors, Emojis ----------------------------- */

function hexToRgb(hex: string) {
  const m = hex.trim().replace("#", "");
  if (m.length !== 6) return { r: 200, g: 200, b: 200 };
  return {
    r: parseInt(m.slice(0, 2), 16),
    g: parseInt(m.slice(2, 4), 16),
    b: parseInt(m.slice(4, 6), 16),
  };
}
function rgbToHue({ r, g, b }: { r: number; g: number; b: number }) {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const d = max - min;
  let h = 0;
  if (d === 0) h = 0;
  else if (max === r) h = ((g - b) / d) % 6;
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  h = Math.round(h * 60);
  if (h < 0) h += 360;
  return h;
}
function hueToEmoji(h: number) {
  const pals = [
    { h: 0, e: "🟥" },
    { h: 30, e: "🟧" },
    { h: 60, e: "🟨" },
    { h: 120, e: "🟩" },
    { h: 210, e: "🟦" },
    { h: 280, e: "🟪" },
  ];
  let best = pals[0];
  let diff = 999;
  for (const p of pals) {
    const d = Math.min(Math.abs(h - p.h), 360 - Math.abs(h - p.h));
    if (d < diff) {
      diff = d;
      best = p;
    }
  }
  return best.e;
}
function colorEmojiFromHex(hex: string) {
  return hueToEmoji(rgbToHue(hexToRgb(hex || "#cccccc")));
}
function personaPrefix(p: Persona) {
  return `${colorEmojiFromHex(p.color)} [${p.name}]`;
}

/* ----------------------------- Personas (inlined) ----------------------------- */

function P(name: string, system: string, instruction: string, color: string): Persona {
  return {
    id: name.toLowerCase().replace(/[^a-z0-9]+/g, "-"),
    enabled: true,
    name,
    system,
    instruction,
    color,
  };
}

/** ALL persona sets merged here. */
const DEFAULT_SETS: PersonaSet[] = [
  // 1) Cross-Functional Team (original baseline)
  {
    id: "cross-functional-team",
    name: "Cross-Functional Team",
    personas: [
      P(
        "Senior Manager",
        "Senior manager prioritizing clarity, risk, outcomes.",
        "Assess clarity of goals, risks, outcomes; give concise suggestions.",
        "#fde047"
      ),
      P(
        "Legal",
        "Corporate counsel focused on compliance.",
        "Flag risky or ambiguous claims; suggest safer wording.",
        "#f9a8d4"
      ),
      P(
        "HR",
        "HR partner focused on inclusivity.",
        "Spot exclusionary tone; suggest inclusive language.",
        "#5eead4"
      ),
      P(
        "Technical Lead",
        "Engineering lead, pragmatic.",
        "Check feasibility, gaps, and technical risks.",
        "#93c5fd"
      ),
      P(
        "Junior Analyst",
        "Detail-oriented analyst.",
        "Call out unclear logic and missing data.",
        "#86efac"
      ),
    ],
  },

  // 2) Marketing Focus Group
  {
    id: "marketing-focus-group",
    name: "Marketing Focus Group",
    personas: [
      P(
        "Midwest Parent",
        "Pragmatic parent balancing budget and trust.",
        "React to clarity, trust, and family benefit; point out confusing or unconvincing claims.",
        "#f59e0b"
      ),
      P(
        "Gen-Z Student",
        "Digital native sensitive to authenticity.",
        "Flag cringe/marketing-speak; suggest authentic tone and concrete examples.",
        "#0ea5e9"
      ),
      P(
        "Retired Veteran",
        "Values respect, responsibility, and clarity.",
        "Request plain language; flag jargon; emphasize credibility.",
        "#6d28d9"
      ),
      P(
        "Small Business Owner",
        "ROI-driven decision maker.",
        "Ask for value proposition and costs; flag fluff.",
        "#16a34a"
      ),
      P(
        "Tech Pro",
        "Detail- and precision-oriented.",
        "Penalize vague claims; request specifications or metrics.",
        "#ef4444"
      ),
    ],
  },

  // 3) Startup Stakeholders
  {
    id: "startup-stakeholders",
    name: "Startup Stakeholders",
    personas: [
      P(
        "Founder",
        "Vision-first, resource constrained.",
        "Evaluate narrative coherence and focus; call out distractions.",
        "#22c55e"
      ),
      P(
        "CTO",
        "Architecture and feasibility oriented.",
        "Challenge technical assumptions and scalability; propose mitigations.",
        "#38bdf8"
      ),
      P(
        "CMO",
        "Go-to-market and positioning focused.",
        "Highlight clarity of ICP, messaging, and conversion path.",
        "#f97316"
      ),
      P(
        "VC Investor",
        "Skeptical; risk-reward framing.",
        "Probe moat, traction, unit economics; remove hand-wavy claims.",
        "#a855f7"
      ),
      P(
        "Customer",
        "Pragmatic buyer.",
        "Call out missing outcomes, ROI, integration blockers.",
        "#ef4444"
      ),
    ],
  },

  // 4) Political Spectrum (balanced tone-checkers)
  {
    id: "political-spectrum",
    name: "Political Spectrum",
    personas: [
      P(
        "Democratic Socialist",
        "Values equity and social safety nets.",
        "Assess alignment with labor rights, affordability, inclusion.",
        "#ef4444"
      ),
      P(
        "Center Left",
        "Pragmatic progressive.",
        "Seek workable policy framing and coalition-building language.",
        "#3b82f6"
      ),
      P(
        "Centrist / Independent",
        "Moderation, evidence, and trade-offs.",
        "Flag partisan tone; request balanced pros/cons.",
        "#64748b"
      ),
      P(
        "Center Right",
        "Market-oriented, institutions-focused.",
        "Request fiscal prudence and limited-overreach framing.",
        "#16a34a"
      ),
      P(
        "MAGA",
        "Populist, anti-elite rhetoric sensitivity.",
        "Flag technocratic tone; ask for concrete local benefits.",
        "#b91c1c"
      ),
      P(
        "Libertarian",
        "Individual liberty, minimal state.",
        "Flag mandates; ask for voluntary/adoption paths.",
        "#eab308"
      ),
    ],
  },

  // 5) Academic Peer Review
  {
    id: "academic-peer-review",
    name: "Academic Peer Review",
    personas: [
      P(
        "Methodologist",
        "Causal inference, validity, reproducibility.",
        "Probe assumptions, threats to validity, and confounds.",
        "#22c55e"
      ),
      P(
        "Statistician",
        "Rigor, power, error rates.",
        "Request effect sizes, CIs, and assumptions checks.",
        "#0ea5e9"
      ),
      P(
        "Domain Expert",
        "Contextual grounding and literature.",
        "Check citations, gaps vs prior work, practical implications.",
        "#a855f7"
      ),
      P(
        "Editor",
        "Structure, clarity, and ethics.",
        "Enforce clarity, concision, and ethical statements.",
        "#ef4444"
      ),
    ],
  },

  // 6) Enterprise Sales Cycle
  {
    id: "enterprise-sales",
    name: "Enterprise Sales Cycle",
    personas: [
      P(
        "Economic Buyer",
        "Budget authority, ROI-first.",
        "Ask for business case, payback, and KPI impact.",
        "#16a34a"
      ),
      P(
        "Champion",
        "Internal sponsor.",
        "Check enablement clarity and rollout path.",
        "#22c55e"
      ),
      P(
        "Security",
        "Risk, compliance, and audit.",
        "Probe data flows, encryption, vendor posture.",
        "#0ea5e9"
      ),
      P(
        "Procurement",
        "Price and T&Cs.",
        "Demand clarity on pricing, SLAs, and liabilities.",
        "#eab308"
      ),
      P(
        "End User",
        "Usability and workflow fit.",
        "Ask for friction points and training needs.",
        "#a855f7"
      ),
    ],
  },

  // 7) Nonprofit Board
  {
    id: "nonprofit-board",
    name: "Nonprofit Board",
    personas: [
      P(
        "Board Chair",
        "Governance and impact.",
        "Stress mission alignment and accountability.",
        "#0ea5e9"
      ),
      P(
        "Program Director",
        "Outcomes and beneficiaries.",
        "Ask for logic model and metrics.",
        "#16a34a"
      ),
      P(
        "Development Lead",
        "Donor narrative and compliance.",
        "Request compelling storytelling with transparency.",
        "#f97316"
      ),
      P(
        "Finance",
        "Budget, reserves, risk.",
        "Scrutinize sustainability and runway.",
        "#64748b"
      ),
    ],
  },

  // 8) Public Sector Review
  {
    id: "public-sector",
    name: "Public Sector Review",
    personas: [
      P(
        "Policy Analyst",
        "Feasibility and trade-offs.",
        "Request alternatives analysis and impacts.",
        "#3b82f6"
      ),
      P(
        "Regulator",
        "Statutory compliance.",
        "Flag conflicts with rules; require mitigation.",
        "#ef4444"
      ),
      P(
        "Civic Advocate",
        "Equity and access.",
        "Ask for outreach, language access, ADA compliance.",
        "#22c55e"
      ),
      P(
        "Budget Office",
        "Costs and sustainability.",
        "Demand operating and capital estimates.",
        "#eab308"
      ),
    ],
  },

  // 9) Product Trio
  {
    id: "product-trio",
    name: "Product Trio",
    personas: [
      P(
        "PM",
        "Outcomes and prioritization.",
        "Clarify problem, users, and success metrics.",
        "#22c55e"
      ),
      P(
        "Design",
        "Usability and accessibility.",
        "Flag confusing flows; suggest simplifications.",
        "#a855f7"
      ),
      P(
        "Engineering",
        "Feasibility and risks.",
        "Call out tech debt, scope creep, critical path.",
        "#0ea5e9"
      ),
    ],
  },

  // 10) Risk & Compliance
  {
    id: "risk-compliance",
    name: "Risk & Compliance",
    personas: [
      P(
        "InfoSec",
        "Threats, data handling, controls.",
        "Request threat model, encryption, access policies.",
        "#ef4444"
      ),
      P(
        "Privacy",
        "Data minimization and consent.",
        "Flag over-collection; require DSRs and retention.",
        "#22c55e"
      ),
      P(
        "Compliance",
        "Standards mapping.",
        "Ask for SOC2/ISO/NIST mappings and audit logs.",
        "#3b82f6"
      ),
    ],
  },

  // 11) Investor Panel
  {
    id: "investor-panel",
    name: "Investor Panel",
    personas: [
      P(
        "Seed Angel",
        "Team & velocity.",
        "Ask for insight loops, shipping cadence.",
        "#f59e0b"
      ),
      P(
        "Series A",
        "Product-market fit and growth.",
        "Probe retention, sales efficiency, expansion.",
        "#22c55e"
      ),
      P(
        "Growth Equity",
        "Scalability and unit economics.",
        "Scrutinize CAC/LTV, margin, capital plan.",
        "#0ea5e9"
      ),
    ],
  },

  // 12) Editorial Board
  {
    id: "editorial-board",
    name: "Editorial Board",
    personas: [
      P(
        "Copy Editor",
        "Grammar and readability.",
        "Enforce clear, concise, active voice.",
        "#a855f7"
      ),
      P(
        "Fact Checker",
        "Verifiability.",
        "Flag unsupported claims and missing citations.",
        "#ef4444"
      ),
      P(
        "Style Editor",
        "Tone and consistency.",
        "Unify terminology and voice guidelines.",
        "#3b82f6"
      ),
    ],
  },

  // 13) Customer Journeys
  {
    id: "customer-journeys",
    name: "Customer Journeys",
    personas: [
      P(
        "New Prospect",
        "First-touch clarity.",
        "Demand clear value proposition and next step.",
        "#22c55e"
      ),
      P(
        "Evaluator",
        "Comparative analysis.",
        "Request differentiators and proof points.",
        "#0ea5e9"
      ),
      P(
        "Champion",
        "Internal selling.",
        "Ask for ROI and rollout plan.",
        "#eab308"
      ),
      P(
        "Administrator",
        "Deployment burden.",
        "Probe SSO, provisioning, support model.",
        "#a855f7"
      ),
    ],
  },

  // 14) Accessibility & Inclusion
  {
    id: "a11y-inclusion",
    name: "Accessibility & Inclusion",
    personas: [
      P(
        "Screen Reader User",
        "Semantic clarity.",
        "Flag ambiguous headings, tables, images without alt text.",
        "#0ea5e9"
      ),
      P(
        "Low Vision",
        "Contrast and scale.",
        "Request high-contrast options and larger tap targets.",
        "#ef4444"
      ),
      P(
        "Neurodivergent",
        "Cognitive load.",
        "Ask for chunking, predictable patterns, reduced noise.",
        "#22c55e"
      ),
      P(
        "Non-native English",
        "Plain language.",
        "Replace idioms; prefer concrete examples.",
        "#eab308"
      ),
    ],
  },

  // 15) Internationalization
  {
    id: "internationalization",
    name: "Internationalization",
    personas: [
      P(
        "EU Market",
        "Data residency and privacy.",
        "Flag US-only assumptions; address GDPR nuances.",
        "#3b82f6"
      ),
      P(
        "APAC Market",
        "Localization and latency.",
        "Ask for translations, local partners, time-zone SLAs.",
        "#22c55e"
      ),
      P(
        "LATAM Market",
        "Payments and support.",
        "Request local payment rails and Spanish/Portuguese support.",
        "#f59e0b"
      ),
    ],
  },

  // 16) Security Review
  {
    id: "security-review",
    name: "Security Review",
    personas: [
      P(
        "AppSec",
        "Secure SDLC.",
        "Ask for threat modeling, code scanning, and secrets handling.",
        "#ef4444"
      ),
      P(
        "InfraSec",
        "Cloud posture.",
        "Probe network segmentation and least privilege.",
        "#0ea5e9"
      ),
      P(
        "Red Team",
        "Abuse cases.",
        "Challenge assumptions; propose controls.",
        "#a855f7"
      ),
    ],
  },

  // 17) Data Science Review
  {
    id: "data-science",
    name: "Data Science Review",
    personas: [
      P(
        "ML Engineer",
        "Production ML",
        "Ask for evaluation metrics, drift, and rollback.",
        "#22c55e"
      ),
      P(
        "Data Scientist",
        "Methodology & bias.",
        "Probe sampling, bias, and interpretability.",
        "#3b82f6"
      ),
      P(
        "Data Engineer",
        "Pipelines & quality.",
        "Demand lineage, quality checks, SLAs.",
        "#f59e0b"
      ),
    ],
  },

  // 18) UX Research Panel
  {
    id: "ux-research",
    name: "UX Research Panel",
    personas: [
      P(
        "New User",
        "Onboarding clarity.",
        "Ask what is confusing in first 5 minutes.",
        "#22c55e"
      ),
      P(
        "Power User",
        "Depth & efficiency.",
        "Push for shortcuts and batch flows.",
        "#0ea5e9"
      ),
      P(
        "Accessibility Advocate",
        "Barrier identification.",
        "Flag non-compliant patterns and alternatives.",
        "#ef4444"
      ),
    ],
  },

  // 19) Ops & Support
  {
    id: "ops-support",
    name: "Operations & Support",
    personas: [
      P(
        "Support Lead",
        "Case drivers.",
        "Predict top issues and KB needs.",
        "#a855f7"
      ),
      P(
        "SRE",
        "Reliability & incident response.",
        "Ask for SLOs and on-call playbooks.",
        "#0ea5e9"
      ),
      P(
        "QA",
        "Defect prevention.",
        "Request test plans and acceptance criteria.",
        "#22c55e"
      ),
    ],
  },

  // 20) Education Stakeholders
  {
    id: "education-stakeholders",
    name: "Education Stakeholders",
    personas: [
      P(
        "Teacher",
        "Classroom practicality.",
        "Ask for lesson alignment and time cost.",
        "#3b82f6"
      ),
      P(
        "Parent",
        "Safety & transparency.",
        "Probe privacy, grading fairness, support.",
        "#f59e0b"
      ),
      P(
        "Student",
        "Clarity & motivation.",
        "Flag confusing wording; suggest examples.",
        "#22c55e"
      ),
      P(
        "Administrator",
        "Policy & budget.",
        "Request compliance and rollout plans.",
        "#64748b"
      ),
    ],
  },
];

/* ----------------------------- Settings ----------------------------- */

function defaultSettings() {
  return {
    provider: { provider: "openrouter" as const, model: "openrouter/auto", openrouterKey: "" },
    personaSetId: DEFAULT_SETS[0].id,
    personaSets: DEFAULT_SETS,
  };
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return defaultSettings();
    const v = JSON.parse(raw);
    if (!v.personaSets?.length) v.personaSets = DEFAULT_SETS;
    if (!v.personaSetId) v.personaSetId = v.personaSets[0].id;
    return v;
  } catch {
    return defaultSettings();
  }
}

function saveSettings() {
  localStorage.setItem(LS_KEY, JSON.stringify(SETTINGS));
}

function currentSet(): PersonaSet {
  const id = SETTINGS.personaSetId;
  return SETTINGS.personaSets.find((s) => s.id === id) || SETTINGS.personaSets[0];
}

/* ----------------------------- UI wiring ----------------------------- */

window.addEventListener("error", (e) =>
  log(`[PF] window.error: ${e.message} @ ${e.filename || ""}:${e.lineno || ""}`)
);
window.addEventListener("unhandledrejection", (ev) =>
  log(`[PF] unhandledrejection: ${String((ev as any).reason)}`)
);

Office.onReady(async () => {
  document.body.style.minWidth = "500px";
  SETTINGS = loadSettings();

  populatePersonaSets();
  hydrateProviderUI();

  // Nav buttons
  const btnSettings = byId<HTMLButtonElement>("btnSettings");
  if (btnSettings) btnSettings.onclick = () => showView("view-settings");
  const btnBack = byId<HTMLButtonElement>("btnBack");
  if (btnBack) btnBack.onclick = () => showView("view-review");

  // Debug panel
  const toggleDebug = byId<HTMLButtonElement>("toggleDebug");
  if (toggleDebug)
    toggleDebug.onclick = () => {
      const p = byId("debugPanel");
      if (!p) return;
      p.classList.toggle("hidden");
      toggleDebug.textContent = p.classList.contains("hidden") ? "Show Debug" : "Hide Debug";
    };
  const clearDbg = byId<HTMLButtonElement>("clearDebug");
  if (clearDbg) clearDbg.onclick = () => {
    const p = byId("debugLog");
    if (p) p.innerHTML = "";
  };

  // Persona set selectors
  const sel = byId<HTMLSelectElement>("personaSet");
  if (sel)
    sel.onchange = (e: any) => {
      SETTINGS.personaSetId = e.target.value;
      saveSettings();
      populatePersonaSets();
    };
  const ssel = byId<HTMLSelectElement>("settingsPersonaSet");
  if (ssel)
    ssel.onchange = (e: any) => {
      SETTINGS.personaSetId = e.target.value;
      saveSettings();
      populatePersonaSets();
    };

  // Provider/model inputs
  const prv = byId<HTMLSelectElement>("provider");
  if (prv)
    prv.onchange = (e: any) => {
      SETTINGS.provider.provider = e.target.value as ProviderConfig["provider"];
      hydrateProviderUI();
      saveSettings();
    };
  const key = byId<HTMLInputElement>("openrouterKey");
  if (key)
    key.oninput = (e: any) => {
      SETTINGS.provider.openrouterKey = e.target.value;
      saveSettings();
    };
  const mdl = byId<HTMLInputElement>("model");
  if (mdl)
    mdl.oninput = (e: any) => {
      SETTINGS.provider.model = e.target.value;
      saveSettings();
    };

  // Settings save/restore
  const saveBtn = byId<HTMLButtonElement>("saveSettings");
  if (saveBtn)
    saveBtn.onclick = () => {
      const set = currentSet();
      set.personas.forEach((p, idx) => {
        const en = byId<HTMLInputElement>(`pe-enabled-${idx}`);
        const sys = byId<HTMLInputElement>(`pe-sys-${idx}`);
        const ins = byId<HTMLInputElement>(`pe-ins-${idx}`);
        const col = byId<HTMLInputElement>(`pe-color-${idx}`);
        if (en) p.enabled = !!en.checked;
        if (sys) p.system = sys.value;
        if (ins) p.instruction = ins.value;
        if (col) p.color = col.value || p.color;
      });
      saveSettings();
      renderPersonaNamesAndLegend();
      toast("Settings saved");
    };

  const restoreBtn = byId<HTMLButtonElement>("restoreDefaults");
  if (restoreBtn)
    restoreBtn.onclick = () => {
      const id = currentSet().id;
      const fresh = DEFAULT_SETS.find((s) => s.id === id);
      if (fresh) {
        const i = SETTINGS.personaSets.findIndex((s) => s.id === id);
        SETTINGS.personaSets[i] = JSON.parse(JSON.stringify(fresh));
        saveSettings();
        populatePersonaSets();
        toast("Default persona set restored");
      }
    };

  // Actions
  const runBtn = byId<HTMLButtonElement>("runBtn");
  if (runBtn) runBtn.onclick = handleRunReview;

  const retryBtn = byId<HTMLButtonElement>("retryBtn");
  if (retryBtn) retryBtn.onclick = handleRetryFailed;

  const exportBtn = byId<HTMLButtonElement>("exportBtn");
  if (exportBtn) exportBtn.onclick = handleExportPDF;

  const clearBtn = byId<HTMLButtonElement>("clearBtn");
  if (clearBtn)
    clearBtn.onclick = async () => {
      const ok = await confirmAsync("Clear all comments", "Delete ALL comments in this document?");
      if (!ok) return;
      const n = await clearAllComments();
      if (n >= 0) toast(n > 0 ? `Deleted ${n} comment(s).` : "No comments found.");
    };

  showView("view-review");
  log("[PF] Office.onReady → UI initialized");
});

/* ----------------------------- UI helpers ----------------------------- */

function populatePersonaSets() {
  const sets = SETTINGS.personaSets;
  const sel = byId<HTMLSelectElement>("personaSet");
  if (sel) {
    sel.innerHTML = "";
    sets.forEach((s) => {
      const o = document.createElement("option");
      o.value = s.id;
      o.textContent = s.name;
      sel.appendChild(o);
    });
    sel.value = SETTINGS.personaSetId;
  }
  const ssel = byId<HTMLSelectElement>("settingsPersonaSet");
  if (ssel) {
    ssel.innerHTML = "";
    sets.forEach((s) => {
      const o = document.createElement("option");
      o.value = s.id;
      o.textContent = s.name;
      ssel.appendChild(o);
    });
    ssel.value = SETTINGS.personaSetId;
  }
  renderPersonaNamesAndLegend();
  renderPersonaEditor();
}

function renderPersonaNamesAndLegend() {
  const set = currentSet();
  const names = byId("personaList");
  if (names) {
    names.textContent = set.personas
      .filter((p) => p.enabled)
      .map((p) => p.name)
      .join(", ");
  }
  const legend = byId("legend");
  if (legend) {
    legend.innerHTML = "";
    set.personas.forEach((p) => {
      const row = document.createElement("div");
      row.style.display = "flex";
      row.style.alignItems = "center";
      row.style.gap = "6px";
      const dot = document.createElement("span");
      dot.style.display = "inline-block";
      dot.style.width = "10px";
      dot.style.height = "10px";
      dot.style.borderRadius = "50%";
      dot.style.background = p.color;
      row.appendChild(dot);
      row.appendChild(document.createTextNode(p.name));
      legend.appendChild(row);
    });
  }
}

function renderPersonaEditor() {
  const set = currentSet();
  const box = byId("personaEditor");
  if (!box) return;
  box.innerHTML = "";
  set.personas.forEach((p, idx) => {
    const card = document.createElement("div");
    card.className = "card";
    card.style.marginBottom = "8px";
    card.innerHTML = `
      <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;">
        <div style="display:flex;gap:8px;align-items:center">
          <input type="checkbox" id="pe-enabled-${idx}" ${p.enabled ? "checked" : ""}/>
          <strong>${escapeHtml(p.name)}</strong>
        </div>
        <div style="display:flex;gap:6px;align-items:center">
          <label>Color</label><input id="pe-color-${idx}" type="color" value="${p.color}"/>
        </div>
      </div>
      <div class="row"><label>System</label><input id="pe-sys-${idx}" type="text" value="${escapeHtml(
        p.system
      )}"/></div>
      <div class="row"><label>Instruction</label><input id="pe-ins-${idx}" type="text" value="${escapeHtml(
        p.instruction
      )}"/></div>
    `;
    box.appendChild(card);
  });
}

function hydrateProviderUI() {
  const pr = byId<HTMLSelectElement>("provider");
  if (pr) pr.value = SETTINGS.provider.provider;
  const key = byId<HTMLInputElement>("openrouterKey");
  if (key) key.value = SETTINGS.provider.openrouterKey || "";
  const mdl = byId<HTMLInputElement>("model");
  if (mdl) mdl.value = SETTINGS.provider.model || "";
  const row = byId("openrouterKeyRow");
  if (row) row.classList.toggle("hidden", SETTINGS.provider.provider !== "openrouter");
}

/* ----------------------------- Progress, status ----------------------------- */

function setProgress(p: number) {
  const bar = byId<HTMLDivElement>("progBar");
  if (bar) bar.style.width = `${Math.max(0, Math.min(100, p))}%`;
}

function setBadgesHost(personas: Persona[]) {
  const host = byId("personaStatus");
  if (!host) return;
  host.innerHTML = "";
  personas.forEach((p) => {
    const row = document.createElement("div");
    row.style.display = "flex";
    row.style.justifyContent = "space-between";
    row.style.marginBottom = "4px";
    row.innerHTML = `<span style="display:inline-flex;align-items:center;gap:6px;">
        <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${p.color};"></span>
        ${escapeHtml(p.name)}
      </span>
      <span id="badge-${p.id}" class="badge">queued</span>`;
    host.appendChild(row);
  });
}

function setBadge(id: string, st: "queued" | "running" | "done" | "error", note?: string) {
  const b = byId("badge-" + id);
  if (!b) return;
  (b as HTMLElement).className =
    "badge " + (st === "done" ? "badge-done" : st === "error" ? "badge-failed" : "");
  (b as HTMLElement).textContent = st + (note ? ` – ${note}` : "");
}

/* ----------------------------- Run / Retry ----------------------------- */

async function handleRunReview() {
  if (RUN_LOCK) {
    toast("Already running…");
    return;
  }
  RUN_LOCK = true;
  try {
    LAST_RESULTS = [];
    const res = byId("results");
    if (res) res.innerHTML = "";
    const stat = byId("personaStatus");
    if (stat) stat.innerHTML = "";
    await runAllEnabledPersonas(false);
  } finally {
    RUN_LOCK = false;
  }
}

async function handleRetryFailed() {
  if (RUN_LOCK) {
    toast("Already running…");
    return;
  }
  RUN_LOCK = true;
  try {
    await runAllEnabledPersonas(true);
  } finally {
    RUN_LOCK = false;
  }
}

async function runAllEnabledPersonas(retryOnly: boolean) {
  const set = currentSet();
  const personas = set.personas.filter((p) => p.enabled);
  if (!personas.length) {
    toast("No personas enabled in this set.");
    return;
  }
  setBadgesHost(personas);
  const docText = await getWholeDocText();
  const total = personas.length;
  let done = 0;
  setProgress(0);

  for (const p of personas) {
    if (retryOnly) {
      const prev = LAST_RESULTS.find((r) => r.personaId === p.id);
      if (prev && prev.status === "done") {
        done++;
        setProgress((done / total) * 100);
        continue;
      }
    }
    setBadge(p.id, "running");
    try {
      const resp = await callLLMForPersona(p, docText);
      const normalized = normalizeResponse(resp);
      const { matched, unmatched } = await applyCommentsForMatchesOnly(p, normalized);
      addResultCard(p, normalized, unmatched);
      upsertResult({
        personaId: p.id,
        personaName: p.name,
        status: "done",
        scores: normalized.scores,
        global_feedback: normalized.global_feedback,
        comments: matched,
        unmatched,
        raw: resp,
      });
      setBadge(p.id, "done");
    } catch (err: any) {
      log(`[PF] Persona ${p.name} error`, err);
      upsertResult({
        personaId: p.id,
        personaName: p.name,
        status: "error",
        error: String((err && err.message) || err),
      });
      setBadge(p.id, "error", String((err && err.message) || "LLM call failed"));
    }
    done++;
    setProgress((done / total) * 100);
  }
  toast("Review finished.");
}

/* ----------------------------- Word helpers ----------------------------- */

async function getWholeDocText(): Promise<string> {
  return Word.run(async (ctx) => {
    const body = ctx.document.body;
    body.load("text");
    await ctx.sync();
    return body.text || "";
  });
}

async function clearAllComments(): Promise<number> {
  return Word.run(async (ctx) => {
    const coll = (ctx.document as any).comments;
    if (!coll || typeof coll.load !== "function") {
      toast("This Word build can’t list comments. Use Review → Delete → Delete All Comments.");
      return -1;
    }
    coll.load("items");
    await ctx.sync();
    let n = 0;
    for (const c of coll.items) {
      (c as any).delete();
      n++;
    }
    await ctx.sync();
    return n;
  });
}

function normalizeQuote(s: string) {
  return (s || "")
    .replace(/[\u2018\u2019\u201A\u201B]/g, "'")
    .replace(/[\u201C\u201D\u201E\u201F]/g, '"')
    .replace(/[\u2013\u2014]/g, "-")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function middleSlice(s: string, max: number) {
  if (s.length <= max) return s;
  const start = Math.max(0, Math.floor((s.length - max) / 2));
  return s.slice(start, start + max);
}

function seedFrom(s: string, which: "first" | "middle" | "last", words: number) {
  const t = s.split(/\s+/).filter(Boolean);
  if (t.length <= words) return s;
  if (which === "first") return t.slice(0, words).join(" ");
  if (which === "last") return t.slice(-words).join(" ");
  const mid = Math.floor(t.length / 2);
  const half = Math.floor(words / 2);
  return t.slice(Math.max(0, mid - half), Math.max(0, mid - half) + words).join(" ");
}

async function findRangeForQuote(ctx: Word.RequestContext, quote: string): Promise<Word.Range | null> {
  const body = ctx.document.body;
  let q = normalizeQuote(quote);
  if (q.length > 260) q = middleSlice(q, 180);

  const trySearch = async (needle: string) => {
    const r = body.search(needle, {
      matchCase: false,
      matchWholeWord: false,
      matchWildcards: false,
      ignoreSpace: true,
      ignorePunct: true,
    });
    r.load("items");
    await ctx.sync();
    return r.items.length ? r.items[0] : null;
  };

  // 1) Direct try
  let r = await trySearch(q);
  if (r) return r;

  // 2) Strip quotes if any
  const dq = q.replace(/^["'“”‘’]+/, "").replace(/["'“”‘’]+$/, "").trim();
  if (dq && dq !== q) {
    r = await trySearch(dq);
    if (r) return r;
  }

  // 3) Seeds (first/middle/last words)
  for (const w of [8, 6, 5]) {
    for (const pos of ["first", "middle", "last"] as const) {
      const s = seedFrom(q, pos, w);
      if (!s) continue;
      r = await trySearch(s);
      if (r) return r;
    }
  }
  return null;
}

async function applyCommentsForMatchesOnly(persona: Persona, data: NormalizedLLM) {
  const matched: NormalizedLLM["comments"] = [];
  const unmatched: Array<{ quote: string; comment: string }> = [];

  if (!Array.isArray(data.comments) || !data.comments.length) {
    return { matched, unmatched };
  }

  for (const [i, c] of data.comments.entries()) {
    const quote = String(c.quote || "").trim();
    const note = String(c.comment || "").trim();
    if (!quote || quote.length < 3) {
      log(`[PF] ${persona.name}: comment #${i + 1} empty/short`, c);
      continue;
    }
    const placed = await Word.run(async (ctx) => {
      const r = await findRangeForQuote(ctx, quote);
      if (!r) return false;
      const cm = (r as any).insertComment(`${personaPrefix(persona)} ${note}`);
      cm.load("id");
      await ctx.sync();
      return true;
    });

    if (placed) {
      matched.push({
        quote,
        spanStart: Number(c.spanStart || 0),
        spanEnd: Number(c.spanEnd || 0),
        comment: note,
      });
    } else {
      unmatched.push({ quote, comment: note });
    }
  }

  return { matched, unmatched };
}

/* ----------------------------- LLM calls ----------------------------- */

function withTimeout<T>(p: Promise<T>, ms = 45000): Promise<T> {
  return new Promise<T>((res, rej) => {
    const t = setTimeout(() => rej(new Error(`Request timed out after ${ms}ms`)), ms);
    p.then(
      (v) => {
        clearTimeout(t);
        res(v);
      },
      (e) => {
        clearTimeout(t);
        rej(e);
      }
    );
  });
}

async function fetchJson(url: string, init: RequestInit) {
  const r = await withTimeout(fetch(url, init));
  let body: any = null;
  let text = "";
  try {
    const ct = r.headers.get("content-type") || "";
    if (ct.includes("application/json")) body = await r.json();
    else {
      text = await r.text();
      try {
        body = JSON.parse(text);
      } catch {
        /* ignore */
      }
    }
  } catch (e) {
    try {
      text = await r.text();
    } catch {
      /* ignore */
    }
  }
  return { ok: (r as any).ok, status: (r as any).status, body, text };
}

async function callLLMForPersona(persona: Persona, docText: string): Promise<any> {
  const META_PROMPT = `
Return ONLY valid JSON matching this schema:
{
  "scores":{"clarity":0-100,"tone":0-100,"alignment":0-100},
  "global_feedback":"short paragraph",
  "comments":[{"quote":"verbatim snippet","spanStart":0,"spanEnd":0,"comment":"..."}]
}
Rules: No extra prose; if you output markdown, fence the JSON as \`\`\`json.
`.trim();

  const sys = `${persona.system}\n\n${META_PROMPT}`;
  const user = `You are acting as: ${persona.name}

INSTRUCTION:
${persona.instruction}

DOCUMENT (plain text):
${docText}`.trim();

  const pr = SETTINGS.provider;
  log(`[PF] Calling LLM → ${pr.provider} / ${pr.model} (${persona.name})`);

  if (pr.provider === "openrouter") {
    if (!pr.openrouterKey) throw new Error("Missing OpenRouter API key.");
    const res = await fetchJson("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${pr.openrouterKey}`,
        "HTTP-Referer": typeof window !== "undefined" ? window.location.origin : "https://example.com",
        "X-Title": "Persona Feedback Add-in",
      },
      body: JSON.stringify({
        model: pr.model || "openrouter/auto",
        messages: [
          { role: "system", content: sys },
          { role: "user", content: user },
        ],
        temperature: 0.2,
      }),
    });
    if (!res.ok) throw new Error(`OpenRouter HTTP ${res.status}: ${res.text || safeJson(res.body)}`);
    const content = res.body?.choices?.[0]?.message?.content ?? "";
    log("[PF] OpenRouter raw", res.body);
    return parseJsonFromText(content);
  } else {
    // Ollama local
    const res = await fetchJson("http://127.0.0.1:11434/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: pr.model || "llama3",
        stream: false,
        messages: [
          { role: "system", content: sys },
          { role: "user", content: user },
        ],
        options: { temperature: 0.2 },
      }),
    });
    if (!res.ok) throw new Error(`Ollama HTTP ${res.status}: ${res.text || safeJson(res.body)}`);
    const content = res.body?.message?.content ?? "";
    log("[PF] Ollama raw", res.body);
    return parseJsonFromText(content);
  }
}

function parseJsonFromText(text: string) {
  const m = text.match(/```json([\s\S]*?)```/i) || text.match(/```([\s\S]*?)```/);
  const raw = m ? m[1] : text;
  try {
    return JSON.parse(raw.trim());
  } catch {
    log("[PF] JSON parse error; full text follows", { text });
    throw new Error("Model returned non-JSON. See Debug for raw output.");
  }
}

function normalizeResponse(resp: any): NormalizedLLM {
  function clamp(n: any) {
    let v = Number(n || 0);
    if (!isFinite(v)) v = 0;
    return Math.max(0, Math.min(100, Math.round(v)));
  }
  return {
    scores: {
      clarity: clamp(resp?.scores?.clarity),
      tone: clamp(resp?.scores?.tone),
      alignment: clamp(resp?.scores?.alignment),
    },
    comments: Array.isArray(resp?.comments) ? resp.comments.slice(0, 12) : [],
    global_feedback: String(resp?.global_feedback || ""),
  };
}

/* ----------------------------- Results UI ----------------------------- */

function barColor(v: number) {
  if (v >= 80) return "#16a34a"; // green
  if (v < 50) return "#dc2626"; // red
  return "#eab308"; // yellow
}
function scoreBar(label: string, val: number) {
  const pct = Math.max(0, Math.min(100, val | 0));
  const c = barColor(pct);
  return `<div style="display:flex;justify-content:space-between;font-size:12px;margin-top:6px;"><span>${label}</span><span>${pct}</span></div>
          <div style="width:100%;height:14px;background:#e5e7eb;border-radius:999px;overflow:hidden;"><div style="height:100%;width:${pct}%;background:${c};"></div></div>`;
}

function addResultCard(persona: Persona, data: NormalizedLLM, unmatched: Array<{ quote: string; comment: string }>) {
  const host = byId("results");
  if (!host) return;
  const card = document.createElement("div");
  card.className = "card";
  card.style.marginBottom = "10px";
  const s = data.scores || { clarity: 0, tone: 0, alignment: 0 };
  card.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
      <div style="display:flex;align-items:center;gap:8px;">
        <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${persona.color};"></span>
        <strong>${escapeHtml(persona.name)}</strong>
      </div>
      <span class="badge badge-done">done</span>
    </div>
    ${scoreBar("Clarity", s.clarity)}${scoreBar("Tone", s.tone)}${scoreBar("Alignment", s.alignment)}
    <div style="margin-top:8px;"><em>${escapeHtml(data.global_feedback)}</em></div>
    ${
      unmatched && unmatched.length
        ? `
      <div style="margin-top:8px;">
        <div style="font-weight:600;margin-bottom:4px;">Unmatched quotes (not inserted):</div>
        <ul style="margin:0 0 0 16px;padding:0;list-style:disc;">
          ${(unmatched || [])
            .slice(0, 6)
            .map(
              (u) =>
                `<li><span style="color:#6b7280">"${escapeHtml(
                  u.quote.slice(0, 160)
                )}${u.quote.length > 160 ? "…" : ""}"</span><br/><span>${escapeHtml(u.comment)}</span></li>`
            )
            .join("")}
        </ul>
      </div>`
        : ""
    }
  `;
  host.appendChild(card);
}

/* ----------------------------- Report export (PDF via print) ----------------------------- */

function buildReportHtml() {
  const setName = currentSet().name;
  const rows = LAST_RESULTS.map((r) => {
    const persona = currentSet().personas.find((p) => p.id === r.personaId);
    const color = persona?.color || "#93c5fd";
    const s = r.scores || { clarity: 0, tone: 0, alignment: 0 };
    return `
      <section style="border:1px solid #e5e7eb;border-radius:10px;padding:14px;margin:10px 0;">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
          <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${color};"></span>
          <strong>${escapeHtml(r.personaName)}</strong>
          <span style="margin-left:auto;padding:2px 8px;border-radius:999px;background:#e5ffe8;border:1px solid #a7f3d0;">${r.status}</span>
        </div>
        ${scoreBar("Clarity", s.clarity)}${scoreBar("Tone", s.tone)}${scoreBar("Alignment", s.alignment)}
        ${r.global_feedback ? `<div style="margin-top:8px;"><em>${escapeHtml(r.global_feedback)}</em></div>` : ""}
      </section>`;
  }).join("");
  return `<!doctype html><html><head><meta charset="utf-8"><title>Persona Feedback Report</title>
  <style>body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif; padding:20px; max-width:800px; margin:0 auto;} h1{font-size:20px;margin:0 0 10px;} .muted{color:#6b7280}</style>
  </head><body>
    <h1>Persona Feedback Report</h1>
    <div class="muted">${new Date().toLocaleString()} • Set: ${escapeHtml(setName)}</div>
    ${rows || "<p class='muted'>No results.</p>"}
    <script>setTimeout(()=>window.print(), 300);</script>
  </body></html>`;
}

function handleExportPDF() {
  const html = buildReportHtml();
  const blob = new Blob([html], { type: "text/html" });
  const url = URL.createObjectURL(blob);
  window.open(url, "_blank");
}

/* ----------------------------- Results state ----------------------------- */

function upsert<T extends { personaId: string }>(arr: T[], item: T) {
  const i = arr.findIndex((x) => x.personaId === item.personaId);
  if (i >= 0) arr[i] = item;
  else arr.push(item);
}
function upsertResult(r: RunResult) {
  upsert(LAST_RESULTS, r);
}
