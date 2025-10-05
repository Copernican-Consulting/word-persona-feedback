/* eslint-disable @typescript-eslint/no-explicit-any */

import "./ui.css";

/* ===============================
   Types
================================= */
type ChatMessage = { role: "system" | "user" | "assistant"; content: string };

type Persona = {
  id: string;
  name: string;
  color?: string;
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
  | { kind: "openrouter"; apiKey: string; model: string }
  | { kind: "ollama"; host?: string; model: string };

type AppSettings = {
  provider?: ProviderConfig;
  selectedSetId?: string;
  personaEnabled?: Record<string, boolean>;
};

/* ===============================
   Persona catalog (merged & embedded)
   13 sets total; includes "Tech Professional" in Marketing Focus
================================= */
const DEFAULT_SETS: PersonaSet[] = [
  {
    id: "cross-functional",
    name: "Cross-Functional Team",
    personas: [
      {
        id: "senior-manager",
        name: "Senior Manager",
        color: "#ef4444",
        system:
          "You are a senior business leader focused on clear executive communication and decision context.",
        instruction:
          "Score clarity, tone, alignment to goals, and business impact. Call out sections that help or hinder executive understanding, and suggest concise rewrites.",
        enabled: true
      },
      {
        id: "legal",
        name: "Legal",
        color: "#2563eb",
        system:
          "You are corporate counsel focused on risk, claims, IP, and contractual language.",
        instruction:
          "Flag ambiguous or risky statements; note claims needing substantiation; suggest safer phrasing and disclaimers.",
        enabled: true
      },
      {
        id: "hr",
        name: "HR Partner",
        color: "#f59e0b",
        system:
          "You are an HR partner focused on inclusive, respectful language and change-management.",
        instruction:
          "Identify wording that could be exclusionary or unclear to broad audiences; propose inclusive alternatives and change-management considerations.",
        enabled: true
      },
      {
        id: "tech-lead",
        name: "Technical Lead",
        color: "#10b981",
        system: "You are a pragmatic engineering lead who values feasibility and risk awareness.",
        instruction:
          "Call out technical risks, missing constraints, unclear dependencies, and unrealistic timelines. Suggest concrete mitigations.",
        enabled: true
      },
      {
        id: "junior-analyst",
        name: "Junior Analyst",
        color: "#8b5cf6",
        system:
          "You are a curious analyst who asks clarifying questions and hunts for missing data.",
        instruction:
          "Point out unclear logic, missing data, and assumptions. Ask 2–3 clarifying questions to de-risk decisions.",
        enabled: true
      }
    ]
  },
  {
    id: "marketing-focus",
    name: "Marketing Focus Group",
    personas: [
      {
        id: "midwest-parent",
        name: "Midwest Parent",
        color: "#f97316",
        system:
          "You are a practical parent from the Midwest focused on value, trust, and family benefit.",
        instruction:
          "Does this feel trustworthy, useful, and affordable for a family like yours? What would make you say yes/no?",
        enabled: true
      },
      {
        id: "genz-student",
        name: "Gen-Z Student",
        color: "#06b6d4",
        system: "You are a Gen-Z student with strong BS detection; you value authenticity and ease.",
        instruction:
          "Call out cringe or corporate tone. Is this exciting and effortless to adopt? What would you share with friends?",
        enabled: true
      },
      {
        id: "retired-veteran",
        name: "Retired Veteran",
        color: "#1e3a8a",
        system: "You are a retired veteran who values clarity, respect, and directness.",
        instruction:
          "React to credibility and plain language. Note any patronizing or vague phrasing; suggest straightforward fixes.",
        enabled: true
      },
      {
        id: "smb-owner",
        name: "Small Business Owner",
        color: "#16a34a",
        system: "You run a small local business and make pragmatic ROI-driven decisions.",
        instruction:
          "Does this save time or make money? Flag impractical steps or fluffy claims; request proof points.",
        enabled: true
      },
      {
        id: "tech-pro",
        name: "Tech Professional",
        color: "#14b8a6",
        system: "You are a hands-on technologist who values precision and evidence.",
        instruction:
          "Call out vague claims, missing technical specifics, and inconsistent terminology. Prefer precise, testable language.",
        enabled: true
      }
    ]
  },
  {
    id: "startup-stakeholders",
    name: "Startup Stakeholders",
    personas: [
      {
        id: "founder-ceo",
        name: "Founder/CEO",
        color: "#ea580c",
        system: "You are a startup founder focused on narrative, mission, and speed to impact.",
        instruction:
          "Is the story compelling and credible? Identify blockers, sharpen the ask, and propose the next concrete milestone.",
        enabled: true
      },
      {
        id: "cto",
        name: "CTO",
        color: "#0ea5e9",
        system: "You are a CTO balancing innovation with execution reality.",
        instruction:
          "Probe architecture tradeoffs, data risks, staffing assumptions, and rollout safety. Suggest lean alternatives.",
        enabled: true
      },
      {
        id: "vp-sales",
        name: "VP Sales",
        color: "#f43f5e",
        system:
          "You are a revenue leader focused on ICP, pain points, repeatability, and proof points.",
        instruction:
          "Flag weak ICP definition, unclear CTA, and missing evidence. Recommend specific customer references or metrics.",
        enabled: true
      },
      {
        id: "design-lead",
        name: "Design Lead",
        color: "#8b5cf6",
        system: "You are a design lead oriented to clarity, hierarchy, and user empathy.",
        instruction:
          "Call out confusing structure, jargon, and missing states/visuals. Suggest hierarchy and microcopy improvements.",
        enabled: true
      },
      {
        id: "investor",
        name: "Investor",
        color: "#22c55e",
        system: "You are an investor seeking traction, team, and market realism.",
        instruction:
          "Challenge market size, differentiation, competitive dynamics, and signs of PMF. Ask critical due-diligence questions.",
        enabled: true
      }
    ]
  },
  {
    id: "political-spectrum",
    name: "Political Spectrum",
    personas: [
      {
        id: "progressive",
        name: "Progressive",
        color: "#22c55e",
        system: "You evaluate proposals through equity, inclusion, and climate justice.",
        instruction:
          "Assess community impact and equitable outcomes; flag exploitative framing; suggest inclusive alternatives.",
        enabled: true
      },
      {
        id: "centrist",
        name: "Centrist",
        color: "#3b82f6",
        system: "You favor pragmatic compromise and institutional stability.",
        instruction:
          "Identify balanced tradeoffs, fiscal prudence, and feasible sequencing. Call out polarization risks.",
        enabled: true
      },
      {
        id: "conservative",
        name: "Conservative",
        color: "#ef4444",
        system: "You value tradition, personal responsibility, and limited government.",
        instruction:
          "Probe unintended consequences, cost burdens, and regulatory overreach. Prefer incremental measures.",
        enabled: true
      },
      {
        id: "libertarian",
        name: "Libertarian",
        color: "#eab308",
        system: "You prize individual freedom and market solutions.",
        instruction:
          "Flag paternalism; prefer voluntary, decentralized approaches; highlight property and privacy concerns.",
        enabled: true
      },
      {
        id: "nonpartisan-analyst",
        name: "Nonpartisan Analyst",
        color: "#64748b",
        system: "You are a neutral analyst focused on evidence and methodology.",
        instruction:
          "Check claims, cite sources, and call out uncertainty clearly. Prefer replication and pre-registration.",
        enabled: true
      },
      {
        id: "local-community",
        name: "Local Community",
        color: "#0ea5e9",
        system: "You represent neighborhood priorities—safety, schools, livability.",
        instruction:
          "Assess real-world impact on daily life; surface practical concerns like traffic, noise, and services.",
        enabled: true
      }
    ]
  },
  {
    id: "academic-review",
    name: "Academic Review Board",
    personas: [
      {
        id: "methods-reviewer",
        name: "Methods Reviewer",
        color: "#22c55e",
        system: "You focus on study design, validity, and reproducibility.",
        instruction:
          "Flag bias, confounders, and missing controls; suggest design fixes; comment on external validity.",
        enabled: true
      },
      {
        id: "literature-reviewer",
        name: "Literature Reviewer",
        color: "#2563eb",
        system: "You ensure claims are situated within the literature.",
        instruction:
          "Point to missing citations, contradictory findings, and prior art; recommend essential sources.",
        enabled: true
      },
      {
        id: "ethics-reviewer",
        name: "Ethics Reviewer",
        color: "#ef4444",
        system: "You evaluate participant risk, consent, and data handling.",
        instruction:
          "Call out ethics gaps and propose mitigation steps; ensure IRB-quality safeguards.",
        enabled: true
      }
    ]
  },
  {
    id: "enterprise-governance",
    name: "Enterprise Governance",
    personas: [
      {
        id: "security",
        name: "Security",
        color: "#0ea5e9",
        system: "You evaluate data security, identity, and threat surface.",
        instruction:
          "Identify data flow risks, compliance gaps, least-privilege violations, and hardening steps.",
        enabled: true
      },
      {
        id: "privacy",
        name: "Privacy",
        color: "#f97316",
        system: "You focus on data minimization and user rights.",
        instruction:
          "Flag over-collection and unclear retention; suggest DPIA needs; ensure lawful bases.",
        enabled: true
      },
      {
        id: "finance",
        name: "Finance",
        color: "#22c55e",
        system: "You evaluate costs, ROI, and budgeting realism.",
        instruction:
          "Check COGS, headcount, hidden costs; validate payback and cash-flow timing; propose phased plans.",
        enabled: true
      }
    ]
  },
  {
    id: "public-sector",
    name: "Public Sector Advisory",
    personas: [
      {
        id: "procurement",
        name: "Procurement",
        color: "#8b5cf6",
        system: "You focus on fair bidding, specs, and vendor risk.",
        instruction:
          "Flag sole-source risk, vague requirements, and weak KPIs; propose measurable success criteria.",
        enabled: true
      },
      {
        id: "counsel",
        name: "Counsel",
        color: "#ef4444",
        system: "You ensure statutory compliance and a defensible record.",
        instruction:
          "Note open-meeting transparency, FOIL/FOIA, conflict handling, and auditability.",
        enabled: true
      },
      {
        id: "ombudsman",
        name: "Ombudsman",
        color: "#10b981",
        system: "You represent citizens’ service quality and recourse.",
        instruction:
          "Surface access barriers and escalation paths; assess fairness and equity impacts.",
        enabled: true
      }
    ]
  },
  {
    id: "nonprofit-board",
    name: "Nonprofit Board",
    personas: [
      {
        id: "board-chair",
        name: "Board Chair",
        color: "#e11d48",
        system: "You drive mission alignment and fiduciary duty.",
        instruction:
          "Call out mission drift, weak metrics, and governance risks; recommend board actions.",
        enabled: true
      },
      {
        id: "treasurer",
        name: "Treasurer",
        color: "#22c55e",
        system: "You scrutinize budgets, reserves, and restricted funds.",
        instruction:
          "Probe sustainability and cash risk; propose realistic pacing; check donor restrictions.",
        enabled: true
      },
      {
        id: "program-director",
        name: "Program Director",
        color: "#3b82f6",
        system: "You ensure program logic and beneficiary outcomes.",
        instruction:
          "Test theory of change, evaluation plan, and capability gaps; request outcome measures.",
        enabled: true
      }
    ]
  },
  {
    id: "product-trio",
    name: "Product Development Trio",
    personas: [
      {
        id: "pm",
        name: "Product Manager",
        color: "#22c55e",
        system: "You balance user value, viability, and timelines.",
        instruction:
          "Is the problem clear? Are scope cuts obvious? What metrics matter? Identify the smallest successful slice.",
        enabled: true
      },
      {
        id: "design",
        name: "Design",
        color: "#8b5cf6",
        system: "You advocate for clarity, accessibility, and hierarchy.",
        instruction:
          "Flag confusing flows and missing states; suggest structure and microcopy improvements.",
        enabled: true
      },
      {
        id: "engineering",
        name: "Engineering",
        color: "#2563eb",
        system: "You are pragmatic about delivery and reliability.",
        instruction:
          "Surface risky assumptions, data/infra needs, rollout plan, and operational constraints.",
        enabled: true
      }
    ]
  },
  {
    id: "support-voices",
    name: "Customer Support Voices",
    personas: [
      {
        id: "support-agent",
        name: "Support Agent",
        color: "#22c55e",
        system: "You care about clarity, sentiment, and deflection rate.",
        instruction:
          "Is the content self-serveable? What macros or links are missing? Flag sentiment risks.",
        enabled: true
      },
      {
        id: "support-manager",
        name: "Support Manager",
        color: "#ef4444",
        system: "You optimize quality, CSAT, and backlog health.",
        instruction:
          "Flag SLA risks, handoff gaps, and escalation criteria. Recommend process adjustments.",
        enabled: true
      },
      {
        id: "documentation",
        name: "Documentation",
        color: "#0ea5e9",
        system: "You maintain help-center clarity and coverage.",
        instruction:
          "Note missing how-to steps, screenshots, and version drift. Suggest specific doc changes.",
        enabled: true
      }
    ]
  },
  {
    id: "localization",
    name: "International Localization",
    personas: [
      {
        id: "emea-reader",
        name: "EMEA Reader",
        color: "#16a34a",
        system: "You check UK/EU spelling, idioms, and cultural references.",
        instruction:
          "Flag US-centric terms; propose neutral alternatives; note localization blockers.",
        enabled: true
      },
      {
        id: "latam-reader",
        name: "LATAM Reader",
        color: "#ef4444",
        system: "You check Spanish/Portuguese borrowings and tone.",
        instruction:
          "Call out untranslatable idioms; suggest simplified phrasing; watch formality.",
        enabled: true
      },
      {
        id: "apac-reader",
        name: "APAC Reader",
        color: "#2563eb",
        system: "You check regional sensibilities and formality norms.",
        instruction:
          "Flag honorifics/collectivist phrasing mismatches; keep clarity first.",
        enabled: true
      }
    ]
  },
  {
    id: "a11y-compliance",
    name: "Accessibility & Compliance",
    personas: [
      {
        id: "plain-language",
        name: "Plain Language",
        color: "#f97316",
        system: "You enforce plain-language guidelines and readability.",
        instruction:
          "Highlight long sentences, passive voice, and jargon; suggest simpler rewrites.",
        enabled: true
      },
      {
        id: "accessibility-auditor",
        name: "Accessibility Auditor",
        color: "#22c55e",
        system: "You care about WCAG and inclusive patterns.",
        instruction:
          "Flag color/contrast risks, keyboard traps, and missing alt text; propose fixes.",
        enabled: true
      },
      {
        id: "compliance-officer",
        name: "Compliance Officer",
        color: "#2563eb",
        system: "You ensure policy/process conformance and approvals.",
        instruction:
          "Note missing disclosures, approvals, and retention guidance; propose compliant language.",
        enabled: true
      }
    ]
  },
  {
    id: "editorial-board",
    name: "Editorial Board",
    personas: [
      {
        id: "copy-editor",
        name: "Copy Editor",
        color: "#22c55e",
        system: "You fix grammar and clarity.",
        instruction: "Suggest concise rewrites; remove redundancy; ensure consistency.",
        enabled: true
      },
      {
        id: "fact-checker",
        name: "Fact Checker",
        color: "#2563eb",
        system: "You verify claims and dates.",
        instruction: "Flag doubtful facts and request sources; recommend citations.",
        enabled: true
      },
      {
        id: "style-editor",
        name: "Style Editor",
        color: "#f43f5e",
        system: "You keep tone and style on-brand.",
        instruction: "Call out off-brand voice and formatting issues; propose style-aligned fixes.",
        enabled: true
      }
    ]
  }
];

/* ===============================
   State
================================= */
let ALL_SETS: PersonaSet[] = DEFAULT_SETS;
let SETTINGS: AppSettings = {};
let ACTIVE_SET: PersonaSet | null = null;
let ACTIVE_PERSONAS: Persona[] = [];

/* ===============================
   DOM helpers
================================= */
const $ = <T extends HTMLElement = HTMLElement>(id: string) =>
  document.getElementById(id) as T | null;

function firstById<T extends HTMLElement = HTMLElement>(ids: string[]): T | null {
  for (const id of ids) {
    const el = document.getElementById(id);
    if (el) return el as T;
  }
  return null;
}

function show(el: HTMLElement | null, on: boolean) {
  if (!el) return;
  el.style.display = on ? "" : "none";
}

function text(el: HTMLElement | null, s: string) {
  if (el) el.textContent = s;
}

function toast(msg: string) {
  console.log("[PF]", msg);
  const t = firstById<HTMLDivElement>(["toast", "pfToast"]);
  if (!t) return;
  t.textContent = msg;
  t.classList.add("show");
  setTimeout(() => t.classList.remove("show"), 2200);
}

async function confirmAsync(message: string): Promise<boolean> {
  const modal = firstById<HTMLDialogElement>(["confirmModal", "pfConfirm"]);
  if (!modal) return window.confirm(message);
  const msg = modal.querySelector<HTMLElement>("[data-confirm-message]");
  if (msg) msg.textContent = message;
  const ok = modal.querySelector<HTMLButtonElement>("[data-confirm-ok]");
  const cancel = modal.querySelector<HTMLButtonElement>("[data-confirm-cancel]");
  return new Promise((resolve) => {
    const done = (v: boolean) => {
      ok?.removeEventListener("click", onOk);
      cancel?.removeEventListener("click", onCancel);
      (modal as any).close?.();
      resolve(v);
    };
    const onOk = () => done(true);
    const onCancel = () => done(false);
    (modal as any).showModal?.();
    ok?.addEventListener("click", onOk);
    cancel?.addEventListener("click", onCancel);
  });
}

/* ===============================
   Color → emoji (for Word comment prefix)
================================= */
function hueFromColor(color?: string): number | null {
  if (!color) return null;
  const c = color.trim().toLowerCase();
  if (c.startsWith("hsl(")) {
    const m = c.match(/hsl\(\s*([0-9.]+)\s*,/);
    return m ? Number(m[1]) : null;
  }
  if (c.startsWith("#")) {
    // hex → hue
    const hex = c.replace("#", "");
    const n = hex.length === 3
      ? hex.split("").map((h) => h + h).join("")
      : hex;
    const r = parseInt(n.slice(0, 2), 16) / 255;
    const g = parseInt(n.slice(2, 4), 16) / 255;
    const b = parseInt(n.slice(4, 6), 16) / 255;
    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    const d = max - min;
    if (d === 0) return 0;
    let h = 0;
    switch (max) {
      case r: h = ((g - b) / d) % 6; break;
      case g: h = (b - r) / d + 2; break;
      case b: h = (r - g) / d + 4; break;
    }
    h = Math.round(h * 60);
    if (h < 0) h += 360;
    return h;
  }
  return null;
}

function emojiForColor(color?: string): string {
  const h = hueFromColor(color);
  if (h === null) return "🟦";
  if (h < 15 || h >= 345) return "🟥";         // red
  if (h < 45) return "🟧";                     // orange
  if (h < 70) return "🟨";                     // yellow
  if (h < 170) return "🟩";                    // green
  if (h < 250) return "🟦";                    // blue
  if (h < 320) return "🟪";                    // purple
  return "🟪";                                 // magenta/pink-ish → purple
}

function makePersonaPrefix(p: Persona) {
  // Tag with [PF] so we can safely delete only our comments later
  const mark = emojiForColor(p.color);
  return `[PF] ${mark} ${p.name}`;
}

/* ===============================
   Settings & View switching
================================= */
const LS_KEY = "pf.settings.v1";

function loadSettings() {
  try {
    SETTINGS = JSON.parse(localStorage.getItem(LS_KEY) || "{}");
  } catch {
    SETTINGS = {};
  }
}
function saveSettings() {
  localStorage.setItem(LS_KEY, JSON.stringify(SETTINGS));
}

function initViewSwitching() {
  const main = firstById<HTMLElement>([
    "mainView",
    "reviewView",
    "mainSection",
    "reviewSection",
    "main",
    "view-main",
    "pfMain"
  ]);
  const settings = firstById<HTMLElement>([
    "settingsView",
    "settingsSection",
    "settings",
    "view-settings",
    "pfSettings"
  ]);
  function switchView(which: "main" | "settings") {
    show(main, which === "main");
    show(settings, which === "settings");
  }
  (window as any)._pfSwitchView = switchView;
  switchView("main");

  const settingsBtn = firstById<HTMLButtonElement>(["settingsBtn", "btnSettings", "pfSettingsBtn"]);
  if (settingsBtn) {
    settingsBtn.onclick = () => {
      const isOpen = settings && settings.style.display !== "none";
      switchView(isOpen ? "main" : "settings");
    };
  }
  const backBtn = firstById<HTMLButtonElement>([
    "backBtn",
    "backToMainBtn",
    "closeSettingsBtn",
    "homeBtn",
    "pfBackBtn"
  ]);
  if (backBtn) backBtn.onclick = () => switchView("main");
  window.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && settings && settings.style.display !== "none") switchView("main");
  });
}

/* ===============================
   Persona UI
================================= */
function selectInitialSet() {
  const desired = SETTINGS.selectedSetId || ALL_SETS[0]?.id;
  ACTIVE_SET = ALL_SETS.find((s) => s.id === desired) || ALL_SETS[0] || null;
  const enabled = SETTINGS.personaEnabled || {};
  ACTIVE_PERSONAS =
    ACTIVE_SET?.personas.map((p) => ({ ...p, enabled: enabled[p.id] ?? true })) || [];
}

function renderPersonaControls() {
  const setSelect =
    firstById<HTMLSelectElement>(["personaSetSelect", "personaSet", "pfPersonaSet"]) ||
    createSetSelect();
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
    selectInitialSet();
    renderPersonaList();
  };
  renderPersonaList();
}

function createSetSelect(): HTMLSelectElement {
  const host =
    firstById<HTMLDivElement>(["personaControls", "personaControlBar", "pfPersonaBar"]) ||
    document.body;
  const sel = document.createElement("select");
  sel.id = "personaSetSelect";
  sel.className = "pf-select";
  host.appendChild(sel);
  return sel;
}

function renderPersonaList() {
  const list =
    firstById<HTMLDivElement>(["personaChecklist", "personaList", "pfPersonaList"]) ||
    createPersonaListMount();
  list.innerHTML = "";
  for (const p of ACTIVE_PERSONAS) {
    const row = document.createElement("label");
    row.className = "pf-row pf-persona-row";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = !!p.enabled;
    cb.onchange = () => {
      p.enabled = cb.checked;
      SETTINGS.personaEnabled = SETTINGS.personaEnabled || {};
      SETTINGS.personaEnabled[p.id] = !!p.enabled;
      saveSettings();
    };

    const sw = document.createElement("span");
    sw.className = "pf-swatch";
    if (p.color) sw.style.background = p.color;

    const nm = document.createElement("span");
    nm.className = "pf-persona-name";
    nm.textContent = p.name;

    row.appendChild(cb);
    row.appendChild(sw);
    row.appendChild(nm);
    list.appendChild(row);
  }
}

function createPersonaListMount(): HTMLDivElement {
  const host =
    firstById<HTMLDivElement>(["personaControls", "personaControlBar", "pfPersonaBar"]) ||
    document.body;
  const div = document.createElement("div");
  div.id = "personaChecklist";
  div.className = "pf-list";
  host.appendChild(div);
  return div;
}

/* ===============================
   Buttons / Wiring
================================= */
function wireButtons() {
  const runBtn = firstById<HTMLButtonElement>(["runBtn", "btnRun", "pfRun"]);
  if (runBtn) runBtn.onclick = () => void runAllEnabledPersonas();

  const retryBtn = firstById<HTMLButtonElement>(["retryBtn", "btnRetry", "pfRetry"]);
  if (retryBtn) retryBtn.onclick = () => void runAllEnabledPersonas();

  const exportBtn = firstById<HTMLButtonElement>(["exportBtn", "btnExport", "pfExport"]);
  if (exportBtn) exportBtn.onclick = () => void exportReport();

  const clearBtn = firstById<HTMLButtonElement>(["clearBtn", "btnClear", "pfClear"]);
  if (clearBtn) clearBtn.onclick = () => void clearOurCommentsOnly();

  const clearAllBtn = firstById<HTMLButtonElement>(["clearAllBtn", "btnClearAll", "pfClearAll"]);
  if (clearAllBtn) {
    clearAllBtn.onclick = async () => {
      const ok = await confirmAsync("Delete ALL comments in the document?");
      if (!ok) return;
      try {
        const n = await clearAllComments();
        if (n === -1) toast("This version of Word doesn't support listing comments.");
        else toast(n > 0 ? `Deleted ${n} comment(s).` : "No comments found.");
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

/* ===============================
   Bootstrap
================================= */
Office.onReady(async () => {
  try {
    document.body.style.minWidth = "520px";
    initViewSwitching();
    loadSettings();
    selectInitialSet();
    renderPersonaControls();
    wireButtons();
    toast("UI initialized");
  } catch (e) {
    console.error(e);
    toast("Init error");
  }
});

/* ===============================
   Core run loop
================================= */
async function runAllEnabledPersonas() {
  if (!ACTIVE_SET) return toast("No persona set selected");
  const enabled = ACTIVE_PERSONAS.filter((p) => p.enabled);
  if (enabled.length === 0) return toast("Enable at least one persona");

  const docText = await getDocumentText();
  if (!docText?.trim()) return toast("Document appears empty");

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

/* ===============================
   Providers
================================= */
async function reviewWithPersona(persona: Persona, documentText: string): Promise<FeedbackJSON> {
  const provider =
    SETTINGS.provider || ({ kind: "openrouter", model: "gpt-4o-mini", apiKey: "" } as ProviderConfig);

  const messages: ChatMessage[] = [
    { role: "system", content: persona.system },
    {
      role: "user",
      content: `${persona.instruction}

Return ONLY valid JSON in the shape:
{"scores":{"<metric>":1-5},"global_feedback":"string","comments":[{"quote":"string","comment":"string","category":"optional"}]}

Document:
<<<DOC
${documentText}
DOC>>>`
    }
  ];

  let raw = "";
  if (provider.kind === "openrouter") {
    if (!provider.apiKey) {
      const k =
        window.prompt("Enter your OpenRouter API key (kept in-memory for this session):", "") || "";
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
    body: JSON.stringify({ model, messages, temperature: 0.2 })
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

/* ===============================
   JSON coercion/validation
================================= */
function coerceJSON(s: string): FeedbackJSON {
  const m = s.match(/\{[\s\S]*\}$/);
  const body = m ? m[0] : s;
  try {
    return JSON.parse(body);
  } catch {
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
  for (const k of Object.keys(o.scores)) {
    const v = Number(o.scores[k]);
    if (!Number.isFinite(v)) delete o.scores[k];
    else o.scores[k] = Math.max(1, Math.min(5, Math.round(v)));
  }
  o.comments = o.comments.map((c: any) => ({
    quote: String(c.quote || "").slice(0, 240),
    comment: String(c.comment || ""),
    category: c.category ? String(c.category) : undefined
  }));
}

/* ===============================
   Word integration
================================= */
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
          matchWholeWord: false
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
        // attach at start to avoid losing feedback — use the enum to satisfy TS
        context.document.body.getRange(Word.RangeLocation.start).insertComment(
          `${prefix} (unmatched) ${payload}`
        );
      }
      await context.sync();
    }
  });
}

function fuzzySnippet(quote: string, docText: string): string | null {
  if (quote.length <= 30) return quote;
  const mid = Math.floor(quote.length / 2);
  const snip = quote.slice(Math.max(0, mid - 15), mid + 15);
  return docText.includes(snip) ? snip : null;
}

/** Delete ALL comments (returns -1 if comments collection is unavailable in this Word build) */
async function clearAllComments(): Promise<number> {
  return Word.run(async (context) => {
    const docAny = context.document as any;
    const comments = docAny.comments; // older typings may not declare Document.comments
    if (!comments) return -1;
    comments.load("items");
    await context.sync();
    const n = comments.items.length;
    for (const c of comments.items) c.delete();
    await context.sync();
    return n;
  });
}

/** Delete only comments we added (prefixed with [PF]) */
async function clearOurCommentsOnly(): Promise<number> {
  return Word.run(async (context) => {
    const docAny = context.document as any;
    const comments = docAny.comments;
    if (!comments) {
      toast("This Word version doesn't support listing comments.");
      return -1;
    }
    comments.load("items");
    await context.sync();
    let n = 0;
    for (const c of comments.items) {
      (c as any).content?.load?.("text");
    }
    await context.sync();
    for (const c of comments.items) {
      const t = (c as any).content?.text ?? "";
      if (/^\[PF\]\s/.test(t)) {
        c.delete();
        n++;
      }
    }
    await context.sync();
    toast(n > 0 ? `Deleted ${n} comment(s) from Feedback Personas` : "No PF comments found.");
    return n;
  });
}

/* ===============================
   Results & export
================================= */
function appendResults(persona: Persona, out: FeedbackJSON) {
  const container = firstById<HTMLDivElement>(["results", "pfResults"]);
  if (!container) return;

  const card = document.createElement("div");
  card.className = "pf-card";

  const h = document.createElement("div");
  h.className = "pf-card-title";
  h.textContent = `${persona.name}`;
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

function exportReport() {
  const container = firstById<HTMLDivElement>(["results", "pfResults"]);
  const html = `
  <html><head><meta charset="utf-8"/><title>Feedback Report</title>
    <style>body{font:14px system-ui,Segoe UI,Roboto,Arial} h1{font-size:20px;margin:0 0 12px} .card{border:1px solid #ddd;border-radius:8px;padding:12px;margin:12px 0} .scores td{padding:2px 6px}</style>
  </head><body><h1>Feedback Report</h1>${container ? container.innerHTML : "<p>No results</p>"}</body></html>`;
  const w = window.open("", "_blank");
  if (!w) return;
  w.document.write(html);
  w.document.close();
  w.focus();
  w.print();
}

/* ===============================
   Expose for debugging
================================= */
(Object.assign(window as any, {
  _pf: {
    runAllEnabledPersonas,
    clearAllComments,
    clearOurCommentsOnly
  }
}) as any);
