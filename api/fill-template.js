/**
 * CTRL PoC Export Service · fill-template (V12)
 * Place at: /api/fill-template.js  (ctrl-poc-service)
 *
 * V12 requirements implemented:
 * - ONLY uses NEW keys (no legacy packSection, no old p3/p4/p5 mappings)
 * - Debug via &debug=1 inspects every critical aspect (payload shape, types, lengths, bands, template selection, layout, page count, etc.)
 * - D1 now sourced from ctrl.summary (PoC_Summaries / PoC_FINAL.summary forwarded) for debug visibility
 * - PDF positions:
 *   Page 1: FullName + Date (no change)
 *   Pages 2–8: header FullName (no change; total 8 pages)
 *   Page 3: exec_summary_para1 + exec_summary_para2
 *   Page 4: ctrl_overview_para1 + ctrl_overview_para2 + chart (chart moved here)
 *   Page 5: ctrl_deepdive_para1 + ctrl_deepdive_para2 + themes_para1 + themes_para2
 *   Page 6: workWith page (C/T/R/L boxes) (was page 8)
 *   Page 7: Act1..Act6
 * - No change to build-link BASE64 convention (GET ?data=...).
 * - No change to template selection & fallback behaviour (still dom/second combos).
 */

export const config = { runtime: "nodejs" };

import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { PDFDocument, StandardFonts } from "pdf-lib";

/* ───────────── utilities ───────────── */
const S = (v, fb = "") => (v == null ? String(fb) : String(v));
const N = (v, fb = 0) => (Number.isFinite(+v) ? +v : fb);
const norm = (s) => S(s).replace(/\s+/g, " ").trim();

function safeJson(obj) {
  try { return JSON.parse(JSON.stringify(obj)); }
  catch { return { _error: "Could not serialise debug object" }; }
}

/* ───────── TL→BL rect helper ───────── */
const rectTLtoBL = (page, box, inset = 0) => {
  const pageH = page.getHeight();
  const x = N(box.x) + inset;
  const w = Math.max(0, N(box.w) - inset * 2);

  const hRaw = N(box.h || 400);
  const h = Math.max(0, hRaw - inset * 2);

  const y = pageH - N(box.y) - hRaw + inset;
  return { x, y, w, h };
};

/* ───────── text box helper ───────── */
function drawTextBox(page, font, text, box, opts = {}) {
  if (!page || !font || !box) return;
  const t0 = String(text ?? "").trim();
  if (!t0) return;

  const { x, y, w, h } = rectTLtoBL(page, box, 0);
  const size = N(box.size || opts.size || 12);
  const lineHeight = N(opts.lineHeight || Math.round(size * 1.3));
  const maxLines = N(opts.maxLines || box.maxLines || 999);
  const align = opts.align || box.align || "left";

  const paragraphs = String(t0).split("\n");
  const lines = [];

  for (const para of paragraphs) {
    const words = para.split(/\s+/).filter(Boolean);
    let line = "";

    for (const word of words) {
      const test = line ? line + " " + word : word;
      const width = font.widthOfTextAtSize(test, size);
      if (width <= w) {
        line = test;
      } else {
        if (line) lines.push(line);
        line = word;
      }
      if (lines.length >= maxLines) break;
    }
    if (lines.length >= maxLines) break;
    if (line) lines.push(line);

    if (lines.length < maxLines) lines.push("");
  }

  while (lines.length && lines[lines.length - 1] === "") lines.pop();

  const clipped = lines.slice(0, maxLines);
  const startY = y + h - lineHeight;

  for (let i = 0; i < clipped.length; i++) {
    const t = clipped[i];
    const tw = font.widthOfTextAtSize(t, size);
    let tx = x;
    if (align === "center") tx = x + (w - tw) / 2;
    if (align === "right") tx = x + (w - tw);
    const ty = startY - i * lineHeight;
    if (ty < y) break;
    page.drawText(t, { x: tx, y: ty, size, font });
  }
}

/* ───────── template loader ───────── */
async function loadTemplateBytesLocal(fname) {
  if (!fname.endsWith(".pdf")) throw new Error(`Invalid template filename: ${fname}`);

  const __file = fileURLToPath(import.meta.url);
  const __dir = path.dirname(__file);

  const candidates = [
    path.join(process.cwd(), "public", fname),
    path.join(__dir, "..", "public", fname),
    path.join(__dir, "public", fname),
    path.join(__dir, fname),
  ];

  let lastErr;
  for (const pth of candidates) {
    try {
      return await fs.readFile(pth);
    } catch (err) {
      lastErr = err;
    }
  }
  throw new Error(
    `Template not found: ${fname}. Tried: ${candidates.join(" | ")}. Last: ${lastErr?.message || "no detail"}`
  );
}

/* ───────── payload parsing (GET ?data=... or POST JSON) ───────── */
async function readPayload(req) {
  if (req.method === "POST") {
    const chunks = [];
    for await (const ch of req) chunks.push(ch);
    const raw = Buffer.concat(chunks).toString("utf8") || "{}";
    try { return JSON.parse(raw); } catch { return {}; }
  }

  const url = new URL(req.url, "http://localhost");
  const dataB64 = url.searchParams.get("data") || "";
  if (!dataB64) return {};

  try {
    const raw = Buffer.from(dataB64, "base64").toString("utf8");
    return JSON.parse(raw);
  } catch {
    return {};
  }
}

/* ───────── dom/second detection ───────── */
function resolveStateKey(any) {
  const s = S(any).trim().toUpperCase();
  const c = s.charAt(0);
  if (["C", "T", "R", "L"].includes(c)) return c;

  const low = S(any).toLowerCase();
  if (low.includes("concealed")) return "C";
  if (low.includes("triggered")) return "T";
  if (low.includes("regulated")) return "R";
  if (low.includes("lead")) return "L";

  return null;
}

function computeDomAndSecondKeys(D) {
  // Prefer explicit keys from build-pdf
  const ctrl = D.ctrl || {};
  const summary = ctrl.summary || {};
  const raw = D.raw || D;

  const domKey =
    resolveStateKey(ctrl.dominantKey) ||
    resolveStateKey(raw.domKey) ||
    resolveStateKey(summary.domKey) ||
    resolveStateKey(summary.domState) ||
    "R";

  const secondKey =
    resolveStateKey(ctrl.secondKey) ||
    resolveStateKey(raw.secondKey) ||
    resolveStateKey(summary.secondKey) ||
    resolveStateKey(summary.secondState) ||
    null;

  // If second is missing, attempt compute from totals (best effort)
  if (secondKey) return { domKey, secondKey };

  const totals =
    summary.ctrlTotals ||
    summary.totals ||
    summary.mix ||
    ctrl.ctrlTotals ||
    ctrl.totals ||
    ctrl.mix ||
    raw.ctrlTotals ||
    raw.totals ||
    raw.mix ||
    {};

  const score = { C: 0, T: 0, R: 0, L: 0 };
  const addTotals = (obj) => {
    if (!obj || typeof obj !== "object") return;
    score.C += Number(obj.C ?? obj.concealed ?? obj.c ?? 0) || 0;
    score.T += Number(obj.T ?? obj.triggered ?? obj.t ?? 0) || 0;
    score.R += Number(obj.R ?? obj.regulated ?? obj.r ?? 0) || 0;
    score.L += Number(obj.L ?? obj.lead ?? obj.l ?? 0) || 0;
  };
  addTotals(totals);

  const ordered = ["C", "T", "R", "L"]
    .filter((k) => k !== domKey)
    .map((k) => [k, score[k]])
    .sort((a, b) => b[1] - a[1]);

  const fallbackSecond = ordered[0]?.[0] || (domKey === "C" ? "T" : "C");
  return { domKey, secondKey: fallbackSecond };
}

/* ───────── radar chart embed (QuickChart) ───────── */
function makeSpiderChartUrl12(bandsRaw) {
  const labels = [
    "C_low","C_mid","C_high","T_low","T_mid","T_high",
    "R_low","R_mid","R_high","L_low","L_mid","L_high",
  ];

  const vals = labels.map((k) => Number(bandsRaw?.[k] || 0));
  const maxVal = Math.max(...vals, 1);
  const scaled = vals.map((v) => (maxVal > 0 ? v / maxVal : 0));

  const cfg = {
    type: "radar",
    data: {
      labels,
      datasets: [{
        label: "",
        data: scaled,
        fill: true,
        borderWidth: 0,
        pointRadius: 0,
        backgroundColor: "rgba(184, 15, 112, 0.35)",
      }],
    },
    options: {
      plugins: { legend: { display: false } },
      scales: {
        r: {
          min: 0, max: 1,
          ticks: { display: false },
          grid: { display: false },
          angleLines: { display: false },
          pointLabels: { display: false },
        },
      },
    },
  };

  const enc = encodeURIComponent(JSON.stringify(cfg));
  return `https://quickchart.io/chart?c=${enc}&format=png&width=800&height=800&backgroundColor=transparent`;
}

async function embedRemoteImage(pdfDoc, url) {
  if (!url) return null;

  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to fetch chart: ${res.status} ${res.statusText}`);

  const buf = new Uint8Array(await res.arrayBuffer());
  const sig = String.fromCharCode(buf[0], buf[1], buf[2], buf[3] || 0);

  if (sig.startsWith("\x89PNG")) return await pdfDoc.embedPng(buf);
  if (sig.startsWith("\xff\xd8")) return await pdfDoc.embedJpg(buf);

  try { return await pdfDoc.embedPng(buf); }
  catch { return await pdfDoc.embedJpg(buf); }
}

async function embedRadarFromBands(pdfDoc, page, box, bandsRaw) {
  if (!pdfDoc || !page || !box || !bandsRaw) return;

  const hasAny =
    bandsRaw && typeof bandsRaw === "object" &&
    Object.values(bandsRaw).some((v) => Number(v) > 0);

  if (!hasAny) return;

  const url = makeSpiderChartUrl12(bandsRaw);
  const img = await embedRemoteImage(pdfDoc, url);
  if (!img) return;

  const H = page.getHeight();
  const { x, y, w, h } = box;
  page.drawImage(img, { x, y: H - y - h, width: w, height: h });
}

/* ───────── DEFAULT LAYOUT (8 pages only) ───────── */
const DEFAULT_LAYOUT = {
  pages: {
    p1: {
      name: { x: 7, y: 473, w: 500, h: 60, size: 30, align: "center", maxLines: 1 },
      date: { x: 210, y: 600, w: 500, h: 40, size: 25, align: "left", maxLines: 1 },
    },

    // Header name on pages 2–8 (used by handler loop)
    p2: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },

    // Page 3
    p3: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      exec1:   { x: 25,  y: 200, w: 550, h: 420, size: 18, align: "left", maxLines: 18 },
      exec2:   { x: 25,  y: 650, w: 550, h: 420, size: 18, align: "left", maxLines: 18 },
    },

    // Page 4 (overview + chart now here)
    p4: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      ov1:     { x: 25,  y: 200, w: 550, h: 420, size: 18, align: "left", maxLines: 18 },
      ov2:     { x: 25,  y: 650, w: 550, h: 420, size: 18, align: "left", maxLines: 18 },
      chart:   { x: 48,  y: 462, w: 500, h: 300 }, // keep style; moved from p5 in older versions
    },

    // Page 5 (deepdive + themes)
    p5: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      dd1:     { x: 25,  y: 200, w: 550, h: 360, size: 18, align: "left", maxLines: 16 },
      dd2:     { x: 25,  y: 590, w: 550, h: 360, size: 18, align: "left", maxLines: 16 },
      th1:     { x: 25,  y: 980, w: 550, h: 260, size: 18, align: "left", maxLines: 12 },
      th2:     { x: 25,  y: 1240, w: 550, h: 260, size: 18, align: "left", maxLines: 12 },
    },

    // Page 6 (workWith)
    p6: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      collabC: { x: 30,  y: 200, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
      collabT: { x: 320, y: 200, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
      collabR: { x: 30,  y: 650, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
      collabL: { x: 320, y: 650, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
    },

    // Page 7 (Acts)
    p7: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      act1: { x: 30, y: 200, w: 590, h: 150, size: 17, align: "left", maxLines: 6 },
      act2: { x: 30, y: 360, w: 590, h: 150, size: 17, align: "left", maxLines: 6 },
      act3: { x: 30, y: 520, w: 590, h: 150, size: 17, align: "left", maxLines: 6 },
      act4: { x: 30, y: 680, w: 590, h: 150, size: 17, align: "left", maxLines: 6 },
      act5: { x: 30, y: 840, w: 590, h: 150, size: 17, align: "left", maxLines: 6 },
      act6: { x: 30, y: 1000,w: 590, h: 150, size: 17, align: "left", maxLines: 6 },
    },

    // Page 8 (header only by default)
    p8: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
    },
  },
};

/* deep merge override into default */
function mergeLayout(overrides = null) {
  const base = JSON.parse(JSON.stringify(DEFAULT_LAYOUT));
  if (!overrides || typeof overrides !== "object") return base;

  const out = JSON.parse(JSON.stringify(base));
  const ov = overrides;

  if (ov.pages && typeof ov.pages === "object") {
    out.pages = out.pages || {};
    for (const pk of Object.keys(ov.pages)) {
      out.pages[pk] = { ...(out.pages[pk] || {}), ...(ov.pages[pk] || {}) };
    }
  }

  for (const k of Object.keys(ov)) {
    if (k === "pages") continue;
    out[k] = ov[k];
  }

  return out;
}

/* ───────── URL-driven layout overrides ───────── */
function applyLayoutOverridesFromUrl(layout, url) {
  if (!layout || !layout.pages || !url || !url.searchParams) return layout;

  const allowed = new Set(["x","y","w","h","size","maxLines","align"]);
  const pages = layout.pages;

  for (const [k, v] of url.searchParams.entries()) {
    if (!k.startsWith("L_")) continue;

    const bits = k.split("_"); // L_p3_exec1_y
    if (bits.length < 4) continue;

    const pageKey = bits[1];
    const boxKey  = bits[2];
    const prop    = bits.slice(3).join("_");

    if (!pages[pageKey] || !pages[pageKey][boxKey]) continue;
    if (!allowed.has(prop)) continue;

    if (prop === "align") {
      const a = String(v || "").toLowerCase();
      if (["left","center","right"].includes(a)) pages[pageKey][boxKey][prop] = a;
      continue;
    }

    const num = Number(v);
    if (!Number.isFinite(num)) continue;
    pages[pageKey][boxKey][prop] = num;
  }

  return layout;
}

/* ───────── input normaliser (NEW KEYS ONLY) ───────── */
function normaliseInput(d = {}) {
  const identity = d.identity || {};
  const ctrl = d.ctrl || {};
  const summary = ctrl.summary || {};
  const text = d.text || {};
  const workWith = d.workWith || {};

  // Name + Date
  const name = norm(identity.fullName || d.fullName || identity.name || "");
  const date = norm(identity.dateLabel || d.dateLbl || d.dateLabel || d.date || "");

  // Bands for chart (no change)
  const bandsRaw =
    (ctrl.bands && typeof ctrl.bands === "object" ? ctrl.bands : null) ||
    (summary.ctrl12 && typeof summary.ctrl12 === "object" ? summary.ctrl12 : null) ||
    (d.bands && typeof d.bands === "object" ? d.bands : null) ||
    {};

  // NEW PAGE TEXT KEYS (ONLY)
  const exec1 = text.exec_summary_para1 ?? "";
  const exec2 = text.exec_summary_para2 ?? "";

  const ov1 = text.ctrl_overview_para1 ?? "";
  const ov2 = text.ctrl_overview_para2 ?? "";

  const dd1 = text.ctrl_deepdive_para1 ?? "";
  const dd2 = text.ctrl_deepdive_para2 ?? "";

  const th1 = text.themes_para1 ?? "";
  const th2 = text.themes_para2 ?? "";

  // Acts (support both Act1 keys and act_1 keys, but ONLY expose Act1..Act6 to renderer)
  const Act1 = text.Act1 ?? text.act1 ?? text.act_1 ?? "";
  const Act2 = text.Act2 ?? text.act2 ?? text.act_2 ?? "";
  const Act3 = text.Act3 ?? text.act3 ?? text.act_3 ?? "";
  const Act4 = text.Act4 ?? text.act4 ?? text.act_4 ?? "";
  const Act5 = text.Act5 ?? text.act5 ?? text.act_5 ?? "";
  const Act6 = text.Act6 ?? text.act6 ?? text.act_6 ?? "";

  // D1 now originates inside summaries (surface for debug only)
  const d1Min = summary.d1Min || summary.d1 || d.d1 || null;
  const d1ByQ = summary.d1ByQ || null;

  return {
    raw: d,
    identity: { name, date },

    ctrl: {
      dominantKey: ctrl.dominantKey || d.domKey || summary.domKey || summary.domState || "",
      secondKey:   ctrl.secondKey   || d.secondKey || summary.secondKey || summary.secondState || "",
      templateKey: ctrl.templateKey || d.templateKey || "",
      summary: summary || {}
    },

    bands: bandsRaw,
    layout: d.layout || null,

    // Page 3
    exec1: norm(exec1),
    exec2: norm(exec2),

    // Page 4
    ov1: norm(ov1),
    ov2: norm(ov2),

    // Page 5
    dd1: norm(dd1),
    dd2: norm(dd2),
    th1: norm(th1),
    th2: norm(th2),

    // Page 6
    workWith: {
      concealed: norm(workWith.concealed || ""),
      triggered: norm(workWith.triggered || ""),
      regulated: norm(workWith.regulated || ""),
      lead:      norm(workWith.lead || ""),
    },

    // Page 7
    acts: {
      Act1: norm(Act1),
      Act2: norm(Act2),
      Act3: norm(Act3),
      Act4: norm(Act4),
      Act5: norm(Act5),
      Act6: norm(Act6),
    },

    // Debug-only D1
    d1: {
      d1Min: okObj(d1Min) ? d1Min : (d1Min ? { value: d1Min } : null),
      d1ByQ: okObj(d1ByQ) ? d1ByQ : null,
    }
  };
}

/* ───────── main handler ───────── */
export default async function handler(req, res) {
  try {
    const url = new URL(req.url, "http://localhost");
    const debug = url.searchParams.get("debug") === "1";

    const payload = await readPayload(req);
    const P = normaliseInput(payload);

    // Layout: default + payload overrides + URL overrides
    let layout = mergeLayout(P.layout);
    layout = applyLayoutOverridesFromUrl(layout, url);
    const L = (layout && layout.pages) ? layout.pages : DEFAULT_LAYOUT.pages;

    // Determine template combo (no change to selection+fallback)
    const { domKey, secondKey } = computeDomAndSecondKeys({ ...payload, ctrl: P.ctrl });
    const combo = `${domKey}${secondKey}`;

    const validCombos = new Set(["CT","CL","CR","TC","TR","TL","RC","RT","RL","LC","LR","LT"]);
    const safeCombo = validCombos.has(combo) ? combo : "CT";
    const tpl = `CTRL_PoC_Assessment_Profile_template_${safeCombo}.pdf`;

    // Debug: audit EVERYTHING important
    if (debug) {
      const bandsKeys = Object.keys(P.bands || {});
      const labels12 = [
        "C_low","C_mid","C_high","T_low","T_mid","T_high",
        "R_low","R_mid","R_high","L_low","L_mid","L_high",
      ];
      const bandsPresent = labels12.filter(k => P.bands?.[k] != null).length;
      const bandsAnyPositive = labels12.some(k => Number(P.bands?.[k] || 0) > 0);

      const dbg = {
        ok: true,
        where: "fill-template:V12:debug",
        method: req.method,
        url: req.url,
        template: { combo, safeCombo, tpl },

        identity: {
          name: P.identity.name,
          nameLen: (P.identity.name || "").length,
          date: P.identity.date,
          dateLen: (P.identity.date || "").length
        },

        pagesExpected: 8,
        layoutKeys: Object.keys(L || {}),
        layoutPageBoxes: Object.fromEntries(
          Object.entries(L || {}).map(([pk, pv]) => [pk, Object.keys(pv || {})])
        ),

        textLengths: {
          exec1: P.exec1.length,
          exec2: P.exec2.length,
          ov1: P.ov1.length,
          ov2: P.ov2.length,
          dd1: P.dd1.length,
          dd2: P.dd2.length,
          th1: P.th1.length,
          th2: P.th2.length,
          collabC: P.workWith.concealed.length,
          collabT: P.workWith.triggered.length,
          collabR: P.workWith.regulated.length,
          collabL: P.workWith.lead.length,
          Act1: P.acts.Act1.length,
          Act2: P.acts.Act2.length,
          Act3: P.acts.Act3.length,
          Act4: P.acts.Act4.length,
          Act5: P.acts.Act5.length,
          Act6: P.acts.Act6.length,
        },

        types: {
          payloadType: typeof payload,
          identityType: typeof payload?.identity,
          ctrlType: typeof payload?.ctrl,
          textType: typeof payload?.text,
          workWithType: typeof payload?.workWith,
          bandsType: typeof payload?.ctrl?.bands,
        },

        bands: {
          keysCount: bandsKeys.length,
          keysPreview: bandsKeys.slice(0, 20),
          present12Count: bandsPresent,
          missing12: labels12.filter(k => P.bands?.[k] == null),
          anyPositive: bandsAnyPositive,
          sample: Object.fromEntries(labels12.map(k => [k, P.bands?.[k] ?? null]).slice(0, 6)),
        },

        d1_from_summary: safeJson(P.d1),

        warnings: [
          ...(P.identity.name ? [] : ["Missing identity.name"]),
          ...(P.identity.date ? [] : ["Missing identity.date"]),
          ...(P.exec1 || P.exec2 ? [] : ["Missing exec_summary_para1/para2"]),
          ...(P.ov1 || P.ov2 ? [] : ["Missing ctrl_overview_para1/para2"]),
          ...(P.dd1 || P.dd2 ? [] : ["Missing ctrl_deepdive_para1/para2"]),
          ...(P.th1 || P.th2 ? [] : ["Missing themes_para1/para2"]),
          ...((P.acts.Act1 || P.acts.Act2 || P.acts.Act3) ? [] : ["Missing Act1/Act2/Act3"]),
          ...(bandsPresent === 12 ? [] : [`Bands not complete (have ${bandsPresent}/12)`]),
        ],
      };

      return res.status(200).json(dbg);
    }

    // Load template
    const pdfBytes = await loadTemplateBytesLocal(tpl);
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

    const pages = pdfDoc.getPages();

    // Guard: expect exactly 8 pages now (but do not hard-crash, just render what exists)
    const pageCount = pages.length;

    // Header name on pages 2–8 (index 1..7)
    const headerName = norm(P.identity.name);
    if (headerName) {
      for (let i = 1; i < pages.length && i <= 7; i++) {
        const pageKey = `p${i + 1}`;
        const box = L?.[pageKey]?.hdrName;
        if (box) drawTextBox(pages[i], font, headerName, box, { maxLines: 1 });
      }
    }

    // Page index mapping (0-based)
    const p1 = pages[0] || null;
    const p2 = pages[1] || null; // header only
    const p3 = pages[2] || null;
    const p4 = pages[3] || null;
    const p5 = pages[4] || null;
    const p6 = pages[5] || null;
    const p7 = pages[6] || null;
    const p8 = pages[7] || null;

    /* Page 1: name + date */
    if (p1 && L.p1) {
      if (L.p1.name && P.identity.name) drawTextBox(p1, font, P.identity.name, L.p1.name, { maxLines: 1 });
      if (L.p1.date && P.identity.date) drawTextBox(p1, font, P.identity.date, L.p1.date, { maxLines: 1 });
    }

    /* Page 3: Exec Summary para1 + para2 */
    if (p3 && L.p3) {
      if (L.p3.exec1) drawTextBox(p3, font, P.exec1, L.p3.exec1, { maxLines: L.p3.exec1.maxLines });
      if (L.p3.exec2) drawTextBox(p3, font, P.exec2, L.p3.exec2, { maxLines: L.p3.exec2.maxLines });
    }

    /* Page 4: CTRL Overview para1 + para2 + CHART (moved here) */
    if (p4 && L.p4) {
      if (L.p4.ov1) drawTextBox(p4, font, P.ov1, L.p4.ov1, { maxLines: L.p4.ov1.maxLines });
      if (L.p4.ov2) drawTextBox(p4, font, P.ov2, L.p4.ov2, { maxLines: L.p4.ov2.maxLines });

      if (L.p4.chart) {
        try {
          await embedRadarFromBands(pdfDoc, p4, L.p4.chart, P.bands || {});
        } catch (e) {
          console.warn("[fill-template V12] Radar chart skipped:", e?.message || String(e));
        }
      }
    }

    /* Page 5: Deep Dive para1 + para2 + Themes para1 + para2 */
    if (p5 && L.p5) {
      if (L.p5.dd1) drawTextBox(p5, font, P.dd1, L.p5.dd1, { maxLines: L.p5.dd1.maxLines });
      if (L.p5.dd2) drawTextBox(p5, font, P.dd2, L.p5.dd2, { maxLines: L.p5.dd2.maxLines });
      if (L.p5.th1) drawTextBox(p5, font, P.th1, L.p5.th1, { maxLines: L.p5.th1.maxLines });
      if (L.p5.th2) drawTextBox(p5, font, P.th2, L.p5.th2, { maxLines: L.p5.th2.maxLines });
    }

    /* Page 6: WorkWith */
    if (p6 && L.p6) {
      if (L.p6.collabC) drawTextBox(p6, font, P.workWith.concealed, L.p6.collabC, { maxLines: L.p6.collabC.maxLines });
      if (L.p6.collabT) drawTextBox(p6, font, P.workWith.triggered, L.p6.collabT, { maxLines: L.p6.collabT.maxLines });
      if (L.p6.collabR) drawTextBox(p6, font, P.workWith.regulated, L.p6.collabR, { maxLines: L.p6.collabR.maxLines });
      if (L.p6.collabL) drawTextBox(p6, font, P.workWith.lead, L.p6.collabL, { maxLines: L.p6.collabL.maxLines });
    }

    /* Page 7: Acts */
    if (p7 && L.p7) {
      if (L.p7.act1) drawTextBox(p7, font, P.acts.Act1, L.p7.act1, { maxLines: L.p7.act1.maxLines });
      if (L.p7.act2) drawTextBox(p7, font, P.acts.Act2, L.p7.act2, { maxLines: L.p7.act2.maxLines });
      if (L.p7.act3) drawTextBox(p7, font, P.acts.Act3, L.p7.act3, { maxLines: L.p7.act3.maxLines });
      if (L.p7.act4) drawTextBox(p7, font, P.acts.Act4, L.p7.act4, { maxLines: L.p7.act4.maxLines });
      if (L.p7.act5) drawTextBox(p7, font, P.acts.Act5, L.p7.act5, { maxLines: L.p7.act5.maxLines });
      if (L.p7.act6) drawTextBox(p7, font, P.acts.Act6, L.p7.act6, { maxLines: L.p7.act6.maxLines });
    }

    // Page 8: header only by default (nothing else to draw)

    const outBytes = await pdfDoc.save();

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Cache-Control", "no-store");
    res.setHeader("X-CTRL-Template", tpl);
    res.setHeader("X-CTRL-PageCount", String(pageCount));
    res.status(200).send(Buffer.from(outBytes));
  } catch (err) {
    console.error("[fill-template V12] CRASH", err);
    res.status(500).json({
      ok: false,
      error: err?.message || String(err),
      stack: err?.stack || null,
    });
  }
}
