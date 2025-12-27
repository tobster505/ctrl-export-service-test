/**
 * CTRL PoC Export Service · fill-template (Starter/PoC flow)
 * Place at: /api/fill-template.js  (ctrl-poc-service)
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

/** TLDR → main → action packer */
function packSection(tldr, main, action) {
  const blocks = [];
  const T = norm(tldr);
  const M = norm(main);
  const A = norm(action);
  if (T) blocks.push(T);
  if (M) blocks.push(M);
  if (A) blocks.push(A);
  return blocks.join("\n\n\n");
}

/* ───────── TL→BL rect helper ───────── */
const rectTLtoBL = (page, box, inset = 0) => {
  const pageH = page.getHeight();
  const x = N(box.x) + inset;
  const w = Math.max(0, N(box.w) - inset * 2);

  // IMPORTANT: use the SAME height value for both h and y-calc
  const hRaw = N(box.h || 400);
  const h = Math.max(0, hRaw - inset * 2);

  const y = pageH - N(box.y) - hRaw + inset;
  return { x, y, w, h };
};

/* ───────── text box helper ───────── */
function drawTextBox(page, font, text, box, opts = {}) {
  if (!page || !font || !box) return;
  const t0 = norm(text);
  if (!t0) return;

  const { x, y, w, h } = rectTLtoBL(page, box, 0);
  const size = N(box.size || opts.size || 12);
  const lineHeight = N(opts.lineHeight || Math.round(size * 1.3));
  const maxLines = N(opts.maxLines || box.maxLines || 999);
  const align = opts.align || box.align || "left";

  // Split words but keep intentional newlines
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
    try {
      return JSON.parse(raw);
    } catch {
      return {};
    }
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

function computeDomAndSecondKeys(P) {
  const raw = P.raw || {};
  const ctrl = raw.ctrl || {};
  const summary = ctrl.summary || {};

  const domKey =
    resolveStateKey(P["p3:dom"]) ||
    resolveStateKey(raw.domState) ||
    resolveStateKey(raw.ctrl?.dominant) ||
    resolveStateKey(summary.domState) ||
    resolveStateKey(ctrl.domState) ||
    "R";

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

  const secondKey = ordered[0]?.[0] || (domKey === "C" ? "T" : "C");
  return { domKey, secondKey };
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

/* ───────── default layout (UPDATED: adds p9 actAnchor box) ───────── */
const DEFAULT_LAYOUT = {
  pages: {
    p1: {
      name: { x: 7, y: 473, w: 500, h: 60, size: 30, align: "center", maxLines: 1 },
      date: { x: 210, y: 600, w: 500, h: 40, size: 25, align: "left", maxLines: 1 },
    },

    // Header name on pages 2–10 (used by handler loop)
    p2:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },

    p3:  {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      domDesc: { x: 25, y: 200, w: 550, h: 900, size: 18, align: "left", maxLines: 38 },
    },

    p4:  {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      spider:  { x: 25, y: 200, w: 550, h: 900, size: 18, align: "left", maxLines: 38 },
    },

    p5:  {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      seqpat:  { x: 25, y: 200, w: 550, h: 900, size: 18, align: "left", maxLines: 38 },
      chart:   { x: 48, y: 462, w: 500, h: 300 },
    },

    p6:  {
      hdrName:   { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      themeExpl: { x: 25, y: 200, w: 550, h: 900, size: 18, align: "left", maxLines: 38 },
    },

    p7:  {
      hdrName:   { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      themesTop: { x: 30,  y: 200, w: 590, h: 900, size: 17, align: "left", maxLines: 42 }, // widened for single-box theme mode
      themesLow: { x: 320, y: 200, w: 300, h: 900, size: 17, align: "left", maxLines: 28 }, // kept for legacy 2-column
    },

    p8:  {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      collabC: { x: 30,  y: 200, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
      collabT: { x: 320, y: 200, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
      collabR: { x: 30,  y: 650, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
      collabL: { x: 320, y: 650, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
    },

    p9:  {
      hdrName:   { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      actAnchor: { x: 25,  y: 200, w: 550, h: 220, size: 20, align: "left", maxLines: 8 }, // NEW: action anchor
    },

    p10: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
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

    const bits = k.split("_"); // L_p3_domDesc_y
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

/* ───────── input normaliser (UPDATED for new payload keys) ───────── */
function normaliseInput(d = {}) {
  const identity = d.identity || {};
  const ctrl = d.ctrl || {};
  const summary = ctrl.summary || {};
  const text = d.text || {};
  const workWith = d.workWith || {};
  const chart = d.chart || {};

  const nameCand =
    (d.person && d.person.fullName) ||
    identity.fullName ||
    identity.name ||
    d["p1:n"] ||
    d.fullName ||
    identity.preferredName ||
    "";

  const dateLbl =
    d.dateLbl ||
    d.date ||
    d["p1:d"] ||
    (d.meta && d.meta.dateLbl) ||
    "";

  const domState =
    d.domState ||
    ctrl.dominant ||
    summary.domState ||
    ctrl.domState ||
    d["p3:dom"] ||
    "";

  const chartUrl = d.chartUrl || chart.url || d["p5:chart"] || "";

  // Prefer the first NON-EMPTY bands object.
  const ctrlBands = (ctrl.bands && typeof ctrl.bands === "object") ? ctrl.bands : null;
  const sumBands  = (summary.bands && typeof summary.bands === "object") ? summary.bands : null;
  const ctrl12    = (summary.ctrl12 && typeof summary.ctrl12 === "object") ? summary.ctrl12 : null;
  const rootBands = (d.bands && typeof d.bands === "object") ? d.bands : null;

  const bandsRaw =
    (ctrlBands && Object.keys(ctrlBands).length ? ctrlBands : null) ||
    (sumBands  && Object.keys(sumBands).length  ? sumBands  : null) ||
    (ctrl12    && Object.keys(ctrl12).length    ? ctrl12    : null) ||
    (rootBands && Object.keys(rootBands).length ? rootBands : null) ||
    {};

  // ---- UPDATED KEY COMPATIBILITY LAYER ----
  // Exec
  const execTLDR = text.execSummary_tldr || "";
  const execMain = text.execSummary || "";
  const execTip  = text.execSummary_tipact || text.tipAction || "";

  // State (Gen2)
  const stateTLDR  = text.state_tldr || text.domState_tldr || "";
  const stateMain  = text.domState || "";
  const bottomMain = text.bottomState || "";
  const stateTip   = text.state_tipact || ""; // optional, new

  // Frequency + sequence
  const freqTLDR = text.frequency_tldr || "";
  const freqMain = text.frequency || "";

  const seqTLDR = text.sequence_tldr || "";
  const seqMain = text.sequence || "";
  const seqTip  = text.sequence_tipact || ""; // new

  // Themes
  // New “single theme block” format:
  const themeTLDR = text.theme_tldr || "";
  const themeMain = text.theme || "";
  const themeTip  = text.theme_tipact || "";

  // Legacy “two box” format:
  const topTLDR = text.themesTop_tldr || themeTLDR || "";
  const lowTLDR = text.themesLow_tldr || "";

  // If new single theme exists, push it into themesTop slot (and leave low empty)
  const themesTopMain = text.themesTop || themeMain || "";
  const themesLowMain = text.themesLow || "";

  // Action Anchor (Gen5)
  const actAnchor = text.act_anchor || text.action_anchor || "";

  const domPlusBottom = bottomMain ? `${norm(stateMain)}\n\n${norm(bottomMain)}` : stateMain;

  return {
    raw: d,
    bands: bandsRaw,
    layout: d.layout || null,

    "p1:n": d["p1:n"] || nameCand || "",
    "p1:d": d["p1:d"] || dateLbl || "",
    "p3:dom": d["p3:dom"] || domState || "",

    // Page 3 (Exec) — TLDR first
    "p3:exec": d["p3:exec"] || execMain || "",
    "p3:tldr": d["p3:tldr"] || execTLDR || "",
    "p3:tip":  d["p3:tip"]  || execTip  || "",

    // Page 4 (State) — TLDR first
    "p4:stateDeep": d["p4:stateDeep"] || domPlusBottom || "",
    "p4:tldr":      d["p4:tldr"]      || stateTLDR || "",
    "p4:action":    d["p4:action"]    || stateTip  || "",

    // Page 5 (Frequency) — TLDR first
    "p5:freq":   d["p5:freq"]   || freqMain || "",
    "p5:tldr":   d["p5:tldr"]   || freqTLDR || "",
    "p5:action": d["p5:action"] || "", // no freq action in your new payload
    "p5:chart":  d["p5:chart"]  || chartUrl || "",

    // Page 6 (Sequence) — TLDR first
    "p6:seq":    d["p6:seq"]    || seqMain || "",
    "p6:tldr":   d["p6:tldr"]   || seqTLDR || "",
    "p6:action": d["p6:action"] || seqTip  || "",

    // Page 7 (Themes) — TLDR first (+ optional theme tip packed into top)
    "p7:themesTop":      d["p7:themesTop"]      || themesTopMain || "",
    "p7:themesLow":      d["p7:themesLow"]      || themesLowMain || "",
    "p7:themesTop_tldr": d["p7:themesTop_tldr"] || topTLDR || "",
    "p7:themesLow_tldr": d["p7:themesLow_tldr"] || lowTLDR || "",
    "p7:themesTop_tip":  themeTip || "",

    // Page 8 (WorkWith)
    "p8:collabC": d["p8:collabC"] || workWith.concealed || "",
    "p8:collabT": d["p8:collabT"] || workWith.triggered || "",
    "p8:collabR": d["p8:collabR"] || workWith.regulated || "",
    "p8:collabL": d["p8:collabL"] || workWith.lead || "",

    // Page 9 (Action Anchor)
    "p9:actAnchor": d["p9:actAnchor"] || actAnchor || "",
  };
}

/* ───────── main handler ───────── */
export default async function handler(req, res) {
  try {
    const url = new URL(req.url, "http://localhost");
    const debug = url.searchParams.get("debug") === "1";

    const payload = await readPayload(req);
    const P = normaliseInput(payload);

    if (debug) {
      return res.status(200).json({
        ok: true,
        where: "fill-template:v3:after_normaliseInput",
        lengths: {
          p3_exec: (P["p3:exec"] || "").length,
          p3_tldr: (P["p3:tldr"] || "").length,
          p3_tip:  (P["p3:tip"]  || "").length,

          p4_main: (P["p4:stateDeep"] || "").length,
          p4_tldr: (P["p4:tldr"] || "").length,
          p4_act:  (P["p4:action"] || "").length,

          p5_main: (P["p5:freq"] || "").length,
          p5_tldr: (P["p5:tldr"] || "").length,

          p6_main: (P["p6:seq"] || "").length,
          p6_tldr: (P["p6:tldr"] || "").length,
          p6_act:  (P["p6:action"] || "").length,

          p7_top: (P["p7:themesTop"] || "").length,
          p7_top_tldr: (P["p7:themesTop_tldr"] || "").length,
          p7_low: (P["p7:themesLow"] || "").length,
          p7_low_tldr: (P["p7:themesLow_tldr"] || "").length,
          p7_top_tip: (P["p7:themesTop_tip"] || "").length,

          p9_anchor: (P["p9:actAnchor"] || "").length,

          bandsKeys: Object.keys(P.bands || {}).length,
        },
        domSecond: safeJson(computeDomAndSecondKeys(P)),
      });
    }

    const { domKey, secondKey } = computeDomAndSecondKeys(P);
    const combo = `${domKey}${secondKey}`;

    const validCombos = new Set(["CT","CL","CR","TC","TR","TL","RC","RT","RL","LC","LR","LT"]);
    const safeCombo = validCombos.has(combo) ? combo : "CT";
    const tpl = `CTRL_PoC_Assessment_Profile_template_${safeCombo}.pdf`;

    const pdfBytes = await loadTemplateBytesLocal(tpl);
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

    // Layout: default + payload overrides + URL overrides
    let layout = mergeLayout(P.layout);
    layout = applyLayoutOverridesFromUrl(layout, url);
    const L = (layout && layout.pages) ? layout.pages : DEFAULT_LAYOUT.pages;

    const pages = pdfDoc.getPages();

    // --- Header name on pages 2–10 (index 1+) ---
    const headerName = norm(P["p1:n"]);
    if (headerName) {
      for (let i = 1; i < pages.length; i++) {
        const pageKey = `p${i + 1}`;
        const box = L?.[pageKey]?.hdrName;
        if (box) drawTextBox(pages[i], font, headerName, box, { maxLines: 1 });
      }
    }

    const p1 = pages[0] || null;
    const p3 = pages[2] || null;
    const p4 = pages[3] || null;
    const p5 = pages[4] || null;
    const p6 = pages[5] || null;
    const p7 = pages[6] || null;
    const p8 = pages[7] || null;
    const p9 = pages[8] || null;

    /* p1: name + date */
    if (p1 && L.p1) {
      if (L.p1.name && P["p1:n"]) drawTextBox(p1, font, P["p1:n"], L.p1.name, { maxLines: 1 });
      if (L.p1.date && P["p1:d"]) drawTextBox(p1, font, P["p1:d"], L.p1.date, { maxLines: 1 });
    }

    /* p3: Exec Summary — TLDR first */
    if (p3 && L.p3?.domDesc) {
      const body = packSection(P["p3:tldr"], P["p3:exec"], P["p3:tip"]);
      drawTextBox(p3, font, body, L.p3.domDesc, { maxLines: L.p3.domDesc.maxLines });
    }

    /* p4: State deep-dive — TLDR first */
    if (p4 && L.p4?.spider) {
      const body = packSection(P["p4:tldr"], P["p4:stateDeep"], P["p4:action"]);
      drawTextBox(p4, font, body, L.p4.spider, { maxLines: L.p4.spider.maxLines });
    }

    /* p5: Frequency — TLDR first + chart */
    if (p5 && L.p5) {
      if (L.p5.seqpat) {
        const body = packSection(P["p5:tldr"], P["p5:freq"], P["p5:action"]);
        drawTextBox(p5, font, body, L.p5.seqpat, { maxLines: L.p5.seqpat.maxLines });
      }

      if (L.p5.chart) {
        const bandsObj = P.bands || {};
        try {
          await embedRadarFromBands(pdfDoc, p5, L.p5.chart, bandsObj);
        } catch (e) {
          console.warn("[fill-template] Radar chart skipped:", e?.message || String(e));
        }
      }
    }

    /* p6: Sequence — TLDR first */
    if (p6 && L.p6?.themeExpl) {
      const body = packSection(P["p6:tldr"], P["p6:seq"], P["p6:action"]);
      drawTextBox(p6, font, body, L.p6.themeExpl, { maxLines: L.p6.themeExpl.maxLines });
    }

    /* p7: Themes — TLDR first (+ optional theme tip packed into the top box) */
    if (p7 && L.p7) {
      if (L.p7.themesTop) {
        const topBody = packSection(P["p7:themesTop_tldr"], P["p7:themesTop"], P["p7:themesTop_tip"]);
        drawTextBox(p7, font, topBody, L.p7.themesTop, { maxLines: L.p7.themesTop.maxLines });
      }
      if (L.p7.themesLow) {
        const lowBody = packSection(P["p7:themesLow_tldr"], P["p7:themesLow"], "");
        drawTextBox(p7, font, lowBody, L.p7.themesLow, { maxLines: L.p7.themesLow.maxLines });
      }
    }

    /* p8: workWith / collaboration */
    if (p8 && L.p8) {
      if (L.p8.collabC && P["p8:collabC"]) drawTextBox(p8, font, P["p8:collabC"], L.p8.collabC, { maxLines: L.p8.collabC.maxLines });
      if (L.p8.collabT && P["p8:collabT"]) drawTextBox(p8, font, P["p8:collabT"], L.p8.collabT, { maxLines: L.p8.collabT.maxLines });
      if (L.p8.collabR && P["p8:collabR"]) drawTextBox(p8, font, P["p8:collabR"], L.p8.collabR, { maxLines: L.p8.collabR.maxLines });
      if (L.p8.collabL && P["p8:collabL"]) drawTextBox(p8, font, P["p8:collabL"], L.p8.collabL, { maxLines: L.p8.collabL.maxLines });
    }

    /* p9: Action Anchor */
    if (p9 && L.p9?.actAnchor && P["p9:actAnchor"]) {
      drawTextBox(p9, font, P["p9:actAnchor"], L.p9.actAnchor, { maxLines: L.p9.actAnchor.maxLines });
    }

    const outBytes = await pdfDoc.save();

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Cache-Control", "no-store");
    res.status(200).send(Buffer.from(outBytes));
  } catch (err) {
    console.error("[fill-template] CRASH", err);
    res.status(500).json({
      ok: false,
      error: err?.message || String(err),
      stack: err?.stack || null,
    });
  }
}
