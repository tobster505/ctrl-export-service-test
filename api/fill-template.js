/**
 * CTRL PoC Export Service · fill-template (Starter/PoC flow)
 * Place at: /api/fill-template.js  (ctrl-poc-service)
 */
export const config = { runtime: "nodejs" };

import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

/* ───────────── utilities ───────────── */
const S = (v, fb = "") => (v == null ? String(fb) : String(v));
const N = (v, fb = 0) => (Number.isFinite(+v) ? +v : fb);
const clamp = (v, lo, hi) => (v < lo ? lo : v > hi ? hi : v);
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

/* brand colour */
const BRAND = { r: 0.72, g: 0.06, b: 0.44 };

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
    } catch (e) {
      return {};
    }
  }

  const url = new URL(req.url, "http://localhost");
  const dataB64 = url.searchParams.get("data") || "";
  if (!dataB64) return {};

  try {
    const raw = Buffer.from(dataB64, "base64").toString("utf8");
    return JSON.parse(raw);
  } catch (e) {
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
  if (typeof fetch !== "function") {
    throw new Error("fetch is not available in this runtime. Ensure Node 18+ on Vercel.");
  }

  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to fetch chart: ${res.status} ${res.statusText}`);

  const buf = new Uint8Array(await res.arrayBuffer());
  const sig = String.fromCharCode(buf[0], buf[1], buf[2], buf[3] || 0);

  if (sig.startsWith("\x89PNG")) return await pdfDoc.embedPng(buf);
  if (sig.startsWith("\xff\xd8")) return await pdfDoc.embedJpg(buf);

  try {
    return await pdfDoc.embedPng(buf);
  } catch {
    return await pdfDoc.embedJpg(buf);
  }
}

/* FIX 1: clean embedRadarFromBands (no duplicate zombie code) */
async function embedRadarFromBands(pdfDoc, page, box, bandsRaw, debug = false) {
  if (!pdfDoc || !page || !box || !bandsRaw) return;

  if (debug) {
    console.log("[fill-template] chart:enter", {
      hasBandsRaw: !!bandsRaw,
      keys: bandsRaw && typeof bandsRaw === "object" ? Object.keys(bandsRaw).length : 0,
      box,
    });
  }

  const hasAny =
    bandsRaw && typeof bandsRaw === "object" &&
    Object.values(bandsRaw).some((v) => Number(v) > 0);

  if (!hasAny) {
    if (debug) console.log("[fill-template] chart:exit:no_values");
    return;
  }

  const url = makeSpiderChartUrl12(bandsRaw);
  if (debug) console.log("[fill-template] chart:url", { len: url.length, preview: url.slice(0, 120) });

  const img = await embedRemoteImage(pdfDoc, url);
  if (!img) {
    if (debug) console.log("[fill-template] chart:exit:no_img");
    return;
  }

  const H = page.getHeight();
  const { x, y, w, h } = box;

  if (debug) console.log("[fill-template] chart:draw", { x, y, w, h, yBL: H - y - h, pageH: H });

  page.drawImage(img, { x, y: H - y - h, width: w, height: h });

  if (debug) console.log("[fill-template] chart:done");
}

/* ───────── default layout (complete) ───────── */
const DEFAULT_LAYOUT = {
  pages: {
    p1: {
      name: { x: 7, y: 473, w: 500, h: 60, size: 30, align: "center", maxLines: 1 },
      date: { x: 210, y: 600, w: 500, h: 40, size: 25, align: "left", maxLines: 1 },
    },

    // Header name on pages 2–10 (used by handler loop)
    p2:  { hdrName: { x: 30, y: 55, w: 560, h: 30, size: 16, align: "left", maxLines: 1 } },

    p3:  {
      hdrName: { x: 30, y: 55, w: 560, h: 30, size: 16, align: "left", maxLines: 1 },
      domDesc: { x: 25, y: 685, w: 550, h: 420, size: 18, align: "left", maxLines: 20 },
    },

    p4:  {
      hdrName: { x: 30, y: 55, w: 560, h: 30, size: 16, align: "left", maxLines: 1 },
      spider:  { x: 25, y: 347, w: 550, h: 420, size: 18, align: "left", maxLines: 20 },
    },

    p5:  {
      hdrName: { x: 30, y: 55, w: 560, h: 30, size: 16, align: "left", maxLines: 1 },
      seqpat:  { x: 25, y: 347, w: 550, h: 420, size: 18, align: "left", maxLines: 20 },
      chart:   { x: 48, y: 462, w: 500, h: 300 },
    },

    p6:  {
      hdrName:   { x: 30, y: 55, w: 560, h: 30, size: 16, align: "left", maxLines: 1 },
      themeExpl: { x: 25, y: 347, w: 550, h: 420, size: 18, align: "left", maxLines: 20 },
    },

    p7:  {
      hdrName:   { x: 30, y: 55, w: 560, h: 30, size: 16, align: "left", maxLines: 1 },
      themesTop: { x: 30,  y: 530, w: 300, h: 420, size: 17, align: "left", maxLines: 12 },
      themesLow: { x: 320, y: 530, w: 300, h: 420, size: 17, align: "left", maxLines: 12 },
    },

    p8:  {
      hdrName: { x: 30, y: 55, w: 560, h: 30, size: 16, align: "left", maxLines: 1 },
      collabC: { x: 30,  y: 530, w: 300, h: 420, size: 17, align: "left", maxLines: 12 },
      collabT: { x: 320, y: 530, w: 300, h: 420, size: 17, align: "left", maxLines: 12 },
      collabR: { x: 30,  y: 960, w: 300, h: 420, size: 17, align: "left", maxLines: 12 },
      collabL: { x: 320, y: 960, w: 300, h: 420, size: 17, align: "left", maxLines: 12 },
    },

    p9:  { hdrName: { x: 30, y: 55, w: 560, h: 30, size: 16, align: "left", maxLines: 1 } },
    p10: { hdrName: { x: 30, y: 55, w: 560, h: 30, size: 16, align: "left", maxLines: 1 } },
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

/* ───────── input normaliser ───────── */
function normaliseInput(d = {}) {
  const identity = d.identity || {};
  const ctrl = d.ctrl || {};
  const summary = ctrl.summary || {};
  const text = d.text || {};
  const workWith = d.workWith || {};
  const actionsObj = d.actions || {};
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
    summary.domState ||
    ctrl.domState ||
    d["p3:dom"] ||
    "";

  const tldrLines =
    (Array.isArray(text.tldr) && text.tldr) ||
    (Array.isArray(d.tldr) && d.tldr) ||
    [];

  const chartUrl = d.chartUrl || chart.url || d["p5:chart"] || "";

  return {
    raw: d,
    identity,
    ctrl,
    summary,
    text,
    workWith,
    chartUrl,
    layout: d.layout || null,

    bands: ctrl.bands || summary.bands || summary.ctrl12 || d.bands || {},

    "p1:n": d["p1:n"] || nameCand || "",
    "p1:d": d["p1:d"] || dateLbl || "",

    "p3:dom": d["p3:dom"] || domState || "",
    "p3:exec": d["p3:exec"] || text.execSummary || "",
    "p3:tldr1": d["p3:tldr1"] || tldrLines[0] || "",
    "p3:tldr2": d["p3:tldr2"] || tldrLines[1] || "",
    "p3:tldr3": d["p3:tldr3"] || tldrLines[2] || "",
    "p3:tldr4": d["p3:tldr4"] || tldrLines[3] || "",
    "p3:tldr5": d["p3:tldr5"] || tldrLines[4] || "",
    "p3:tip": d["p3:tip"] || text.tipAction || "",

    "p4:stateDeep": d["p4:stateDeep"] || text.stateSubInterpretation || "",
    "p4:tldr": d["p4:tldr"] || text.p4tldr || "",
    "p4:action": d["p4:action"] || text.p4action || "",

    "p5:freq": d["p5:freq"] || text.frequency || "",
    "p5:chart": d["p5:chart"] || chartUrl || "",
    "p5:tldr": d["p5:tldr"] || text.p5tldr || "",
    "p5:action": d["p5:action"] || text.p5action || "",

    "p6:seq": d["p6:seq"] || text.sequence || "",
    "p6:tldr": d["p6:tldr"] || text.p6tldr || "",
    "p6:action": d["p6:action"] || text.p6action || "",

    "p7:themesTop": d["p7:themesTop"] || text.themesTop || text.themes_top || "",
    "p7:themesLow": d["p7:themesLow"] || text.themesLow || text.themes_low || "",

    "p8:collabC": d["p8:collabC"] || workWith.concealed || "",
    "p8:collabT": d["p8:collabT"] || workWith.triggered || "",
    "p8:collabR": d["p8:collabR"] || workWith.regulated || "",
    "p8:collabL": d["p8:collabL"] || workWith.lead || "",
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
      // FIX 3: add an explicit “bands source” debug line
      const raw = (payload && typeof payload === "object") ? payload : {};
      const rawCtrl = raw.ctrl || {};
      const rawSummary = (rawCtrl && rawCtrl.summary) ? rawCtrl.summary : {};

      const bandsFrom = {
        "payload.ctrl.bands": rawCtrl && typeof rawCtrl.bands === "object" ? Object.keys(rawCtrl.bands).length : 0,
        "payload.ctrl.summary.bands": rawSummary && typeof rawSummary.bands === "object" ? Object.keys(rawSummary.bands).length : 0,
        "payload.ctrl.summary.ctrl12": rawSummary && typeof rawSummary.ctrl12 === "object" ? Object.keys(rawSummary.ctrl12).length : 0,
        "payload.bands": raw && typeof raw.bands === "object" ? Object.keys(raw.bands).length : 0,
        "P.bands": P.bands && typeof P.bands === "object" ? Object.keys(P.bands).length : 0,
      };

      return res.status(200).json({
        ok: true,
        where: "fill-template:after_normaliseInput",
        gotDataParam: url.searchParams.has("data"),
        payloadTopKeys: payload && typeof payload === "object" ? Object.keys(payload).slice(0, 25) : [],
        lengths: {
          p7themesTop: (P["p7:themesTop"] || "").length,
          p7themesLow: (P["p7:themesLow"] || "").length,
          p8c: (P["p8:collabC"] || "").length,
          p8t: (P["p8:collabT"] || "").length,
          p8r: (P["p8:collabR"] || "").length,
          p8l: (P["p8:collabL"] || "").length,
          bandsKeys: Object.keys(P.bands || {}).length,
        },
        // NEW: chart bands source lens
        bandsFrom,
        samples: {
          themesTop: (P["p7:themesTop"] || "").slice(0, 140),
          collabC: (P["p8:collabC"] || "").slice(0, 140),
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

    const layout = mergeLayout(P.layout);
    const L = (layout && layout.pages) ? layout.pages : DEFAULT_LAYOUT.pages;

    const pages = pdfDoc.getPages();

    // --- Header name on pages 2–10 (index 1+) ---
    const headerName = norm(P["p1:n"]);
    if (headerName) {
      for (let i = 1; i < pages.length; i++) {
        const pageKey = `p${i + 1}`; // page index 1 => p2
        const box = L?.[pageKey]?.hdrName;
        if (box) drawTextBox(pages[i], font, headerName, box, { maxLines: box.maxLines || 1 });
      }
    }

    // SAFETY: do not assume page count
    const p1 = pages[0] || null;
    const p3 = pages[2] || null;
    const p4 = pages[3] || null;
    const p5 = pages[4] || null;
    const p6 = pages[5] || null;
    const p7 = pages[6] || null;
    const p8 = pages[7] || null;

    /* p1: name + date */
    if (p1 && L.p1) {
      if (L.p1.name && P["p1:n"]) drawTextBox(p1, font, P["p1:n"], L.p1.name, { maxLines: 1 });
      if (L.p1.date && P["p1:d"]) drawTextBox(p1, font, P["p1:d"], L.p1.date, { maxLines: 1 });
    }

    /* p3: TLDR → main → action */
    if (p3 && L.p3?.domDesc) {
      const exec = norm(P["p3:exec"]);
      const tldrs = [P["p3:tldr1"], P["p3:tldr2"], P["p3:tldr3"], P["p3:tldr4"], P["p3:tldr5"]]
        .map(norm)
        .filter(Boolean);
      const tip = norm(P["p3:tip"]);

      const blocks = [];
      if (tldrs.length) blocks.push(tldrs.join("\n\n"));
      if (exec) blocks.push(exec);
      if (tip) blocks.push(tip);

      drawTextBox(p3, font, blocks.join("\n\n\n"), L.p3.domDesc, { maxLines: L.p3.domDesc.maxLines });
    }

    /* p4: TLDR → main → action */
    if (p4 && L.p4?.spider) {
      const body = packSection(P["p4:tldr"], P["p4:stateDeep"], P["p4:action"]);
      drawTextBox(p4, font, body, L.p4.spider, { maxLines: L.p4.spider.maxLines });
    }

    /* p5: TLDR → main → action + chart */
    if (p5 && L.p5) {
      if (L.p5.seqpat) {
        const body = packSection(P["p5:tldr"], P["p5:freq"], P["p5:action"]);
        drawTextBox(p5, font, body, L.p5.seqpat, { maxLines: L.p5.seqpat.maxLines });
      }

      /* FIX 2: keep the try/catch inside the chart block, add chart debug */
      if (L.p5.chart) {
        const bandsObj = P.bands || {};
        const bandKeys = (bandsObj && typeof bandsObj === "object") ? Object.keys(bandsObj) : [];
        const labels = ["C_low","C_mid","C_high","T_low","T_mid","T_high","R_low","R_mid","R_high","L_low","L_mid","L_high"];
        const vals = labels.map((k) => Number(bandsObj?.[k] || 0));
        const hasAny = vals.some((v) => v > 0);
        const maxVal = Math.max(...vals, 0);
        const sumVal = vals.reduce((a, b) => a + b, 0);

        // 2nd debug line regarding the chart (runtime logging)
        console.log("[fill-template] chart:bandsMeta", {
          keys: bandKeys.length,
          sampleKeys: bandKeys.slice(0, 12),
          hasAny,
          maxVal,
          sum: sumVal,
          first6: vals.slice(0, 6),
          last6: vals.slice(6),
          box: L.p5.chart,
          pageH: p5.getHeight(),
          pageW: p5.getWidth(),
        });

        try {
          await embedRadarFromBands(pdfDoc, p5, L.p5.chart, bandsObj, false);
        } catch (e) {
          console.warn("[fill-template] Radar chart skipped:", e?.message || String(e));
        }
      }
    }

    /* p6: TLDR → main → action */
    if (p6 && L.p6?.themeExpl) {
      const body = packSection(P["p6:tldr"], P["p6:seq"], P["p6:action"]);
      drawTextBox(p6, font, body, L.p6.themeExpl, { maxLines: L.p6.themeExpl.maxLines });
    }

    /* p7: themes (Top + Low) */
    if (p7 && L.p7) {
      if (L.p7.themesTop && P["p7:themesTop"]) {
        drawTextBox(p7, font, P["p7:themesTop"], L.p7.themesTop, { maxLines: L.p7.themesTop.maxLines });
      }
      if (L.p7.themesLow && P["p7:themesLow"]) {
        drawTextBox(p7, font, P["p7:themesLow"], L.p7.themesLow, { maxLines: L.p7.themesLow.maxLines });
      }
    }

    /* p8: workWith / collaboration */
    if (p8 && L.p8) {
      if (L.p8.collabC && P["p8:collabC"]) drawTextBox(p8, font, P["p8:collabC"], L.p8.collabC, { maxLines: L.p8.collabC.maxLines });
      if (L.p8.collabT && P["p8:collabT"]) drawTextBox(p8, font, P["p8:collabT"], L.p8.collabT, { maxLines: L.p8.collabT.maxLines });
      if (L.p8.collabR && P["p8:collabR"]) drawTextBox(p8, font, P["p8:collabR"], L.p8.collabR, { maxLines: L.p8.collabR.maxLines });
      if (L.p8.collabL && P["p8:collabL"]) drawTextBox(p8, font, P["p8:collabL"], L.p8.collabL, { maxLines: L.p8.collabL.maxLines });
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
