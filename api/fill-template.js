/**
 * CTRL PoC Export Service · fill-template (V11)
 * Place at: /api/fill-template.js  (ctrl-poc-service)
 *
 * V11 changes vs V9:
 * - Keeps V9 working template loader + rendering pipeline
 * - Updates input mapping to match current Botpress payload
 * - Page 9 now renders 4 actions: text.act_1..act_4 (no more anchor required)
 * - Backwards compatible: if act_anchor exists, it will still render (fallback)
 */

export const config = { runtime: "nodejs" };

import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

/* ───────────── small utils ───────────── */
const S = (v, fb = "") => (v == null ? String(fb) : String(v));
const N = (v, fb = 0) => (Number.isFinite(+v) ? +v : fb);
const norm = (s) => S(s).replace(/\s+/g, " ").trim();

function safeJson(obj) {
  try { return JSON.parse(JSON.stringify(obj)); }
  catch { return { _error: "Could not serialise debug object" }; }
}

function safeJsonParse(s, fb = {}) {
  try { return JSON.parse(String(s || "")); }
  catch { return fb; }
}

function clampStrForFilename(s) {
  return S(s)
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^A-Za-z0-9_\-]/g, "")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function parseDateLabelToYYYYMMDD(dateLbl) {
  const s = S(dateLbl).trim();

  // Accept: "19 Dec 2025" / "19 December 2025"
  const m = s.match(/^(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})$/);
  if (m) {
    const dd = String(m[1]).padStart(2, "0");
    const monRaw = m[2].toLowerCase();
    const yyyy = m[3];

    const map = {
      jan: "01", january: "01",
      feb: "02", february: "02",
      mar: "03", march: "03",
      apr: "04", april: "04",
      may: "05",
      jun: "06", june: "06",
      jul: "07", july: "07",
      aug: "08", august: "08",
      sep: "09", sept: "09", september: "09",
      oct: "10", october: "10",
      nov: "11", november: "11",
      dec: "12", december: "12",
    };

    const mm = map[monRaw] || map[monRaw.slice(0, 3)];
    if (mm) return `${yyyy}-${mm}-${dd}`;
  }

  // Accept: "2025-12-19"
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;

  // Fallback: filename-safe label
  return clampStrForFilename(s || "date");
}

function makeOutputFilename(fullName, dateLbl) {
  const parts = S(fullName).trim().split(/\s+/).filter(Boolean);
  const first = parts[0] || "First";
  const last = parts.length > 1 ? parts[parts.length - 1] : "Surname";
  const datePart = parseDateLabelToYYYYMMDD(dateLbl);

  const fn = clampStrForFilename(first);
  const ln = clampStrForFilename(last);
  return `CTRL_${fn}_${ln}_${datePart}.pdf`;
}

/* ───────── TLDR formatting ───────── */
function formatTLDR(tldr) {
  const s = S(tldr).trim();
  if (!s) return "";

  if (s.includes("\n")) {
    return s
      .split("\n")
      .map((ln) => ln.trim())
      .filter(Boolean)
      .map((ln) => ln.startsWith("•") ? ln : `• ${ln.replace(/^-\s*/, "")}`)
      .join("\n");
  }

  const parts = s.split("•").map((x) => x.trim()).filter(Boolean);
  if (parts.length <= 1) {
    const guess = s.split(/(?<=[.!?])\s+/).map((x) => x.trim()).filter(Boolean);
    return guess.map((x) => (x.startsWith("•") ? x : `• ${x}`)).join("\n");
  }
  return parts.map((x) => `• ${x}`).join("\n");
}

/* ───────── template loader (V9 style) ───────── */
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

/* ───────── drawing helpers ───────── */
function rectTLtoBL(page, box) {
  const H = page.getHeight();
  return { x: box.x, y: H - box.y - box.h, w: box.w, h: box.h };
}

function splitToLines(font, text, maxWidth, size) {
  const words = S(text).replace(/\r/g, "").split(/\s+/).filter(Boolean);
  const lines = [];
  let line = "";

  for (const w of words) {
    const test = line ? `${line} ${w}` : w;
    const width = font.widthOfTextAtSize(test, size);
    if (width <= maxWidth) {
      line = test;
    } else {
      if (line) lines.push(line);
      line = w;
    }
  }
  if (line) lines.push(line);
  return lines;
}

function drawTextBox(page, font, text, box, opts = {}) {
  const txt = S(text || "").replace(/\r/g, "").trim();
  if (!txt) return;

  const { x, y, w, h } = rectTLtoBL(page, box);
  const size = N(box.size ?? 12);
  const lineGap = N(box.lineGap ?? 2);
  const pad = N(box.pad ?? 0);

  const alignRaw = String(opts.align ?? box.align ?? "left").toLowerCase();
  const align = (alignRaw === "centre") ? "center" : alignRaw;

  const maxLines = N(opts.maxLines ?? box.maxLines ?? 50);
  const lineHeight = size + lineGap;

  if (box.bg === true) {
    page.drawRectangle({ x, y, width: w, height: h, color: rgb(1, 1, 1), borderWidth: 0 });
  }

  const paragraphs = txt.split("\n");
  let cursorY = y + h - pad - size;
  let linesUsed = 0;

  for (let p = 0; p < paragraphs.length; p++) {
    const para = paragraphs[p].trim();
    if (!para) continue;

    const lines = splitToLines(font, para, w - pad * 2, size);
    for (const ln of lines) {
      if (linesUsed >= maxLines) return;

      const lw = font.widthOfTextAtSize(ln, size);
      let dx = x + pad;
      if (align === "center") dx = x + (w - lw) / 2;
      if (align === "right")  dx = x + w - pad - lw;

      page.drawText(ln, { x: dx, y: cursorY, size, font, color: rgb(0, 0, 0) });
      cursorY -= lineHeight;
      linesUsed += 1;

      if (cursorY < y + pad) return;
    }

    // Paragraph gap (small)
    if (p < paragraphs.length - 1) {
      cursorY -= Math.max(0, lineGap);
      if (cursorY < y + pad) return;
    }
  }
}

function drawLabelAndBody(page, fontB, font, label, body, box, bodyOpts = {}) {
  const L = norm(label);
  const B = S(body || "").replace(/\r/g, "").trim();
  if (!L && !B) return;

  const { x, y, w, h } = rectTLtoBL(page, box);
  const size = N(box.size ?? 12);
  const lineGap = N(box.lineGap ?? 2);
  const pad = N(box.pad ?? 0);
  const maxLines = N(box.maxLines ?? 50);

  if (box.bg === true) {
    page.drawRectangle({ x, y, width: w, height: h, color: rgb(1, 1, 1), borderWidth: 0 });
  }

  const lineHeight = size + lineGap;
  let cursorY = y + h - pad - size;

  if (L) {
    page.drawText(L, { x: x + pad, y: cursorY, size, font: fontB, color: rgb(0, 0, 0) });
    cursorY -= lineHeight;
  }

  if (B) {
    const bodyBox = {
      x: box.x,
      y: box.y + (L ? lineHeight : 0),
      w: box.w,
      h: box.h - (L ? lineHeight : 0),
      size: box.size,
      align: (box.align ?? "left"),
      maxLines: Math.max(1, maxLines - (L ? 1 : 0)),
      lineGap: box.lineGap,
      pad: box.pad,
      bg: false,
    };
    drawTextBox(page, font, B, bodyBox, { ...bodyOpts, maxLines: bodyBox.maxLines });
  }
}

/* ───────── chart (kept as V9) ───────── */
function makeSpiderChartUrl12(bands) {
  // NOTE: kept exactly as V9 (including bigger labels + rotation)
  const labels = [
    "C_low","C_mid","C_high",
    "T_low","T_mid","T_high",
    "R_low","R_mid","R_high",
    "L_low","L_mid","L_high"
  ];

  const displayLabels = [
    "C (E)","C (D)","C (E)",
    "T (E)","T (D)","T (E)",
    "R (E)","R (D)","R (E)",
    "L (E)","L (D)","L (E)"
  ];

  const data = labels.map((k) => {
    const v = bands?.[k];
    const n = Number(v);
    return Number.isFinite(n) ? Math.max(0, Math.min(1, n)) : 0;
  });

  // v9 colour set (kept)
  const colours = [
    "rgba(180, 180, 180, 0.75)","rgba(180, 180, 180, 0.75)","rgba(180, 180, 180, 0.75)",
    "rgba(255, 140, 0, 0.75)","rgba(255, 140, 0, 0.75)","rgba(255, 140, 0, 0.75)",
    "rgba(0, 140, 255, 0.75)","rgba(0, 140, 255, 0.75)","rgba(0, 140, 255, 0.75)",
    "rgba(160, 0, 255, 0.75)","rgba(160, 0, 255, 0.75)","rgba(160, 0, 255, 0.75)"
  ];

  const startAngle = -Math.PI / 4;

  const cfg = {
    type: "polarArea",
    data: {
      labels: displayLabels,
      datasets: [{
        data,
        backgroundColor: colours,
        borderWidth: 3,
        borderColor: "rgba(0, 0, 0, 0.20)",
      }],
    },
    options: {
      plugins: { legend: { display: false } },
      startAngle,
      scales: {
        r: {
          startAngle,
          min: 0,
          max: 1,
          ticks: { display: false },
          grid: { display: true },
          angleLines: { display: false },
          pointLabels: {
            display: true,
            padding: 14,
            font: { size: 26, weight: "bold" },
            centerPointLabels: true,
          },
        },
      },
    },
  };

  const enc = encodeURIComponent(JSON.stringify(cfg));
  return `https://quickchart.io/chart?c=${enc}&format=png&width=900&height=900&backgroundColor=transparent&version=4`;
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

async function embedRadarFromBandsOrUrl(pdfDoc, page, box, bandsRaw, chartUrl) {
  if (!pdfDoc || !page || !box) return;

  let url = S(chartUrl).trim();
  if (!url) {
    const hasAny =
      bandsRaw && typeof bandsRaw === "object" &&
      Object.values(bandsRaw).some((v) => Number(v) > 0);
    if (!hasAny) return;
    url = makeSpiderChartUrl12(bandsRaw);
  }

  const img = await embedRemoteImage(pdfDoc, url);
  if (!img) return;

  const H = page.getHeight();
  page.drawImage(img, { x: box.x, y: H - box.y - box.h, width: box.w, height: box.h });
}

/* ───────── DEFAULT LAYOUT (V9 coords + V11 Page9 split) ───────── */
const DEFAULT_LAYOUT = {
  pages: {
    p1: {
      name: { x: 60, y: 458, w: 500, h: 60, size: 30, align: "center", maxLines: 1 },
      date: { x: 230, y: 613, w: 500, h: 40, size: 25, align: "left", maxLines: 1 },
    },

    // Header name on pages 2–10
    p2:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p3:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p4:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p5:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p6:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p7:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p8:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p9:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p10: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },

    // Page 3
    p3TLDR: { domDesc: { x: 25, y: 310, w: 550, h: 210, size: 15, align: "left", maxLines: 10 } },
    p3main: { domDesc: { x: 25, y: 515, w: 550, h: 520, size: 15, align: "left", maxLines: 26 } },
    p3act:  { domDesc: { x: 25, y: 700, w: 550, h: 140, size: 15, align: "left", maxLines: 8 } },

    // Page 4
    p4TLDR: { spider: { x: 25, y: 160, w: 550, h: 220, size: 14, align: "left", maxLines: 10 } },
    p4main: { spider: { x: 25, y: 330, w: 550, h: 520, size: 14, align: "left", maxLines: 26 } },
    p4act:  { spider: { x: 25, y: 650, w: 550, h: 140, size: 14, align: "left", maxLines: 8 } },

    // Page 5 (text + chart)
    p5TLDR: { seqpat: { x: 25, y: 170, w: 170, h: 220, size: 15, align: "left", maxLines: 10 } },
    p5main: { seqpat: { x: 25, y: 530, w: 550, h: 520, size: 15, align: "left", maxLines: 28 } },
    p5act:  { seqpat: { x: 25, y: 680, w: 550, h: 140, size: 15, align: "left", maxLines: 8 } },
    p5chart:{ chart:  { x: 250, y: 160, w: 320, h: 320 } },

    // Page 6
    p6TLDR: { themeExpl: { x: 25, y: 160, w: 550, h: 220, size: 16, align: "left", maxLines: 10 } },
    p6main: { themeExpl: { x: 25, y: 300, w: 550, h: 520, size: 16, align: "left", maxLines: 26 } },
    p6act:  { themeExpl: { x: 25, y: 560, w: 550, h: 140, size: 16, align: "left", maxLines: 8 } },

    // Page 7
    p7TLDR: { themesTop: { x: 25, y: 160, w: 550, h: 220, size: 15, align: "left", maxLines: 10 } },
    p7main: { themesTop: { x: 25, y: 350, w: 260, h: 520, size: 15, align: "left", maxLines: 26 } },
    p7act:  { themesTop: { x: 25, y: 640, w: 260, h: 140, size: 15, align: "left", maxLines: 8 } },

    // ✅ themesLow (right column)
    p7: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      themesLow: { x: 320, y: 210, w: 290, h: 900, size: 15, align: "left", maxLines: 28 },
    },

    // Page 8: 2x2 grid
    p8grid: {
      collabC: { x: 30,  y: 300, w: 270, h: 420, size: 14, align: "left", maxLines: 14 },
      collabT: { x: 320, y: 300, w: 260, h: 420, size: 14, align: "left", maxLines: 14 },
      collabR: { x: 30,  y: 575, w: 260, h: 420, size: 14, align: "left", maxLines: 14 },
      collabL: { x: 320, y: 575, w: 260, h: 420, size: 14, align: "left", maxLines: 14 },
    },

    // Page 9: Actions (V11 = 4 boxes; derived from old anchor coords)
    p9: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },

      // legacy box kept (if act_anchor ever sent again)
      actAnchor: { x: 25, y: 300, w: 550, h: 220, size: 16, align: "left", maxLines: 8 },

      // new split boxes
      act1: { x: 25, y: 300, w: 550, h: 50, size: 16, align: "left", maxLines: 2 },
      act2: { x: 25, y: 355, w: 550, h: 50, size: 16, align: "left", maxLines: 2 },
      act3: { x: 25, y: 410, w: 550, h: 50, size: 16, align: "left", maxLines: 2 },
      act4: { x: 25, y: 465, w: 550, h: 50, size: 16, align: "left", maxLines: 2 },
    },
  },
};

function deepMerge(target, source) {
  const out = JSON.parse(JSON.stringify(target || {}));
  if (!source || typeof source !== "object") return out;

  for (const k of Object.keys(source)) {
    const sv = source[k];
    const tv = out[k];
    if (sv && typeof sv === "object" && !Array.isArray(sv)) {
      out[k] = deepMerge(tv && typeof tv === "object" ? tv : {}, sv);
    } else {
      out[k] = sv;
    }
  }
  return out;
}

/* ───────── URL-driven layout overrides ───────── */
function applyLayoutOverridesFromUrl(layout, url) {
  if (!layout || !layout.pages || !url || !url.searchParams) return { layout, applied: [], ignored: [] };

  const allowed = new Set(["x","y","w","h","size","maxLines","align","pad","lineGap","bg"]);
  const pages = layout.pages;

  const applied = [];
  const ignored = [];

  for (const [k, v] of url.searchParams.entries()) {
    if (!k.startsWith("L_")) continue;

    const bits = k.split("_");
    if (bits.length < 4) { ignored.push({ k, v, why: "bits<4" }); continue; }

    const pageKey = bits[1];
    const boxKey  = bits[2];
    const prop    = bits.slice(3).join("_");

    if (!pages[pageKey] || !pages[pageKey][boxKey]) { ignored.push({ k, v, why: "unknown box" }); continue; }
    if (!allowed.has(prop)) { ignored.push({ k, v, why: "prop not allowed" }); continue; }

    if (prop === "align") {
      const a0 = String(v || "").toLowerCase();
      const a = (a0 === "centre") ? "center" : a0;
      if (["left","center","right"].includes(a)) {
        pages[pageKey][boxKey][prop] = a;
        applied.push({ k, v });
      } else {
        ignored.push({ k, v, why: "bad align" });
      }
      continue;
    }

    if (prop === "bg") {
      pages[pageKey][boxKey][prop] = (String(v) === "1" || String(v).toLowerCase() === "true");
      applied.push({ k, v });
      continue;
    }

    const num = Number(v);
    if (!Number.isFinite(num)) { ignored.push({ k, v, why: "not a number" }); continue; }

    pages[pageKey][boxKey][prop] = num;
    applied.push({ k, v });
  }

  return { layout, applied, ignored };
}

/* ───────── input normaliser (updated for your current Botpress payload) ───────── */
function normaliseInput(d = {}) {
  const identity = d.identity || {};
  const text = d.text || {};
  const workWith = d.workWith || {};
  const ctrl = d.ctrl || {};
  const summary = ctrl.summary || {};

  const name =
    (d.person && d.person.fullName) ||
    identity.fullName ||
    identity.name ||
    d["p1:n"] ||
    d.fullName ||
    identity.preferredName ||
    "";

  const email = identity.email || d.email || "";
  const dateLbl = d.dateLbl || d.date || d["p1:d"] || identity.dateLabel || (d.meta && d.meta.dateLbl) || "";

  // bands for chart (prefer ctrl.bands)
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

  // p3
  const p3_tldr = formatTLDR(text.execSummary_tldr || "");
  const p3_main = S(text.execSummary || "");
  const p3_act  = S(text.execSummary_tipact || text.tipAction || "");

  // p4
  const p4_tldr = formatTLDR(text.state_tldr || "");
  const p4_dom  = S(text.domState || "");
  const p4_bot  = S(text.bottomState || "");
  const p4_main = [p4_dom, p4_bot].map((s) => S(s).trim()).filter(Boolean).join("\n\n");
  const p4_act  = S(text.state_tipact || "");

  // p5
  const p5_tldr = formatTLDR(text.frequency_tldr || "");
  const p5_main = S(text.frequency || "");
  const p5_act  = ""; // not currently sent; safe blank

  // p6
  const p6_tldr = formatTLDR(text.sequence_tldr || "");
  const p6_main = S(text.sequence || "");
  const p6_act  = S(text.sequence_tipact || "");

  // p7
  const p7_tldr = formatTLDR(text.theme_tldr || "");
  const p7_main = S(text.theme || "");
  const p7_act  = S(text.theme_tipact || "");
  const p7_low  = S(text.themesLow || text.theme_low || "");

  // p8 (workWith)
  const p8_C = S(workWith.concealed || "");
  const p8_T = S(workWith.triggered || "");
  const p8_R = S(workWith.regulated || "");
  const p8_L = S(workWith.lead || "");

  // p9 (actions) — NEW: act_1..act_4
  const a1 = S(text.act_1 || "");
  const a2 = S(text.act_2 || "");
  const a3 = S(text.act_3 || "");
  const a4 = S(text.act_4 || "");

  // legacy fallback (if anything still sends it)
  const legacyAnchor = S(text.act_anchor || "");

  // Also build a combined fallback string if you ever want one
  const combinedActs = [a1, a2, a3, a4].map((x) => norm(x)).filter(Boolean).map((x) => `• ${x}`).join("\n");

  return {
    // identity
    "p1:n": norm(name),
    "p1:e": norm(email),
    "p1:d": norm(dateLbl),

    // bands for chart
    bands: bandsRaw,

    // p3–p7
    "p3:tldr": p3_tldr,
    "p3:main": p3_main,
    "p3:act":  p3_act,

    "p4:tldr": p4_tldr,
    "p4:main": p4_main,
    "p4:act":  p4_act,

    "p5:tldr": p5_tldr,
    "p5:main": p5_main,
    "p5:act":  p5_act,

    "p6:tldr": p6_tldr,
    "p6:main": p6_main,
    "p6:act":  p6_act,

    "p7:tldr": p7_tldr,
    "p7:main": p7_main,
    "p7:act":  p7_act,
    "p7:low":  p7_low,

    // p8
    "p8:C": p8_C,
    "p8:T": p8_T,
    "p8:R": p8_R,
    "p8:L": p8_L,

    // p9 (new)
    "p9:act1": a1,
    "p9:act2": a2,
    "p9:act3": a3,
    "p9:act4": a4,

    // p9 legacy (only used if you want)
    "p9:anchor": legacyAnchor || combinedActs,
  };
}

/* ───────── main handler ───────── */
export default async function handler(req, res) {
  try {
    const dataParam = req.query?.data || "";
    const decoded = Buffer.from(String(dataParam || ""), "base64").toString("utf-8");
    const payload = safeJsonParse(decoded, {});

    // Template file name (V9 style)
    const tpl = payload?.tpl || payload?.template || payload?.pdfTpl || "template.pdf";
    const pdfBytes = await loadTemplateBytesLocal(tpl);
    const pdfDoc = await PDFDocument.load(pdfBytes);

    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontB = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    // Layout (default + payload + URL overrides)
    const baseLayout = deepMerge(DEFAULT_LAYOUT, payload?.layout || {});
    const url = new URL(req.url, `http://${req.headers.host}`);
    const { layout: layoutWithOverrides, applied, ignored } = applyLayoutOverridesFromUrl(baseLayout, url);

    // Normalised payload-to-keys map
    const P = normaliseInput(payload);

    // Final layout access
    const L = layoutWithOverrides.pages || layoutWithOverrides;

    const pages = pdfDoc.getPages();

    // p1
    if (pages[0] && L.p1) {
      if (L.p1.name && P["p1:n"]) drawTextBox(pages[0], fontB, P["p1:n"], L.p1.name, { maxLines: 1 });
      if (L.p1.date && P["p1:d"]) drawTextBox(pages[0], font,  P["p1:d"], L.p1.date, { maxLines: 1 });
    }

    // headers p2–p10
    const headerName = norm(P["p1:n"]);
    if (headerName) {
      for (let i = 1; i < pages.length; i++) {
        const pk = `p${i + 1}`;
        const box = L?.[pk]?.hdrName;
        if (box) drawTextBox(pages[i], font, headerName, box, { maxLines: 1 });
      }
    }

    const p3 = pages[2] || null;
    const p4 = pages[3] || null;
    const p5 = pages[4] || null;
    const p6 = pages[5] || null;
    const p7 = pages[6] || null;
    const p8 = pages[7] || null;
    const p9 = pages[8] || null;

    if (p3) {
      if (L.p3TLDR?.domDesc) drawLabelAndBody(p3, fontB, font, "TLDR",  P["p3:tldr"], L.p3TLDR.domDesc);
      if (L.p3main?.domDesc) drawLabelAndBody(p3, fontB, font, "",      P["p3:main"], L.p3main.domDesc);
      if (L.p3act?.domDesc)  drawLabelAndBody(p3, fontB, font, "Action", P["p3:act"],  L.p3act.domDesc);
    }

    if (p4) {
      if (L.p4TLDR?.spider) drawLabelAndBody(p4, fontB, font, "TLDR",  P["p4:tldr"], L.p4TLDR.spider);
      if (L.p4main?.spider) drawLabelAndBody(p4, fontB, font, "",      P["p4:main"], L.p4main.spider);
      if (L.p4act?.spider)  drawLabelAndBody(p4, fontB, font, "Action", P["p4:act"],  L.p4act.spider);
    }

    if (p5) {
      if (L.p5TLDR?.seqpat) drawLabelAndBody(p5, fontB, font, "TLDR",  P["p5:tldr"], L.p5TLDR.seqpat);
      if (L.p5main?.seqpat) drawLabelAndBody(p5, fontB, font, "",      P["p5:main"], L.p5main.seqpat);
      if (L.p5act?.seqpat)  drawLabelAndBody(p5, fontB, font, "Action", P["p5:act"],  L.p5act.seqpat);

      if (L.p5chart?.chart) {
        try {
          await embedRadarFromBandsOrUrl(pdfDoc, p5, L.p5chart.chart, P.bands || {}, payload?.spiderChartUrl || payload?.chart?.spiderUrl || "");
        } catch (e) {
          console.warn("[fill-template:v11] Radar chart skipped:", e?.message || String(e));
        }
      }
    }

    if (p6) {
      if (L.p6TLDR?.themeExpl) drawLabelAndBody(p6, fontB, font, "TLDR",  P["p6:tldr"], L.p6TLDR.themeExpl);
      if (L.p6main?.themeExpl) drawLabelAndBody(p6, fontB, font, "",      P["p6:main"], L.p6main.themeExpl);
      if (L.p6act?.themeExpl)  drawLabelAndBody(p6, fontB, font, "Action", P["p6:act"],  L.p6act.themeExpl);
    }

    if (p7) {
      if (L.p7TLDR?.themesTop) drawLabelAndBody(p7, fontB, font, "TLDR",  P["p7:tldr"], L.p7TLDR.themesTop);
      if (L.p7main?.themesTop) drawLabelAndBody(p7, fontB, font, "",      P["p7:main"], L.p7main.themesTop);
      if (L.p7act?.themesTop)  drawLabelAndBody(p7, fontB, font, "Action", P["p7:act"],  L.p7act.themesTop);

      if (L.p7?.themesLow && P["p7:low"]) {
        drawTextBox(p7, font, P["p7:low"], L.p7.themesLow, { maxLines: L.p7.themesLow.maxLines });
      }
    }

    if (p8 && L.p8grid) {
      if (L.p8grid.collabC && P["p8:C"]) drawTextBox(p8, font, P["p8:C"], L.p8grid.collabC, { maxLines: L.p8grid.collabC.maxLines });
      if (L.p8grid.collabT && P["p8:T"]) drawTextBox(p8, font, P["p8:T"], L.p8grid.collabT, { maxLines: L.p8grid.collabT.maxLines });
      if (L.p8grid.collabR && P["p8:R"]) drawTextBox(p8, font, P["p8:R"], L.p8grid.collabR, { maxLines: L.p8grid.collabR.maxLines });
      if (L.p8grid.collabL && P["p8:L"]) drawTextBox(p8, font, P["p8:L"], L.p8grid.collabL, { maxLines: L.p8grid.collabL.maxLines });
    }

    // Page 9 — NEW: 4 actions
    if (p9 && L.p9) {
      if (L.p9.act1 && P["p9:act1"]) drawTextBox(p9, font, P["p9:act1"], L.p9.act1, { maxLines: L.p9.act1.maxLines });
      if (L.p9.act2 && P["p9:act2"]) drawTextBox(p9, font, P["p9:act2"], L.p9.act2, { maxLines: L.p9.act2.maxLines });
      if (L.p9.act3 && P["p9:act3"]) drawTextBox(p9, font, P["p9:act3"], L.p9.act3, { maxLines: L.p9.act3.maxLines });
      if (L.p9.act4 && P["p9:act4"]) drawTextBox(p9, font, P["p9:act4"], L.p9.act4, { maxLines: L.p9.act4.maxLines });

      // fallback: if no individual acts, render legacy combined anchor (won’t harm)
      const noneActs = !P["p9:act1"] && !P["p9:act2"] && !P["p9:act3"] && !P["p9:act4"];
      if (noneActs && L.p9.actAnchor && P["p9:anchor"]) {
        drawTextBox(p9, font, P["p9:anchor"], L.p9.actAnchor, { maxLines: L.p9.actAnchor.maxLines });
      }
    }

    // Diagnostics (kept lightweight; won’t break rendering)
    const missing = { identity: [], text: [], workWith: [], actions: [] };

    if (!P["p1:n"]) missing.identity.push("identity.fullName");
    if (!P["p1:e"]) missing.identity.push("identity.email");
    if (!P["p1:d"]) missing.identity.push("identity.dateLabel/dateLbl");

    if (!P["p3:tldr"]) missing.text.push("text.execSummary_tldr");
    if (!P["p3:main"]) missing.text.push("text.execSummary");
    if (!P["p3:act"])  missing.text.push("text.execSummary_tipact");

    if (!P["p4:tldr"]) missing.text.push("text.state_tldr");
    if (!P["p4:main"]) missing.text.push("text.domState+bottomState");
    if (!P["p4:act"])  missing.text.push("text.state_tipact");

    if (!P["p5:main"]) missing.text.push("text.frequency");
    if (!P["p6:main"]) missing.text.push("text.sequence");
    if (!P["p7:main"]) missing.text.push("text.theme");

    if (!P["p8:C"]) missing.workWith.push("workWith.concealed");
    if (!P["p8:T"]) missing.workWith.push("workWith.triggered");
    if (!P["p8:R"]) missing.workWith.push("workWith.regulated");
    if (!P["p8:L"]) missing.workWith.push("workWith.lead");

    if (!P["p9:act1"]) missing.actions.push("text.act_1");
    if (!P["p9:act2"]) missing.actions.push("text.act_2");
    if (!P["p9:act3"]) missing.actions.push("text.act_3");
    if (!P["p9:act4"]) missing.actions.push("text.act_4");

    const outBytes = await pdfDoc.save();

    const outName = makeOutputFilename(P["p1:n"], P["p1:d"]);
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Cache-Control", "no-store");
    res.setHeader("Content-Disposition", `inline; filename="${outName}"`);

    // Optional debug headers (safe + tiny)
    res.setHeader("x-ctrl-layout-overrides-applied", String(applied.length));
    res.setHeader("x-ctrl-layout-overrides-ignored", String(ignored.length));
    res.setHeader("x-ctrl-missing-count", String(
      missing.identity.length + missing.text.length + missing.workWith.length + missing.actions.length
    ));

    res.status(200).send(Buffer.from(outBytes));
  } catch (err) {
    console.error("[fill-template:v11] CRASH", err);
    res.status(500).json({
      ok: false,
      error: err?.message || String(err),
      stack: err?.stack || null,
    });
  }
}
