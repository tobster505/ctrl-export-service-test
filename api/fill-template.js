/**
 * CTRL PoC Export Service · fill-template (V8)
 * Place at: /api/fill-template.js  (ctrl-poc-service)
 *
 * Fixes in V8:
 * 1) Output filename: PoC_Profile_{Firstname}_{Surname}_{Date}.pdf
 * 2) URL coordinate overrides (full) supported for ALL boxes (TLDR/main/action, etc.)
 * 3) Header on pages 2–10 becomes just "FullName" (by masking template header text)
 * 4) Pages 3–7 render: TLDR (bold) + bullet lines + Main + Action (bold) + action paragraph
 * 5) Default align-left and URL overrideable width (w) so text stops falling off page
 * 6) Spider/radar chart embedded on Page 5 (from bands if chartUrl missing)
 * 7) Page 8 WorkWith blocks are 2x2 grid (not a single row)
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

  return `PoC_Profile_${fn}_${ln}_${datePart}.pdf`;
}

/* ───────── TL→BL rect helper ───────── */
const rectTLtoBL = (page, box) => {
  const pageH = page.getHeight();
  const x = N(box.x);
  const w = Math.max(0, N(box.w));
  const h = Math.max(0, N(box.h));
  const y = pageH - N(box.y) - h;
  return { x, y, w, h };
};

/* ───────── text wrapping + drawing ───────── */
function wrapText(font, text, size, w) {
  const raw = S(text);
  const paragraphs = raw.split("\n");
  const lines = [];

  for (const para of paragraphs) {
    const words = para.split(/\s+/).filter(Boolean);
    let line = "";

    if (!words.length) {
      lines.push("");
      continue;
    }

    for (const word of words) {
      const test = line ? `${line} ${word}` : word;
      const width = font.widthOfTextAtSize(test, size);
      if (width <= w) {
        line = test;
      } else {
        if (line) lines.push(line);
        line = word;
      }
    }
    if (line) lines.push(line);
  }

  // trim trailing blanks
  while (lines.length && lines[lines.length - 1] === "") lines.pop();
  return lines;
}

function drawTextBox(page, font, text, box, opts = {}) {
  if (!page || !font || !box) return;
  const t0 = S(text);
  if (!t0.trim()) return;

  const { x, y, w, h } = rectTLtoBL(page, box);
  const size = N(opts.size ?? box.size ?? 12);
  const lineGap = N(opts.lineGap ?? box.lineGap ?? 2);
  const maxLines = N(opts.maxLines ?? box.maxLines ?? 999);
  const alignRaw = String(opts.align ?? box.align ?? "left").toLowerCase();
  const align = (alignRaw === "centre") ? "center" : alignRaw;

  const pad = N(opts.pad ?? box.pad ?? 0);

  // Optional: mask background (useful to overwrite template header text)
  if (opts.bg === true || box.bg === true) {
    page.drawRectangle({
      x,
      y,
      width: w,
      height: h,
      color: rgb(1, 1, 1),
      borderWidth: 0,
    });
  }

  const innerW = Math.max(0, w - pad * 2);
  const lines = wrapText(font, t0.replace(/\r/g, ""), size, innerW).slice(0, maxLines);

  const lineHeight = size + lineGap;
  let cursorY = y + h - pad - size;

  for (let i = 0; i < lines.length; i++) {
    const ln = lines[i];
    if (cursorY < y + pad) break;

    let dx = x + pad;
    if (align !== "left") {
      const lw = font.widthOfTextAtSize(ln, size);
      if (align === "center") dx = x + (w - lw) / 2;
      if (align === "right") dx = x + w - pad - lw;
    }

    page.drawText(ln, { x: dx, y: cursorY, size, font, color: rgb(0, 0, 0) });
    cursorY -= lineHeight;
  }
}

/* Bold label (first line), then wrapped body below */
function drawLabelAndBody(page, fontB, font, label, body, box, bodyOpts = {}) {
  const L = norm(label);
  const B = S(body || "").replace(/\r/g, "").trim();
  if (!L && !B) return;

  const { x, y, w, h } = rectTLtoBL(page, box);
  const size = N(box.size ?? 12);
  const lineGap = N(box.lineGap ?? 2);
  const pad = N(box.pad ?? 0);
  const maxLines = N(box.maxLines ?? 50);

  // optional mask
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
    // body box beneath label
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

/* ───────── TLDR formatting ───────── */
function formatTLDR(tldr) {
  const s = S(tldr).trim();
  if (!s) return "";

  // If already multi-line, keep, but normalise bullets a bit
  if (s.includes("\n")) {
    return s
      .split("\n")
      .map((ln) => ln.trim())
      .filter(Boolean)
      .map((ln) => ln.startsWith("•") ? ln : `• ${ln.replace(/^-\s*/, "")}`)
      .join("\n");
  }

  // Split "• a • b • c" into lines
  const parts = s.split("•").map((x) => x.trim()).filter(Boolean);
  if (parts.length <= 1) {
    // fallback: split by sentence-ish
    const guess = s.split(/(?<=[.!?])\s+/).map((x) => x.trim()).filter(Boolean);
    return guess.map((x) => (x.startsWith("•") ? x : `• ${x}`)).join("\n");
  }
  return parts.map((x) => `• ${x}`).join("\n");
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

function computeDomAndSecondKeys(P) {
  const raw = P.raw || {};
  const ctrl = raw.ctrl || {};
  const summary = (ctrl.summary || raw.ctrl?.summary || {}) || {};

  const domKey =
    resolveStateKey(P.domKey) ||
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
  return { domKey, secondKey, templateKey: `${domKey}${secondKey}` };
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

async function embedRadarFromBandsOrUrl(pdfDoc, page, box, bandsRaw, chartUrl) {
  if (!pdfDoc || !page || !box) return;

  // Prefer explicit chart URL if provided
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

/* ───────── default layout (supports your URL override scheme) ───────── */
const DEFAULT_LAYOUT = {
  pages: {
    p1: {
      name: { x: 7, y: 473, w: 500, h: 60, size: 30, align: "center", maxLines: 1 },
      date: { x: 210, y: 600, w: 500, h: 40, size: 25, align: "left", maxLines: 1 },
    },

    // Header name on pages 2–10 (mask template header and write just the name)
    p2:  { hdrName: { x: 360, y: 44, w: 520, h: 34, size: 13, align: "left", maxLines: 1} },
    p3:  { hdrName: { x: 360, y: 44, w: 520, h: 34, size: 13, align: "left", maxLines: 1} },
    p4:  { hdrName: { x: 360, y: 44, w: 520, h: 34, size: 13, align: "left", maxLines: 1} },
    p5:  { hdrName: { x: 360, y: 44, w: 520, h: 34, size: 13, align: "left", maxLines: 1} },
    p6:  { hdrName: { x: 360, y: 44, w: 520, h: 34, size: 13, align: "left", maxLines: 1} },
    p7:  { hdrName: { x: 360, y: 44, w: 520, h: 34, size: 13, align: "left", maxLines: 1} },
    p8:  { hdrName: { x: 360, y: 44, w: 520, h: 34, size: 13, align: "left", maxLines: 1} },
    p9:  { hdrName: { x: 360, y: 44, w: 520, h: 34, size: 13, align: "left", maxLines: 1} },
    p10: { hdrName: { x: 360, y: 44, w: 520, h: 34, size: 13, align: "left", maxLines: 1} },

    // Pages 3–7: split boxes (TLDR / main / action) to match your coordinate file format
    p3TLDR: { domDesc: { x: 25, y: 310, w: 550, h: 210, size: 15, align: "left", maxLines: 10 } },
    p3main: { domDesc: { x: 25, y: 450, w: 550, h: 520, size: 15, align: "left", maxLines: 26 } },
    p3act:  { domDesc: { x: 25, y: 700, w: 550, h: 140, size: 15, align: "left", maxLines: 8 } },

    p4TLDR: { spider: { x: 25, y: 160, w: 550, h: 220, size: 15, align: "left", maxLines: 10 } },
    p4main: { spider: { x: 25, y: 340, w: 550, h: 520, size: 15, align: "left", maxLines: 26 } },
    p4act:  { spider: { x: 25, y: 630, w: 550, h: 140, size: 15, align: "left", maxLines: 8 } },

    p5TLDR: { seqpat: { x: 25, y: 170, w: 210, h: 220, size: 14, align: "left", maxLines: 10 } },
    p5main: { seqpat: { x: 25, y: 340, w: 210, h: 430, size: 14, align: "left", maxLines: 22 } },
    p5act:  { seqpat: { x: 25, y: 710, w: 210, h: 140, size: 14, align: "left", maxLines: 8 } },
    p5chart:{ chart:  { x: 300, y: 320, w: 320, h: 320 } },

    p6TLDR: { themeExpl: { x: 25, y: 170, w: 550, h: 220, size: 15, align: "left", maxLines: 10 } },
    p6main: { themeExpl: { x: 25, y: 340, w: 550, h: 520, size: 15, align: "left", maxLines: 26 } },
    p6act:  { themeExpl: { x: 25, y: 630, w: 550, h: 140, size: 15, align: "left", maxLines: 8 } },

    p7TLDR: { themesTop: { x: 30, y: 170, w: 590, h: 220, size: 15, align: "left", maxLines: 10 } },
    p7main: { themesTop: { x: 30, y: 340, w: 590, h: 520, size: 15, align: "left", maxLines: 26 } },
    p7act:  { themesTop: { x: 30, y: 630, w: 590, h: 140, size: 15, align: "left", maxLines: 8 } },

    // Page 8: 2x2 grid
    p8grid: {
      collabC: { x: 30,  y: 200, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
      collabT: { x: 320, y: 200, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
      collabR: { x: 30,  y: 650, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
      collabL: { x: 320, y: 650, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
    },

    // Page 9 (action anchor)
    p9: { actAnchor: { x: 25, y: 200, w: 550, h: 220, size: 20, align: "left", maxLines: 8 } },
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

/* ───────── URL-driven layout overrides ─────────
   Supports your exact pattern:
   L_p3TLDR_domDesc_x=...
   L_p3main_domDesc_w=...
   L_p5chart_chart_x=...
*/
function applyLayoutOverridesFromUrl(layout, url) {
  if (!layout || !layout.pages || !url || !url.searchParams) return { layout, applied: [], ignored: [] };

  const allowed = new Set(["x","y","w","h","size","maxLines","align","pad","lineGap","bg"]);
  const pages = layout.pages;

  const applied = [];
  const ignored = [];

  for (const [k, v] of url.searchParams.entries()) {
    if (!k.startsWith("L_")) continue;

    const bits = k.split("_"); // L_p3TLDR_domDesc_y
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

/* ───────── input normaliser (matches your current payload shape) ───────── */
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
  const dateLbl = d.dateLbl || d.date || d["p1:d"] || (d.meta && d.meta.dateLbl) || "";

  // Bands (12)
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

  // Key pieces
  const domState =
    d.domState ||
    ctrl.dominant ||
    summary.domState ||
    ctrl.domState ||
    d["p3:dom"] ||
    "";

  // Page 3: Exec
  const p3_tldr = formatTLDR(text.execSummary_tldr || "");
  const p3_main = S(text.execSummary || "");
  const p3_act  = S(text.execSummary_tipact || text.tipAction || "");

  // Page 4: State deep-dive
  const p4_tldr = formatTLDR(text.state_tldr || text.domState_tldr || "");
  const p4_main = S(text.domState || "");
  const p4_act  = S(text.state_tipact || "");

  // Page 5: Frequency (plus chart)
  const p5_tldr = formatTLDR(text.frequency_tldr || "");
  const p5_main = S(text.frequency || "");
  const p5_act  = S(text.frequency_tipact || ""); // optional
  const chartUrl = S(d.chartUrl || d.chart?.url || d["p5:chart"] || "");

  // Page 6: Sequence
  const p6_tldr = formatTLDR(text.sequence_tldr || "");
  const p6_main = S(text.sequence || "");
  const p6_act  = S(text.sequence_tipact || "");

  // Page 7: Themes (single block)
  const p7_tldr = formatTLDR(text.theme_tldr || text.themesTop_tldr || "");
  const p7_main = S(text.theme || text.themesTop || "");
  const p7_act  = S(text.theme_tipact || text.themesTop_tip || "");

  // Page 8: WorkWith blocks
  const p8C = S(workWith.concealed || "");
  const p8T = S(workWith.triggered || "");
  const p8R = S(workWith.regulated || "");
  const p8L = S(workWith.lead || "");

  // Page 9: Action anchor
  const p9_anchor = S(text.act_anchor || text.action_anchor || "");

  return {
    raw: d,
    identity: { fullName: name, email, dateLabel: dateLbl },

    bands: bandsRaw,
    layout: d.layout || null,

    "p1:n": name,
    "p1:d": dateLbl,
    "p3:dom": domState,

    "p3:tldr": p3_tldr,
    "p3:main": p3_main,
    "p3:act":  p3_act,

    "p4:tldr": p4_tldr,
    "p4:main": p4_main,
    "p4:act":  p4_act,

    "p5:tldr": p5_tldr,
    "p5:main": p5_main,
    "p5:act":  p5_act,
    "p5:chartUrl": chartUrl,

    "p6:tldr": p6_tldr,
    "p6:main": p6_main,
    "p6:act":  p6_act,

    "p7:tldr": p7_tldr,
    "p7:main": p7_main,
    "p7:act":  p7_act,

    "p8:C": p8C,
    "p8:T": p8T,
    "p8:R": p8R,
    "p8:L": p8L,

    "p9:anchor": p9_anchor,
  };
}

/* ───────── master probe summary (like your current debug output) ───────── */
function buildMasterProbe(P, domSecond) {
  const fullName = S(P.identity?.fullName || "");
  const email = S(P.identity?.email || "");
  const dateLabel = S(P.identity?.dateLabel || P["p1:d"] || "");

  const bands = P.bands || {};
  const bandKeys = Object.keys(bands);
  const required12 = [
    "C_low","C_mid","C_high","T_low","T_mid","T_high",
    "R_low","R_mid","R_high","L_low","L_mid","L_high",
  ];
  const present12 = required12.filter((k) => bandKeys.includes(k)).length;

  const textKeys = 0
    + (P["p3:tldr"] ? 1 : 0) + (P["p3:main"] ? 1 : 0) + (P["p3:act"] ? 1 : 0)
    + (P["p4:tldr"] ? 1 : 0) + (P["p4:main"] ? 1 : 0) + (P["p4:act"] ? 1 : 0)
    + (P["p5:tldr"] ? 1 : 0) + (P["p5:main"] ? 1 : 0) + (P["p5:act"] ? 1 : 0)
    + (P["p6:tldr"] ? 1 : 0) + (P["p6:main"] ? 1 : 0) + (P["p6:act"] ? 1 : 0)
    + (P["p7:tldr"] ? 1 : 0) + (P["p7:main"] ? 1 : 0) + (P["p7:act"] ? 1 : 0)
    + (P["p9:anchor"] ? 1 : 0);

  const workWithKeys = 0
    + (P["p8:C"] ? 1 : 0)
    + (P["p8:T"] ? 1 : 0)
    + (P["p8:R"] ? 1 : 0)
    + (P["p8:L"] ? 1 : 0);

  const missing = {
    identity: [],
    ctrl: [],
    text: [],
    workWith: [],
  };

  if (!fullName.trim()) missing.identity.push("identity.fullName");
  if (!email.trim()) missing.identity.push("identity.email");
  if (!dateLabel.trim()) missing.identity.push("identity.dateLabel");

  if (!domSecond?.domKey) missing.ctrl.push("ctrl.summary.dominantKey");
  if (!domSecond?.secondKey) missing.ctrl.push("ctrl.summary.secondKey");
  if (!domSecond?.templateKey) missing.ctrl.push("ctrl.summary.templateKey");

  // minimal checks
  if (!P["p3:tldr"]) missing.text.push("text.execSummary_tldr");
  if (!P["p3:main"]) missing.text.push("text.execSummary");
  if (!P["p3:act"]) missing.text.push("text.execSummary_tipact");

  if (!P["p5:main"]) missing.text.push("text.frequency");
  if (!P["p6:main"]) missing.text.push("text.sequence");
  if (!P["p7:main"]) missing.text.push("text.theme");

  if (!P["p8:C"]) missing.workWith.push("workWith.concealed");
  if (!P["p8:T"]) missing.workWith.push("workWith.triggered");
  if (!P["p8:R"]) missing.workWith.push("workWith.regulated");
  if (!P["p8:L"]) missing.workWith.push("workWith.lead");

  return {
    ok: true,
    where: "fill-template:v8:master_probe:summary",
    domSecond: safeJson(domSecond),
    identity: {
      fullName: { has: !!fullName, len: fullName.length, preview: fullName.slice(0, 40) },
      email: { has: !!email, len: email.length, preview: email.slice(0, 40) },
      dateLabel: { has: !!dateLabel, len: dateLabel.length, preview: dateLabel.slice(0, 40) },
    },
    counts: {
      questions: 0,               // (kept to match your existing probe schema)
      bandsKeys: bandKeys.length,
      bandsPresent12: present12,
      textKeys,
      workWithKeys,
      freqTldrExtraKeys: [],
    },
    lengths: {
      p3_tldr: (P["p3:tldr"] || "").length,
      p3_main: (P["p3:main"] || "").length,
      p3_act:  (P["p3:act"]  || "").length,

      p4_tldr: (P["p4:tldr"] || "").length,
      p4_main: (P["p4:main"] || "").length,
      p4_act:  (P["p4:act"]  || "").length,

      p5_tldr: (P["p5:tldr"] || "").length,
      p5_main: (P["p5:main"] || "").length,
      p5_act:  (P["p5:act"]  || "").length,

      p6_tldr: (P["p6:tldr"] || "").length,
      p6_main: (P["p6:main"] || "").length,
      p6_act:  (P["p6:act"]  || "").length,

      p7_tldr: (P["p7:tldr"] || "").length,
      p7_main: (P["p7:main"] || "").length,
      p7_act:  (P["p7:act"]  || "").length,

      p9_anchor: (P["p9:anchor"] || "").length,
      bandsKeys: bandKeys.length,
      bandsPresent12: present12,
      freqTldrExtraKeys: 0,
    },
    missing,
    previews: {
      execSummary_tldr: (P["p3:tldr"] || "").slice(0, 140),
      execSummary: (P["p3:main"] || "").slice(0, 140),
      frequency_tldr: (P["p5:tldr"] || "").slice(0, 140),
      frequency: (P["p5:main"] || "").slice(0, 140),
      sequence_tldr: (P["p6:tldr"] || "").slice(0, 140),
      theme_tldr: (P["p7:tldr"] || "").slice(0, 140),
      act_anchor: (P["p9:anchor"] || "").slice(0, 140),
      workWith_triggered: (P["p8:T"] || "").slice(0, 140),
    },
  };
}

/* ───────── main handler ───────── */
export default async function handler(req, res) {
  try {
    const url = new URL(req.url, "http://localhost");
    const debug = url.searchParams.get("debug") === "1";

    const payload = await readPayload(req);
    const P = normaliseInput(payload);
    const domSecond = computeDomAndSecondKeys(P);

    // Layout: default + payload overrides + URL overrides (full)
    let layout = deepMerge(DEFAULT_LAYOUT, P.layout || {});
    const { layout: layout2, applied, ignored } = applyLayoutOverridesFromUrl(layout, url);
    layout = layout2;
    const L = layout.pages || DEFAULT_LAYOUT.pages;

    if (debug) {
      // include applied/ignored overrides so you can see why coords "did not move"
      const probe = buildMasterProbe(P, domSecond);
      probe.layoutOverrides = { applied, ignored };
      return res.status(200).json(probe);
    }

    // Template resolve
    const validCombos = new Set(["CT","CL","CR","TC","TR","TL","RC","RT","RL","LC","LR","LT"]);
    const safeCombo = validCombos.has(domSecond.templateKey) ? domSecond.templateKey : "CT";
    const tpl = `CTRL_PoC_Assessment_Profile_template_${safeCombo}.pdf`;

    const pdfBytes = await loadTemplateBytesLocal(tpl);
    const pdfDoc = await PDFDocument.load(pdfBytes);

    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontB = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    const pages = pdfDoc.getPages();

    // p1 (cover): name + date
    if (pages[0] && L.p1) {
      if (L.p1.name && P["p1:n"]) drawTextBox(pages[0], fontB, P["p1:n"], L.p1.name, { maxLines: 1 });
      if (L.p1.date && P["p1:d"]) drawTextBox(pages[0], font,  P["p1:d"], L.p1.date, { maxLines: 1 });
    }

// Header (pages 2–10): write just FullName
const headerName = norm(P["p1:n"]);
if (headerName) {
  for (let i = 1; i < pages.length; i++) {
    const pk = `p${i + 1}`;
    const box = L?.[pk]?.hdrName;
    if (box) drawTextBox(pages[i], font, headerName, box, { maxLines: 1 });
  }
}


    // Index mapping (template is 10 pages: 1..10 => 0..9)
    const p3 = pages[2] || null;
    const p4 = pages[3] || null;
    const p5 = pages[4] || null;
    const p6 = pages[5] || null;
    const p7 = pages[6] || null;
    const p8 = pages[7] || null;
    const p9 = pages[8] || null;

    // Page 3 — TLDR / Main / Action
    if (p3) {
      if (L.p3TLDR?.domDesc) drawLabelAndBody(p3, fontB, font, "TLDR",  P["p3:tldr"], L.p3TLDR.domDesc);
      if (L.p3main?.domDesc) drawLabelAndBody(p3, fontB, font, "",      P["p3:main"], L.p3main.domDesc);
      if (L.p3act?.domDesc)  drawLabelAndBody(p3, fontB, font, "Action", P["p3:act"],  L.p3act.domDesc);
    }

    // Page 4 — TLDR / Main / Action
    if (p4) {
      if (L.p4TLDR?.spider) drawLabelAndBody(p4, fontB, font, "TLDR",  P["p4:tldr"], L.p4TLDR.spider);
      if (L.p4main?.spider) drawLabelAndBody(p4, fontB, font, "",      P["p4:main"], L.p4main.spider);
      if (L.p4act?.spider)  drawLabelAndBody(p4, fontB, font, "Action", P["p4:act"],  L.p4act.spider);
    }

    // Page 5 — TLDR / Main / Action + radar chart
    if (p5) {
      if (L.p5TLDR?.seqpat) drawLabelAndBody(p5, fontB, font, "TLDR",  P["p5:tldr"], L.p5TLDR.seqpat);
      if (L.p5main?.seqpat) drawLabelAndBody(p5, fontB, font, "",      P["p5:main"], L.p5main.seqpat);
      if (L.p5act?.seqpat)  drawLabelAndBody(p5, fontB, font, "Action", P["p5:act"],  L.p5act.seqpat);

      if (L.p5chart?.chart) {
        try {
          await embedRadarFromBandsOrUrl(pdfDoc, p5, L.p5chart.chart, P.bands || {}, P["p5:chartUrl"]);
        } catch (e) {
          console.warn("[fill-template:v8] Radar chart skipped:", e?.message || String(e));
        }
      }
    }

    // Page 6 — TLDR / Main / Action
    if (p6) {
      if (L.p6TLDR?.themeExpl) drawLabelAndBody(p6, fontB, font, "TLDR",  P["p6:tldr"], L.p6TLDR.themeExpl);
      if (L.p6main?.themeExpl) drawLabelAndBody(p6, fontB, font, "",      P["p6:main"], L.p6main.themeExpl);
      if (L.p6act?.themeExpl)  drawLabelAndBody(p6, fontB, font, "Action", P["p6:act"],  L.p6act.themeExpl);
    }

    // Page 7 — TLDR / Main / Action (single theme block)
    if (p7) {
      if (L.p7TLDR?.themesTop) drawLabelAndBody(p7, fontB, font, "TLDR",  P["p7:tldr"], L.p7TLDR.themesTop);
      if (L.p7main?.themesTop) drawLabelAndBody(p7, fontB, font, "",      P["p7:main"], L.p7main.themesTop);
      if (L.p7act?.themesTop)  drawLabelAndBody(p7, fontB, font, "Action", P["p7:act"],  L.p7act.themesTop);
    }

    // Page 8 — WorkWith 2x2 grid
    if (p8 && L.p8grid) {
      if (L.p8grid.collabC && P["p8:C"]) drawTextBox(p8, font, P["p8:C"], L.p8grid.collabC, { maxLines: L.p8grid.collabC.maxLines });
      if (L.p8grid.collabT && P["p8:T"]) drawTextBox(p8, font, P["p8:T"], L.p8grid.collabT, { maxLines: L.p8grid.collabT.maxLines });
      if (L.p8grid.collabR && P["p8:R"]) drawTextBox(p8, font, P["p8:R"], L.p8grid.collabR, { maxLines: L.p8grid.collabR.maxLines });
      if (L.p8grid.collabL && P["p8:L"]) drawTextBox(p8, font, P["p8:L"], L.p8grid.collabL, { maxLines: L.p8grid.collabL.maxLines });
    }

    // Page 9 — Action anchor
    if (p9 && L.p9?.actAnchor && P["p9:anchor"]) {
      drawTextBox(p9, font, P["p9:anchor"], L.p9.actAnchor, { maxLines: L.p9.actAnchor.maxLines });
    }

    const outBytes = await pdfDoc.save();

    // Filename fix
    const outName = makeOutputFilename(P["p1:n"], P["p1:d"]);
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Cache-Control", "no-store");
    res.setHeader("Content-Disposition", `inline; filename="${outName}"`);
    res.status(200).send(Buffer.from(outBytes));
  } catch (err) {
    console.error("[fill-template:v8] CRASH", err);
    res.status(500).json({
      ok: false,
      error: err?.message || String(err),
      stack: err?.stack || null,
    });
  }
}
