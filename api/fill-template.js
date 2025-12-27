/**
 * CTRL PoC Export Service · fill-template (V12.1)
 *
 * V12.1 fixes:
 * 1) Reinstates V9 coordinates for:
 *    - Page 1 FullName + Date
 *    - Header name on pages 2–8
 * 2) Reinstates V9 chart rendering (polarArea, colours, style),
 *    but renders it on Page 4 (not Page 5).
 * 3) Fixes "Gen text missing" by accepting BOTH:
 *    - new keys: exec_summary_para1 / exec_summary_para2, etc
 *    - old single-block keys: exec_summary, ctrl_overview, ctrl_deepdive, themes
 *      and splitting into para1/para2 when needed.
 * 4) Keeps template selection + fallback behaviour unchanged.
 * 5) Keeps base64 data decoding unchanged.
 * 6) Adds robust &debug=1 probe (every major key + lengths + missing list).
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
const okObj = (o) => o && typeof o === "object" && !Array.isArray(o);
const okArr = (a) => Array.isArray(a);

function safeJson(obj) {
  try { return JSON.parse(JSON.stringify(obj)); }
  catch { return { _error: "Could not serialise debug object" }; }
}

/* ───────── filename helpers (unchanged) ───────── */
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
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
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
    if (!words.length) { lines.push(""); continue; }
    for (const word of words) {
      const test = line ? `${line} ${word}` : word;
      const width = font.widthOfTextAtSize(test, size);
      if (width <= w) line = test;
      else { if (line) lines.push(line); line = word; }
    }
    if (line) lines.push(line);
  }
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

  if (opts.bg === true || box.bg === true) {
    page.drawRectangle({
      x, y, width: w, height: h,
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

/* ───────── paragraph splitting (for missing para1/para2) ───────── */
function splitToTwoParas(s) {
  const raw = S(s).replace(/\r/g, "").trim();
  if (!raw) return ["", ""];
  // Prefer blank-line split
  const parts = raw.split(/\n\s*\n/).map((p) => p.trim()).filter(Boolean);
  if (parts.length >= 2) return [parts[0], parts.slice(1).join("\n\n")];
  // Fallback: sentence split roughly halfway
  const sentences = raw.split(/(?<=[.!?])\s+/).filter(Boolean);
  if (sentences.length >= 2) {
    const mid = Math.ceil(sentences.length / 2);
    return [sentences.slice(0, mid).join(" ").trim(), sentences.slice(mid).join(" ").trim()];
  }
  return [raw, ""];
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
    try { return await fs.readFile(pth); }
    catch (err) { lastErr = err; }
  }

  throw new Error(
    `Template not found: ${fname}. Tried: ${candidates.join(" | ")}. Last: ${lastErr?.message || "no detail"}`
  );
}

/* ───────── payload parsing (unchanged) ───────── */
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

/* ───────── dom/second detection (kept) ───────── */
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
    resolveStateKey(raw.dominantKey) ||
    resolveStateKey(summary.dominant) ||
    resolveStateKey(summary.domState) ||
    resolveStateKey(raw.ctrl?.dominant) ||
    resolveStateKey(raw.domState) ||
    "R";

  const secondKey =
    resolveStateKey(P.secondKey) ||
    resolveStateKey(raw.secondKey) ||
    resolveStateKey(summary.secondState) ||
    resolveStateKey(raw.secondState) ||
    (domKey === "R" ? "T" : "R");

  return { domKey, secondKey, templateKey: `${domKey}${secondKey}` };
}

/* ───────── radar/polar chart embed (FROM V9) ───────── */
function makeSpiderChartUrl12(bandsRaw) {
  const keys = [
    "C_low","C_mid","C_high",
    "T_low","T_mid","T_high",
    "R_low","R_mid","R_high",
    "L_low","L_mid","L_high",
  ];

  const displayLabels = [
    "", "Concealed", "",
    "", "Triggered", "",
    "", "Regulated", "",
    "", "Lead", ""
  ];

  const vals = keys.map((k) => Number(bandsRaw?.[k] || 0));
  const maxVal = Math.max(...vals, 1);
  const data = vals.map((v) => (maxVal > 0 ? v / maxVal : 0));

  const CTRL_COLOURS = {
    C: {
      low:  "rgba(230, 228, 225, 0.55)",
      mid:  "rgba(184, 180, 174, 0.55)",
      high: "rgba(110, 106, 100, 0.55)",
    },
    T: {
      low:  "rgba(244, 225, 198, 0.55)",
      mid:  "rgba(211, 155,  74, 0.55)",
      high: "rgba(154,  94,  26, 0.55)",
    },
    R: {
      low:  "rgba(226, 236, 230, 0.55)",
      mid:  "rgba(143, 183, 161, 0.55)",
      high: "rgba( 79, 127, 105, 0.55)",
    },
    L: {
      low:  "rgba(230, 220, 227, 0.55)",
      mid:  "rgba(164, 135, 159, 0.55)",
      high: "rgba( 94,  63,  90, 0.55)",
    },
  };

  const colours = keys.map((k) => {
    const state = k[0];
    const tier = k.split("_")[1];
    return CTRL_COLOURS[state]?.[tier] || "rgba(0,0,0,0.10)";
  });

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
    const hasAny = bandsRaw && typeof bandsRaw === "object" &&
      Object.values(bandsRaw).some((v) => Number(v) > 0);
    if (!hasAny) return;
    url = makeSpiderChartUrl12(bandsRaw);
  }

  const img = await embedRemoteImage(pdfDoc, url);
  if (!img) return;

  const H = page.getHeight();
  page.drawImage(img, {
    x: box.x,
    y: H - box.y - box.h,
    width: box.w,
    height: box.h
  });
}

/* ───────── DEFAULT LAYOUT (V9 coords + new 8-page map) ───────── */
const DEFAULT_LAYOUT = {
  pages: {
    // Page 1 — V9 coords (keep)
    p1: {
      name: { x: 60,  y: 458, w: 500, h: 60, size: 30, align: "center", maxLines: 1 },
      date: { x: 230, y: 613, w: 500, h: 40, size: 25, align: "left",   maxLines: 1 },
    },

    // Header name on pages 2–8 — V9 coords (keep)
    p2: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p3: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p4: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p5: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p6: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p7: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p8: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },

    // Page 3 — Executive summary (your new para1/para2 boxes)
    p3Text: {
      exec1: { x: 25, y: 310, w: 550, h: 250, size: 15, align: "left", maxLines: 13 },
      exec2: { x: 25, y: 570, w: 550, h: 420, size: 15, align: "left", maxLines: 22 },
    },

    // Page 4 — CTRL overview + CHART (chart moved here)
    p4Text: {
      ov1: { x: 25, y: 160, w: 550, h: 240, size: 14, align: "left", maxLines: 13 },
      ov2: { x: 25, y: 410, w: 550, h: 420, size: 14, align: "left", maxLines: 23 },
      // chart box: use the SAME V9 chart coords, but on Page 4 now
      chart: { x: 250, y: 160, w: 320, h: 320 },
    },

    // Page 5 — Deepdive + Themes
    p5Text: {
      dd1: { x: 25, y: 170, w: 550, h: 240, size: 15, align: "left", maxLines: 13 },
      dd2: { x: 25, y: 420, w: 550, h: 310, size: 15, align: "left", maxLines: 17 },
      th1: { x: 25, y: 740, w: 550, h: 160, size: 15, align: "left", maxLines: 9 },
      th2: { x: 25, y: 910, w: 550, h: 160, size: 15, align: "left", maxLines: 9 },
    },

    // Page 6 — WorkWith (4 boxes)
    p6WorkWith: {
      collabC: { x: 30,  y: 300, w: 270, h: 420, size: 14, align: "left", maxLines: 14 },
      collabT: { x: 320, y: 300, w: 260, h: 420, size: 14, align: "left", maxLines: 14 },
      collabR: { x: 30,  y: 575, w: 260, h: 420, size: 14, align: "left", maxLines: 14 },
      collabL: { x: 320, y: 575, w: 260, h: 420, size: 14, align: "left", maxLines: 14 },
    },

    // Page 7 — Actions (Act1..Act6)
    p7Actions: {
      act1: { x: 30, y: 240, w: 550, h: 95, size: 16, align: "left", maxLines: 5 },
      act2: { x: 30, y: 345, w: 550, h: 95, size: 16, align: "left", maxLines: 5 },
      act3: { x: 30, y: 450, w: 550, h: 95, size: 16, align: "left", maxLines: 5 },
      act4: { x: 30, y: 555, w: 550, h: 95, size: 16, align: "left", maxLines: 5 },
      act5: { x: 30, y: 660, w: 550, h: 95, size: 16, align: "left", maxLines: 5 },
      act6: { x: 30, y: 765, w: 550, h: 95, size: 16, align: "left", maxLines: 5 },
    },
  },
};

/* ───────── input normaliser (NEW KEY SUPPORT + backward compatible) ───────── */
function normaliseInput(d = {}) {
  const identity = okObj(d.identity) ? d.identity : {};
  const text = okObj(d.text) ? d.text : {};
  const workWith = okObj(d.workWith) ? d.workWith : {};
  const ctrl = okObj(d.ctrl) ? d.ctrl : {};
  const summary = okObj(ctrl.summary) ? ctrl.summary : {};

  const fullName =
    S(identity.fullName || d.fullName || d.FullName || summary?.identity?.fullName || "").trim();
  const dateLabel =
    S(identity.dateLabel || d.dateLbl || d.date || d.Date || summary?.dateLbl || "").trim();

  // Bands (chart): prefer ctrl.summary.ctrl12 then ctrl12/bands
  const bandsRaw =
    (okObj(summary.ctrl12) && Object.keys(summary.ctrl12).length ? summary.ctrl12 : null) ||
    (okObj(d.bands) && Object.keys(d.bands).length ? d.bands : null) ||
    (okObj(ctrl.bands) && Object.keys(ctrl.bands).length ? ctrl.bands : null) ||
    {};

  // ---- Text mapping: NEW keys first, else fallback + split ----
  const exec1 = S(text.exec_summary_para1 || "");
  const exec2 = S(text.exec_summary_para2 || "");
  const execBlock = S(text.exec_summary || "");
  const [execA, execB] = splitToTwoParas(execBlock);

  const ov1 = S(text.ctrl_overview_para1 || "");
  const ov2 = S(text.ctrl_overview_para2 || "");
  const ovBlock = S(text.ctrl_overview || "");
  const [ovA, ovB] = splitToTwoParas(ovBlock);

  const dd1 = S(text.ctrl_deepdive_para1 || "");
  const dd2 = S(text.ctrl_deepdive_para2 || "");
  const ddBlock = S(text.ctrl_deepdive || "");
  const [ddA, ddB] = splitToTwoParas(ddBlock);

  const th1 = S(text.themes_para1 || "");
  const th2 = S(text.themes_para2 || "");
  const thBlock = S(text.themes || "");
  const [thA, thB] = splitToTwoParas(thBlock);

  // Actions: prefer explicit Act1..Act6, else text.act_1..act_6, else array
  const act1 = S(text.Act1 || text.act_1 || "");
  const act2 = S(text.Act2 || text.act_2 || "");
  const act3 = S(text.Act3 || text.act_3 || "");
  const act4 = S(text.Act4 || text.act_4 || "");
  const act5 = S(text.Act5 || text.act_5 || "");
  const act6 = S(text.Act6 || text.act_6 || "");

  const actsArr = okArr(text.actions) ? text.actions.map((x) => S(x)) : [];
  const actFromArr = (i) => S(actsArr[i] || "");

  const chartUrl =
    S(d.spiderChartUrl || d.spider_chart_url || d.chartUrl || text.chartUrl || "").trim() ||
    S(d.chart?.spiderUrl || d.chart?.url || "").trim() ||
    S(summary?.chart?.spiderUrl || "").trim();

  return {
    raw: d,
    identity: { fullName, dateLabel },
    bands: bandsRaw,

    // resolved text (always filled if any source exists)
    exec_summary_para1: exec1 || execA,
    exec_summary_para2: exec2 || execB,

    ctrl_overview_para1: ov1 || ovA,
    ctrl_overview_para2: ov2 || ovB,

    ctrl_deepdive_para1: dd1 || ddA,
    ctrl_deepdive_para2: dd2 || ddB,

    themes_para1: th1 || thA,
    themes_para2: th2 || thB,

    Act1: act1 || actFromArr(0),
    Act2: act2 || actFromArr(1),
    Act3: act3 || actFromArr(2),
    Act4: act4 || actFromArr(3),
    Act5: act5 || actFromArr(4),
    Act6: act6 || actFromArr(5),

    workWith: {
      concealed: S(workWith.concealed || ""),
      triggered: S(workWith.triggered || ""),
      regulated: S(workWith.regulated || ""),
      lead: S(workWith.lead || ""),
    },

    chartUrl,
  };
}

/* ───────── debug probe ───────── */
function buildProbe(P, domSecond, tpl) {
  const bands = P.bands || {};
  const required12 = [
    "C_low","C_mid","C_high","T_low","T_mid","T_high",
    "R_low","R_mid","R_high","L_low","L_mid","L_high",
  ];
  const keys = Object.keys(bands);
  const missing12 = required12.filter((k) => !keys.includes(k));

  const text = {
    exec1: S(P.exec_summary_para1),
    exec2: S(P.exec_summary_para2),
    ov1: S(P.ctrl_overview_para1),
    ov2: S(P.ctrl_overview_para2),
    dd1: S(P.ctrl_deepdive_para1),
    dd2: S(P.ctrl_deepdive_para2),
    th1: S(P.themes_para1),
    th2: S(P.themes_para2),
    act1: S(P.Act1),
    act2: S(P.Act2),
    act3: S(P.Act3),
    act4: S(P.Act4),
    act5: S(P.Act5),
    act6: S(P.Act6),
  };

  const warnings = [];
  if (!text.exec1.trim() && !text.exec2.trim()) warnings.push("Missing exec_summary_para1/para2 (and no fallback content)");
  if (!text.ov1.trim() && !text.ov2.trim()) warnings.push("Missing ctrl_overview_para1/para2 (and no fallback content)");
  if (!text.dd1.trim() && !text.dd2.trim()) warnings.push("Missing ctrl_deepdive_para1/para2 (and no fallback content)");
  if (!text.th1.trim() && !text.th2.trim()) warnings.push("Missing themes_para1/para2 (and no fallback content)");

  return {
    ok: true,
    where: "fill-template:V12.1:debug",
    template: tpl,
    domSecond: safeJson(domSecond),
    identity: {
      fullName: P.identity.fullName,
      dateLabel: P.identity.dateLabel,
      nameLen: P.identity.fullName.length,
      dateLen: P.identity.dateLabel.length,
    },
    bands: {
      keysCount: keys.length,
      present12Count: required12.length - missing12.length,
      missing12,
      anyPositive: Object.values(bands).some((v) => Number(v) > 0),
      sample: Object.fromEntries(required12.slice(0, 6).map((k) => [k, bands[k]])),
    },
    textLengths: Object.fromEntries(Object.entries(text).map(([k, v]) => [k, v.length])),
    chart: {
      chartUrlProvided: !!P.chartUrl,
      chartUrlPreview: P.chartUrl ? P.chartUrl.slice(0, 80) : "",
    },
    workWithLengths: {
      concealed: (P.workWith?.concealed || "").length,
      triggered: (P.workWith?.triggered || "").length,
      regulated: (P.workWith?.regulated || "").length,
      lead: (P.workWith?.lead || "").length,
    },
    warnings,
  };
}

/* ───────── main handler ───────── */
export default async function handler(req, res) {
  try {
    const url = new URL(req.url, "http://localhost");
    const debug = url.searchParams.get("debug") === "1";

    const payload = await readPayload(req);
    const P = normaliseInput(payload);
    const domSecond = computeDomAndSecondKeys({ raw: payload, domKey: payload?.dominantKey, secondKey: payload?.secondKey });

    // Template selection & fallback: unchanged
    const validCombos = new Set(["CT","CL","CR","TC","TR","TL","RC","RT","RL","LC","LR","LT"]);
    const safeCombo = validCombos.has(domSecond.templateKey) ? domSecond.templateKey : "CT";
    const tpl = {
      combo: domSecond.templateKey,
      safeCombo,
      tpl: `CTRL_PoC_Assessment_Profile_template_${safeCombo}.pdf`,
    };

    if (debug) return res.status(200).json(buildProbe(P, domSecond, tpl));

    const pdfBytes = await loadTemplateBytesLocal(tpl.tpl);
    const pdfDoc = await PDFDocument.load(pdfBytes);

    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontB = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    const pages = pdfDoc.getPages();
    const L = DEFAULT_LAYOUT.pages;

    // Page 1
    if (pages[0]) {
      drawTextBox(pages[0], fontB, P.identity.fullName, L.p1.name, { maxLines: 1 });
      drawTextBox(pages[0], font,  P.identity.dateLabel, L.p1.date, { maxLines: 1 });
    }

    // Header name pages 2–8
    const headerName = norm(P.identity.fullName);
    if (headerName) {
      for (let i = 1; i < Math.min(pages.length, 8); i++) {
        const pk = `p${i + 1}`;
        const box = L?.[pk]?.hdrName;
        if (box) drawTextBox(pages[i], font, headerName, box, { maxLines: 1 });
      }
    }

    // Page mapping (8 pages total)
    const p3 = pages[2] || null; // page 3
    const p4 = pages[3] || null; // page 4 (chart now here)
    const p5 = pages[4] || null; // page 5
    const p6 = pages[5] || null; // page 6 (workwith)
    const p7 = pages[6] || null; // page 7 (actions)
    // p8 = pages[7] header only

    // Page 3 text
    if (p3) {
      drawTextBox(p3, font, P.exec_summary_para1, L.p3Text.exec1);
      drawTextBox(p3, font, P.exec_summary_para2, L.p3Text.exec2);
    }

    // Page 4 text + chart
    if (p4) {
      drawTextBox(p4, font, P.ctrl_overview_para1, L.p4Text.ov1);
      drawTextBox(p4, font, P.ctrl_overview_para2, L.p4Text.ov2);

      try {
        await embedRadarFromBandsOrUrl(pdfDoc, p4, L.p4Text.chart, P.bands || {}, P.chartUrl);
      } catch (e) {
        console.warn("[fill-template:V12.1] Chart skipped:", e?.message || String(e));
      }
    }

    // Page 5 deepdive + themes
    if (p5) {
      drawTextBox(p5, font, P.ctrl_deepdive_para1, L.p5Text.dd1);
      drawTextBox(p5, font, P.ctrl_deepdive_para2, L.p5Text.dd2);
      drawTextBox(p5, font, P.themes_para1, L.p5Text.th1);
      drawTextBox(p5, font, P.themes_para2, L.p5Text.th2);
    }

    // Page 6 workwith
    if (p6) {
      drawTextBox(p6, font, P.workWith?.concealed, L.p6WorkWith.collabC);
      drawTextBox(p6, font, P.workWith?.triggered, L.p6WorkWith.collabT);
      drawTextBox(p6, font, P.workWith?.regulated, L.p6WorkWith.collabR);
      drawTextBox(p6, font, P.workWith?.lead,      L.p6WorkWith.collabL);
    }

    // Page 7 actions
    if (p7) {
      drawTextBox(p7, font, P.Act1, L.p7Actions.act1);
      drawTextBox(p7, font, P.Act2, L.p7Actions.act2);
      drawTextBox(p7, font, P.Act3, L.p7Actions.act3);
      drawTextBox(p7, font, P.Act4, L.p7Actions.act4);
      drawTextBox(p7, font, P.Act5, L.p7Actions.act5);
      drawTextBox(p7, font, P.Act6, L.p7Actions.act6);
    }

    const outBytes = await pdfDoc.save();
    const outName = makeOutputFilename(P.identity.fullName, P.identity.dateLabel);

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Cache-Control", "no-store");
    res.setHeader("Content-Disposition", `inline; filename="${outName}"`);
    res.status(200).send(Buffer.from(outBytes));
  } catch (err) {
    console.error("[fill-template:V12.1] CRASH", err);
    res.status(500).json({
      ok: false,
      error: err?.message || String(err),
      stack: err?.stack || null,
    });
  }
}
