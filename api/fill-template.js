// === CTRL PoC fill-template — V6.1 ===
// Changes vs V5 (targeted, per Toby request):
// V6.1: tighten TLDR bullet spacing (no other changes).
// 1) Remove underline separators ("-----") and replace section headings with bold labels only (no line underlines)
// 2) TLDR bullet points: normalise so each bullet is on its own line
// 3) Remove middle titles like "State Deep Dive" / "Executive Summary" (static page title already exists in PDF)
// 4) Add per-section layout overrides per page:
//    - p3TLDR / p3main / p3act
//    - p4TLDR / p4main / p4act
//    - p5TLDR / p5main (act optional)
//    - p6TLDR / p6main / p6act
//    - p7TLDR / p7main / p7act
//    URL example: &L_p3TLDR_domDesc_x=45&L_p3main_domDesc_y=420&L_p3act_domDesc_w=520
//
// IMPORTANT: Everything else remains aligned with V5 logic/payload mapping.

export const config = { runtime: "nodejs" };

import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";

/* ───────── basics ───────── */

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const PUBLIC_DIR = path.join(__dirname, "..", "public");

const TEMPLATE_PREFIX = "CTRL_PoC_Assessment_Profile_template_";
const TEMPLATE_SUFFIX = ".pdf";
const VALID_TEMPLATE_KEYS = new Set([
  "CT","CR","CL",
  "TC","TR","TL",
  "RC","RT","RL",
  "LC","LT","LR",
]);

/* ───────── layout ───────── */

const DEFAULT_LAYOUT = {
  v: 6,
  pages: {
    p1: {
      name: { x: 7, y: 473, w: 500, h: 60, size: 30, align: "center", maxLines: 1 },
      date: { x: 210, y: 600, w: 500, h: 40, size: 25, align: "left", maxLines: 1 },
    },

    // p2–p10 header name
    p2: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },

    // p3 physical page: header lives here
    p3: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    // p3 sections (same physical page 3)
    p3TLDR: { domDesc: { x: 25, y: 200, w: 550, h: 220, size: 18, align: "left", maxLines: 10 } },
    p3main: { domDesc: { x: 25, y: 430, w: 550, h: 520, size: 18, align: "left", maxLines: 26 } },
    p3act:  { domDesc: { x: 25, y: 960, w: 550, h: 140, size: 18, align: "left", maxLines: 8 } },

    // p4
    p4: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p4TLDR: { spider: { x: 25, y: 200, w: 550, h: 220, size: 18, align: "left", maxLines: 10 } },
    p4main: { spider: { x: 25, y: 430, w: 550, h: 520, size: 18, align: "left", maxLines: 26 } },
    p4act:  { spider: { x: 25, y: 960, w: 550, h: 140, size: 18, align: "left", maxLines: 8 } },

    // p5 (Frequency text + chart)
    p5: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      chart: { x: 48, y: 462, w: 500, h: 300 },
    },
    p5TLDR: { seqpat: { x: 25, y: 200, w: 550, h: 220, size: 18, align: "left", maxLines: 10 } },
    p5main: { seqpat: { x: 25, y: 430, w: 550, h: 520, size: 18, align: "left", maxLines: 28 } },
    // p5act is optional; uncomment if you later add it
    // p5act: { seqpat: { x: 25, y: 960, w: 550, h: 140, size: 18, align: "left", maxLines: 8 } },

    // p6 (Sequence)
    p6: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p6TLDR: { themeExpl: { x: 25, y: 200, w: 550, h: 220, size: 18, align: "left", maxLines: 10 } },
    p6main: { themeExpl: { x: 25, y: 430, w: 550, h: 520, size: 18, align: "left", maxLines: 26 } },
    p6act:  { themeExpl: { x: 25, y: 960, w: 550, h: 140, size: 18, align: "left", maxLines: 8 } },

    // p7 (Themes)
    p7: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      themesLow: { x: 320, y: 200, w: 300, h: 900, size: 17, align: "left", maxLines: 28 },
    },
    p7TLDR: { themesTop: { x: 30, y: 200, w: 590, h: 220, size: 17, align: "left", maxLines: 10 } },
    p7main: { themesTop: { x: 30, y: 430, w: 590, h: 520, size: 17, align: "left", maxLines: 26 } },
    p7act:  { themesTop: { x: 30, y: 960, w: 590, h: 140, size: 17, align: "left", maxLines: 8 } },

    // p8 (WorkWith / Collaboration)
    p8: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      collabC: { x: 30, y: 200, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
      collabT: { x: 320, y: 200, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
      collabR: { x: 30, y: 650, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
      collabL: { x: 320, y: 650, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
    },

    // p9 (Action Anchor)
    p9: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      actAnchor: { x: 25, y: 200, w: 550, h: 220, size: 20, align: "left", maxLines: 8 },
    },

    p10: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
  },
};

function deepClone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function mergeLayout(base, override) {
  const out = deepClone(base);
  if (!override || typeof override !== "object") return out;
  if (override.v != null) out.v = override.v;
  if (override.pages && typeof override.pages === "object") {
    out.pages = out.pages || {};
    for (const [pk, pv] of Object.entries(override.pages)) {
      out.pages[pk] = out.pages[pk] || {};
      if (pv && typeof pv === "object") {
        for (const [bk, bv] of Object.entries(pv)) {
          out.pages[pk][bk] = { ...(out.pages[pk][bk] || {}), ...(bv || {}) };
        }
      }
    }
  }
  return out;
}

/* ───────── URL layout overrides ───────── */

function parseLayoutParamKey(key) {
  if (!key || typeof key !== "string") return null;

  // Strip optional leading "L_"
  const k = key.startsWith("L_") || key.startsWith("l_") ? key.slice(2) : key;

  // Style: p3TLDR_domDesc_x   (NEW, V6)
  // or:    p3_domDesc_x       (legacy)
  const bits = k.split("_");
  if (bits.length >= 3 && /^p\d+/i.test(bits[0])) {
    const pageKey = bits[0];                 // allow p3TLDR / p3main / p3act etc
    const boxKey = bits[1];
    const prop = bits.slice(2).join("_");
    return { pageKey, boxKey, prop };
  }

  // Legacy compact style: p1_namex
  const m = k.match(/^(p\d+)_([A-Za-z0-9]+)(x|y|w|h|size|maxLines|align)$/i);
  if (m) return { pageKey: m[1], boxKey: m[2], prop: m[3] };

  return null;
}

function applyLayoutOverridesFromUrl(layout, url) {
  const applied = [];
  const ignored = [];
  if (!layout || !layout.pages || !url) return { applied, ignored };

  const sp = url.searchParams;
  for (const [rawKey, rawVal] of sp.entries()) {
    const parsed = parseLayoutParamKey(rawKey);
    if (!parsed) continue;

    const { pageKey, boxKey, prop } = parsed;
    const page = layout.pages[pageKey];
    const box = page ? page[boxKey] : null;

    if (!page || !box || typeof box !== "object") {
      ignored.push({ key: rawKey, reason: "unknown_box_or_page" });
      continue;
    }

    const allow = new Set(["x","y","w","h","size","maxLines","align","titleSize","lineGap","pad"]);
    const canSet = Object.prototype.hasOwnProperty.call(box, prop) || allow.has(prop);
    if (!canSet) {
      ignored.push({ key: rawKey, reason: "prop_not_allowed" });
      continue;
    }

    let v;
    if (prop === "align") {
      v = String(rawVal || "").toLowerCase();
      if (!["left","center","right"].includes(v)) v = "left";
    } else if (rawVal === "" || rawVal == null) {
      ignored.push({ key: rawKey, reason: "empty_value" });
      continue;
    } else {
      const n = Number(rawVal);
      if (Number.isFinite(n)) v = n;
      else {
        ignored.push({ key: rawKey, reason: "not_a_number" });
        continue;
      }
    }

    box[prop] = v;
    applied.push({ key: rawKey, pageKey, boxKey, prop, value: v });
  }

  return { applied, ignored };
}

/* ───────── text safety (WinAnsi friendly) ───────── */

function safeText(s) {
  if (s == null) return "";
  let out = String(s);

  out = out.replace(/\u00A0|\u202F/g, " ");
  out = out.replace(/\uFEFF/g, "");
  out = out.replace(/[\u2500-\u257F]/g, "-");
  out = out.replace(/\u2013|\u2014|\u2212/g, "-");
  out = out.replace(/\u2018|\u2019/g, "'");
  out = out.replace(/\u201C|\u201D/g, '"');
  out = out.replace(/\u2026/g, "...");
  out = out.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, " ");

  return out;
}

function norm(s) {
  const t = safeText(s);
  return t.trim();
}

// TLDR bullet normaliser:
// - Turns inline bullets into separate lines
// - Keeps bullets as "• " lines (still wraps if too long, but at least 1 bullet per paragraph/line start)
function normaliseTldrBullets(s) {
  let t = norm(s);
  if (!t) return "";
  // convert " . • " or " • " into new lines
  t = t.replace(/\s*[•·]\s*/g, "\n• ");
  // also convert " - " used as bullets into new lines when it looks like a list
  t = t.replace(/\n-\s+/g, "\n• ");
  // remove duplicate bullet markers
  t = t.replace(/\n•\s*•\s*/g, "\n• ");
  // trim each line
  t = t.split("\n").map(x => x.trim()).filter(x => x.length > 0).join("\n");
  // ensure every line starts with bullet if it is a multi-line tldr
  const lines = t.split("\n");
  if (lines.length > 1) {
    t = lines.map((ln, i) => (ln.startsWith("•") ? ln : (i === 0 ? ln : ("• " + ln)))).join("\n");
  }
  return t;
}

/* ───────── PDF helpers ───────── */

function rectTLtoBL(page, box) {
  const H = page.getHeight();
  return { x: box.x, y: H - box.y - box.h, w: box.w, h: box.h };
}

// Basic word-wrap text box (single font)
function drawTextBox(page, font, text, box, opts = {}) {
  const t = norm(text);
  if (!t) return;

  const maxLines = Number.isFinite(opts.maxLines) ? opts.maxLines : (box.maxLines ?? 50);
  const lineGap = Number.isFinite(opts.lineGap) ? opts.lineGap : (box.lineGap ?? 2);
  const paraGap = Number.isFinite(opts.paraGap) ? opts.paraGap : 1;
  const pad = Number.isFinite(opts.pad) ? opts.pad : (box.pad ?? 0);

  const size = opts.size ?? box.size ?? 12;
  const align = opts.align ?? box.align ?? "left";

  const { x, y, w, h } = rectTLtoBL(page, box);

  const lines = [];
  const paras = t.split("\n");
  for (const para of paras) {
    const words = para.split(/\s+/).filter(Boolean);
    let line = "";
    for (const word of words) {
      const next = line ? `${line} ${word}` : word;
      const width = font.widthOfTextAtSize(next, size);
      if (width <= (w - pad * 2)) {
        line = next;
      } else {
        if (line) lines.push(line);
        line = word;
      }
      if (lines.length >= maxLines) break;
    }
    if (lines.length >= maxLines) break;
    if (line) lines.push(line);
    if (paras.length > 1 && para !== paras[paras.length - 1] && lines.length < maxLines) {
      for (let g = 0; g < paraGap && lines.length < maxLines; g++) lines.push("");
    }
    if (lines.length >= maxLines) break;
  }

  const lineHeight = size + lineGap;
  let cursorY = y + h - pad - size;

  for (let i = 0; i < lines.length && i < maxLines; i++) {
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

// Draw a bold label on the first line, then body below (no underline)
function drawLabelAndBody(page, fontB, font, label, body, box, bodyOpts = {}) {
  const L = norm(label);
  const B = norm(body);
  if (!L && !B) return;

  const size = box.size ?? 12;
  const lineGap = box.lineGap ?? 2;
  const pad = box.pad ?? 0;
  const maxLines = box.maxLines ?? 50;

  const { x, y, w, h } = rectTLtoBL(page, box);
  const lineHeight = size + lineGap;

  let cursorY = y + h - pad - size;

  // label (bold)
  if (L) {
    page.drawText(L, { x: x + pad, y: cursorY, size, font: fontB, color: rgb(0, 0, 0) });
    cursorY -= lineHeight;
  }

  // body (wrapped)
  if (B) {
    // temporary "box" for body beneath the label
    const bodyBox = {
      x: box.x,
      y: box.y + (L ? lineHeight : 0),
      w: box.w,
      h: box.h - (L ? lineHeight : 0),
      size: box.size,
      align: box.align ?? "left",
      maxLines: Math.max(1, maxLines - (L ? 1 : 0)),
      lineGap: box.lineGap,
      pad: box.pad,
    };
    drawTextBox(page, font, B, bodyBox, { maxLines: bodyBox.maxLines, ...bodyOpts });
  }
}

/* ───────── radar chart embed (QuickChart) ───────── */

function makeSpiderChartUrl12(bandsRaw) {
  const labels = [
    "C_low","C_mid","C_high","T_low","T_mid","T_high",
    "R_low","R_mid","R_high","L_low","L_mid","L_high",
  ];

  const vals = labels.map((k) => Number(bandsRaw?.[k] ?? 0));
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
        borderWidth: 2,
        pointRadius: 2,
      }],
    },
    options: {
      plugins: { legend: { display: false } },
      scales: {
        r: {
          beginAtZero: true,
          min: 0,
          max: 1,
          ticks: { display: false },
          pointLabels: { font: { size: 10 } },
        },
      },
    },
  };

  const q = encodeURIComponent(JSON.stringify(cfg));
  return `https://quickchart.io/chart?width=900&height=900&backgroundColor=transparent&c=${q}`;
}

async function fetchPngBytes(url) {
  const r = await fetch(url);
  if (!r.ok) throw new Error(`chart fetch failed: ${r.status}`);
  const ab = await r.arrayBuffer();
  return new Uint8Array(ab);
}

/* ───────── payload normalisation ───────── */

function normalisePayload(raw) {
  const ctrl = raw?.ctrl || raw?.ct || raw?.payload?.ctrl || {};
  const summary = ctrl?.summary || raw?.ctrl?.summary || raw?.ctrl?.results || {};

  const identity = raw?.identity || summary?.identity || {};
  const fullName = identity?.fullName || raw?.FullName || raw?.fullName || summary?.FullName || "";
  const email = identity?.email || raw?.Email || raw?.email || summary?.Email || "";

  const text = raw?.text || ctrl?.text || {};
  const workWith = raw?.workWith || ctrl?.workWith || {};

  const bands = raw?.bands || ctrl?.bands || summary?.ctrl12 || raw?.ctrl12 || {};
  const dominantKey = raw?.dominantKey || ctrl?.dominantKey || raw?.domKey || "";
  const secondKey = raw?.secondKey || ctrl?.secondKey || raw?.secKey || "";
  const templateKey = raw?.templateKey || ctrl?.templateKey || raw?.tplKey || "";

  const P = {};

  // p1
  P["p1:name"] = norm(fullName);
  P["p1:date"] = norm(raw?.dateLbl || identity?.dateLabel || raw?.dateLabel || "");

  // header name
  P["hdrName"] = norm(fullName);

  // p3
  P["p3:tldr"] = normaliseTldrBullets(text.execSummary_tldr || text.p3_exec_tldr || "");
  P["p3:exec"] = norm(text.execSummary || text.p3_exec || "");
  P["p3:act"]  = norm(text.execSummary_tipact || text.p3_exec_tipact || "");

  // p4
  P["p4:tldr"] = normaliseTldrBullets(text.state_tldr || text.p4_state_tldr || "");
  P["p4:dom"]  = norm(text.domState || text.p4_dom || "");
  P["p4:bottom"] = norm(text.bottomState || text.p4_bottom_state || "");
  P["p4:act"]  = norm(text.state_tipact || text.p4_state_tipact || "");

  // p5
  P["p5:tldr"] = normaliseTldrBullets(text.frequency_tldr || text.p5_freq_tldr || "");
  P["p5:freq"] = norm(text.frequency || text.p5_freq || "");

  // p6
  P["p6:tldr"] = normaliseTldrBullets(text.sequence_tldr || text.p6_seq_tldr || "");
  P["p6:seq"]  = norm(text.sequence || text.p6_seq || "");
  P["p6:act"]  = norm(text.sequence_tipact || text.p6_seq_tipact || "");

  // p7
  P["p7:tldr"] = normaliseTldrBullets(text.theme_tldr || text.p7_theme_tldr || "");
  P["p7:theme"] = norm(text.theme || text.p7_theme || "");
  P["p7:act"]  = norm(text.theme_tipact || text.p7_theme_tipact || "");
  P["p7:themesLow"] = norm(text.themeLow || text.p7_theme_low || "");

  // p8
  P["p8:collabC"] = norm(workWith?.concealed || "");
  P["p8:collabT"] = norm(workWith?.triggered || "");
  P["p8:collabR"] = norm(workWith?.regulated || "");
  P["p8:collabL"] = norm(workWith?.lead || "");

  // p9
  P["p9:actAnchor"] = norm(text.act_anchor || text.p9_act_anchor || "");

  return { P, fullName: P["p1:name"], email, bands, dominantKey, secondKey, templateKey };
}

/* ───────── template loading ───────── */

async function loadTemplateBytes(templateKey) {
  if (!VALID_TEMPLATE_KEYS.has(templateKey)) {
    throw new Error(`invalid templateKey "${templateKey}" (expected one of ${Array.from(VALID_TEMPLATE_KEYS).join(", ")})`);
  }
  const file = `${TEMPLATE_PREFIX}${templateKey}${TEMPLATE_SUFFIX}`;
  const fp = path.join(PUBLIC_DIR, file);
  return fs.readFile(fp);
}

/* ───────── main handler ───────── */

export default async function handler(req, res) {
  try {
    const url = new URL(req.url, "http://localhost");
    const debugMode = ["1","true","yes"].includes(String(url.searchParams.get("debug") || "").toLowerCase());

    let raw = null;

    const dataB64 = url.searchParams.get("data");
    if (dataB64) {
      try {
        const json = Buffer.from(decodeURIComponent(dataB64), "base64").toString("utf8");
        raw = JSON.parse(json);
      } catch (e) {
        const json = Buffer.from(dataB64, "base64").toString("utf8");
        raw = JSON.parse(json);
      }
    } else if (req.method === "POST") {
      let body = "";
      for await (const chunk of req) body += chunk;
      raw = body ? JSON.parse(body) : {};
    } else {
      raw = {};
    }

    const { P, fullName, bands, dominantKey, secondKey, templateKey: tplFromPayload } = normalisePayload(raw);

    const domKey = String(dominantKey || "").toUpperCase().slice(0, 1);
    const secKey = String(secondKey || "").toUpperCase().slice(0, 1);
    const templateKey = (tplFromPayload && String(tplFromPayload).toUpperCase()) || `${domKey}${secKey}`;

    const payloadLayout = raw?.layout || raw?.ctrl?.layout || raw?.ct?.layout || null;
    const layout = mergeLayout(DEFAULT_LAYOUT, payloadLayout);

    const { applied: layoutApplied, ignored: layoutIgnored } = applyLayoutOverridesFromUrl(layout, url);

    if (debugMode) {
      res.setHeader("Content-Type", "application/json; charset=utf-8");
      res.status(200).send(JSON.stringify({
        ok: true,
        v: 6,
        templateKey,
        identity: { fullName },
        layout,
        layoutOverrides: { applied: layoutApplied, ignored: layoutIgnored },
      }, null, 2));
      return;
    }

    const templateBytes = await loadTemplateBytes(templateKey);
    const pdfDoc = await PDFDocument.load(templateBytes);

    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontB = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    const pages = pdfDoc.getPages();
    const L = layout.pages;

    // p1
    if (pages[0] && L.p1) {
      drawTextBox(pages[0], fontB, P["p1:name"], L.p1.name, { maxLines: L.p1.name.maxLines });
      drawTextBox(pages[0], font,  P["p1:date"], L.p1.date, { maxLines: L.p1.date.maxLines });
    }

    // p2–p10 header name
    for (let i = 1; i < pages.length; i++) {
      const pk = `p${i+1}`;
      if (L[pk]?.hdrName) drawTextBox(pages[i], font, P["hdrName"], L[pk].hdrName, { maxLines: 1 });
    }

    // p3 (physical page index 2)
    if (pages[2]) {
      if (L.p3TLDR?.domDesc) drawLabelAndBody(pages[2], fontB, font, "TLDR", P["p3:tldr"], L.p3TLDR.domDesc, { lineGap: 2, paraGap: 0 });
      if (L.p3main?.domDesc) drawLabelAndBody(pages[2], fontB, font, "",     P["p3:exec"], L.p3main.domDesc);
      if (L.p3act?.domDesc)  drawLabelAndBody(pages[2], fontB, font, "Key action", P["p3:act"], L.p3act.domDesc);
    }

    // p4 (physical page index 3)
    if (pages[3]) {
      const main = [P["p4:dom"], P["p4:bottom"]].filter(Boolean).join("\n\n");
      if (L.p4TLDR?.spider) drawLabelAndBody(pages[3], fontB, font, "TLDR", P["p4:tldr"], L.p4TLDR.spider, { lineGap: 2, paraGap: 0 });
      if (L.p4main?.spider) drawLabelAndBody(pages[3], fontB, font, "",     main,         L.p4main.spider);
      if (L.p4act?.spider)  drawLabelAndBody(pages[3], fontB, font, "Key action", P["p4:act"], L.p4act.spider);
    }

    // p5 (physical page index 4) + chart
    if (pages[4]) {
      if (L.p5TLDR?.seqpat) drawLabelAndBody(pages[4], fontB, font, "TLDR", P["p5:tldr"], L.p5TLDR.seqpat, { lineGap: 2, paraGap: 0 });
      if (L.p5main?.seqpat) drawLabelAndBody(pages[4], fontB, font, "",     P["p5:freq"], L.p5main.seqpat);

      if (L.p5?.chart) {
        const chartUrl = makeSpiderChartUrl12(bands || {});
        try {
          const png = await fetchPngBytes(chartUrl);
          const img = await pdfDoc.embedPng(png);
          const { x, y, w, h } = rectTLtoBL(pages[4], { ...L.p5.chart, h: L.p5.chart.h ?? 300, w: L.p5.chart.w ?? 300 });
          pages[4].drawImage(img, { x, y, width: w, height: h });
        } catch {}
      }
    }

    // p6 (physical page index 5)
    if (pages[5]) {
      if (L.p6TLDR?.themeExpl) drawLabelAndBody(pages[5], fontB, font, "TLDR", P["p6:tldr"], L.p6TLDR.themeExpl, { lineGap: 2, paraGap: 0 });
      if (L.p6main?.themeExpl) drawLabelAndBody(pages[5], fontB, font, "",     P["p6:seq"],  L.p6main.themeExpl);
      if (L.p6act?.themeExpl)  drawLabelAndBody(pages[5], fontB, font, "Key action", P["p6:act"], L.p6act.themeExpl);
    }

    // p7 (physical page index 6)
    if (pages[6]) {
      if (L.p7TLDR?.themesTop) drawLabelAndBody(pages[6], fontB, font, "TLDR", P["p7:tldr"],  L.p7TLDR.themesTop, { lineGap: 2, paraGap: 0 });
      if (L.p7main?.themesTop) drawLabelAndBody(pages[6], fontB, font, "",     P["p7:theme"], L.p7main.themesTop);
      if (L.p7act?.themesTop)  drawLabelAndBody(pages[6], fontB, font, "Key action", P["p7:act"], L.p7act.themesTop);
      if (L.p7?.themesLow)     drawTextBox(pages[6], font, P["p7:themesLow"], L.p7.themesLow, { maxLines: L.p7.themesLow.maxLines });
    }

    // p8
    if (pages[7] && L.p8) {
      if (L.p8.collabC) drawTextBox(pages[7], font, P["p8:collabC"], L.p8.collabC, { maxLines: L.p8.collabC.maxLines });
      if (L.p8.collabT) drawTextBox(pages[7], font, P["p8:collabT"], L.p8.collabT, { maxLines: L.p8.collabT.maxLines });
      if (L.p8.collabR) drawTextBox(pages[7], font, P["p8:collabR"], L.p8.collabR, { maxLines: L.p8.collabR.maxLines });
      if (L.p8.collabL) drawTextBox(pages[7], font, P["p8:collabL"], L.p8.collabL, { maxLines: L.p8.collabL.maxLines });
    }

    // p9
    if (pages[8] && L.p9?.actAnchor) {
      drawTextBox(pages[8], font, P["p9:actAnchor"], L.p9.actAnchor, { maxLines: L.p9.actAnchor.maxLines });
    }

    const pdfBytes = await pdfDoc.save();
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", 'inline; filename="CTRL_PoC_Report.pdf"');
    res.status(200).send(Buffer.from(pdfBytes));
  } catch (err) {
    res.setHeader("Content-Type", "application/json; charset=utf-8");
    res.status(500).send(JSON.stringify({
      ok: false,
      error: String(err?.message || err),
      stack: (err?.stack || "").split("\n").slice(0, 10),
    }, null, 2));
  }
}
