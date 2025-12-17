// === CTRL PoC fill-template — V5 ===
// Changes vs V4 (targeted):
// 1) URL layout overrides now support BOTH styles:
//    - L_p1_name_x=45 (recommended)
//    - p1_namex=45 or p1_name_x=45 (legacy)
//    Applied/ignored overrides are echoed when ?debug=1
// 2) Packed narrative blocks now add safe section titles with simple hyphen underlines
//    (avoids WinAnsi errors from box-drawing characters like '─').
// 3) Extra text sanitisation to avoid WinAnsi-encoding crashes.
//
// IMPORTANT: Keep the rest of the rendering logic and payload mapping the same.

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
  v: 5,
  pages: {
    p1: {
      name: { x: 7, y: 473, w: 500, h: 60, size: 30, align: "center", maxLines: 1 },
      date: { x: 210, y: 600, w: 500, h: 40, size: 25, align: "left", maxLines: 1 },
    },

    // p2–p10 header name (if pages exist)
    p2: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p3: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      domDesc: { x: 25, y: 200, w: 550, h: 900, size: 18, align: "left", maxLines: 38 },
    },
    p4: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      spider: { x: 25, y: 200, w: 550, h: 900, size: 18, align: "left", maxLines: 38 },
    },
    p5: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      seqpat: { x: 25, y: 200, w: 550, h: 900, size: 18, align: "left", maxLines: 38 },
      chart: { x: 48, y: 462, w: 500, h: 300 },
    },
    p6: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      themeExpl: { x: 25, y: 200, w: 550, h: 900, size: 18, align: "left", maxLines: 38 },
    },
    p7: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      themesTop: { x: 30, y: 200, w: 590, h: 900, size: 17, align: "left", maxLines: 42 },
      themesLow: { x: 320, y: 200, w: 300, h: 900, size: 17, align: "left", maxLines: 28 },
    },
    p8: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      collabC: { x: 30, y: 200, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
      collabT: { x: 320, y: 200, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
      collabR: { x: 30, y: 650, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
      collabL: { x: 320, y: 650, w: 300, h: 420, size: 17, align: "left", maxLines: 14 },
    },
    p9: {
      hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 },
      actAnchor: { x: 25, y: 200, w: 550, h: 220, size: 20, align: "left", maxLines: 8 },
    },
    // p10: keep header override-ready if your template has 10 pages
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

/* ───────── URL layout overrides (V5) ───────── */

// Supported param styles:
//
// 1) L_p1_name_x=45   (recommended)
// 2) p1_name_x=45     (legacy-alt)
// 3) p1_namex=45      (legacy)
//
// Notes:
// - pageKey must be like p1..p10 (string)
// - boxKey must match the layout box name (e.g. name, date, domDesc, seqpat, chart)
// - prop is x|y|w|h|size|maxLines|align|... any key in the box object
//
function parseLayoutParamKey(key) {
  if (!key || typeof key !== "string") return null;

  // Strip optional leading "L_"
  const k = key.startsWith("L_") || key.startsWith("l_") ? key.slice(2) : key;

  // Style 1/2: p1_name_x
  const bits = k.split("_");
  if (bits.length >= 3 && /^p\d+$/i.test(bits[0])) {
    const pageKey = bits[0];
    const boxKey = bits[1];
    const prop = bits.slice(2).join("_");
    return { pageKey, boxKey, prop };
  }

  // Style 3: p1_namex (or p5_chartw)
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

    // Only allow properties that already exist in the box,
    // OR a safe allow-list for numeric geometry + text config.
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

  // Strip common trouble-makers
  out = out.replace(/\u00A0|\u202F/g, " ");              // NBSP / narrow NBSP
  out = out.replace(/\uFEFF/g, "");                     // BOM
  out = out.replace(/[\u2500-\u257F]/g, "-");           // box drawing
  out = out.replace(/\u2013|\u2014|\u2212/g, "-");       // en/em/minus
  out = out.replace(/\u2018|\u2019/g, "'");             // curly apostrophes
  out = out.replace(/\u201C|\u201D/g, '"');             // curly quotes
  out = out.replace(/\u2026/g, "...");                  // ellipsis

  // Remove stray control chars (keep \n and \t)
  out = out.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, " ");

  return out;
}

function norm(s) {
  const t = safeText(s);
  return t.trim();
}

function titleBlock(title, body) {
  const t = norm(title);
  const b = norm(body);
  if (!b) return ""; // do not show empty sections
  const underline = "-".repeat(Math.min(Math.max(t.length, 4), 40));
  return `${t}\n${underline}\n${b}`;
}

function packTriptych({ tldr, main, act, mainTitle }) {
  const blocks = [];
  blocks.push(titleBlock("TLDR", tldr));
  blocks.push(titleBlock(mainTitle || "Summary", main));
  blocks.push(titleBlock("Key Action", act));
  return blocks.filter(Boolean).join("\n\n");
}

/* ───────── PDF helpers ───────── */

function rectTLtoBL(page, box) {
  const H = page.getHeight();
  return { x: box.x, y: H - box.y - box.h, w: box.w, h: box.h };
}

function drawTextBox(page, font, text, box, opts = {}) {
  const t = norm(text);
  if (!t) return;

  const maxLines = Number.isFinite(opts.maxLines) ? opts.maxLines : (box.maxLines ?? 50);
  const lineGap = Number.isFinite(opts.lineGap) ? opts.lineGap : (box.lineGap ?? 2);
  const pad = Number.isFinite(opts.pad) ? opts.pad : (box.pad ?? 0);

  const size = opts.size ?? box.size ?? 12;
  const align = opts.align ?? box.align ?? "left";

  const { x, y, w, h } = rectTLtoBL(page, box);

  // primitive word wrap
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
    // preserve blank line between paragraphs (but not after last)
    if (paras.length > 1 && para !== paras[paras.length - 1] && lines.length < maxLines) {
      lines.push("");
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
  P["p3:tldr"] = norm(text.execSummary_tldr || text.p3_exec_tldr || "");
  P["p3:exec"] = norm(text.execSummary || text.p3_exec || "");
  P["p3:act"]  = norm(text.execSummary_tipact || text.p3_exec_tipact || "");

  // p4
  P["p4:tldr"] = norm(text.state_tldr || text.p4_state_tldr || "");
  P["p4:dom"]  = norm(text.domState || text.p4_dom || "");
  P["p4:bottom"] = norm(text.bottomState || text.p4_bottom_state || "");
  P["p4:act"]  = norm(text.state_tipact || text.p4_state_tipact || "");

  // p5
  P["p5:tldr"] = norm(text.frequency_tldr || text.p5_freq_tldr || "");
  P["p5:freq"] = norm(text.frequency || text.p5_freq || "");

  // p6
  P["p6:tldr"] = norm(text.sequence_tldr || text.p6_seq_tldr || "");
  P["p6:seq"]  = norm(text.sequence || text.p6_seq || "");
  P["p6:act"]  = norm(text.sequence_tipact || text.p6_seq_tipact || "");

  // p7
  P["p7:tldr"] = norm(text.theme_tldr || text.p7_theme_tldr || "");
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

    // payload: prefer ?data= base64(json). Fallback to JSON body.
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

    // Template selection: prefer explicit templateKey; else build from dom+second (e.g. T + L => TL)
    const domKey = String(dominantKey || "").toUpperCase().slice(0, 1);
    const secKey = String(secondKey || "").toUpperCase().slice(0, 1);
    const templateKey = (tplFromPayload && String(tplFromPayload).toUpperCase()) || `${domKey}${secKey}`;

    // Layout: default -> payload override -> URL overrides
    const payloadLayout = raw?.layout || raw?.ctrl?.layout || raw?.ct?.layout || null;
    const layout = mergeLayout(DEFAULT_LAYOUT, payloadLayout);

    const { applied: layoutApplied, ignored: layoutIgnored } = applyLayoutOverridesFromUrl(layout, url);

    if (debugMode) {
      res.setHeader("Content-Type", "application/json; charset=utf-8");
      res.status(200).send(JSON.stringify({
        ok: true,
        v: 5,
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
      drawTextBox(pages[0], font, P["p1:date"], L.p1.date, { maxLines: L.p1.date.maxLines });
    }

    // p2–p10 header name
    for (let i = 1; i < pages.length; i++) {
      const pk = `p${i+1}`;
      if (L[pk]?.hdrName) drawTextBox(pages[i], font, P["hdrName"], L[pk].hdrName, { maxLines: 1 });
    }

    // p3
    if (pages[2] && L.p3?.domDesc) {
      const packed = packTriptych({
        tldr: P["p3:tldr"],
        main: P["p3:exec"],
        act: P["p3:act"],
        mainTitle: "Executive Summary",
      });
      drawTextBox(pages[2], font, packed, L.p3.domDesc, { maxLines: L.p3.domDesc.maxLines });
    }

    // p4
    if (pages[3] && L.p4?.spider) {
      const main = [P["p4:dom"], P["p4:bottom"]].filter(Boolean).join("\n\n");
      const packed = packTriptych({
        tldr: P["p4:tldr"],
        main,
        act: P["p4:act"],
        mainTitle: "State Deep Dive",
      });
      drawTextBox(pages[3], font, packed, L.p4.spider, { maxLines: L.p4.spider.maxLines });
    }

    // p5 + chart
    if (pages[4] && L.p5) {
      const packed = packTriptych({
        tldr: P["p5:tldr"],
        main: P["p5:freq"],
        act: "",
        mainTitle: "Frequency",
      });
      if (L.p5.seqpat && packed) drawTextBox(pages[4], font, packed, L.p5.seqpat, { maxLines: L.p5.seqpat.maxLines });

      if (L.p5.chart) {
        const chartUrl = makeSpiderChartUrl12(bands || {});
        try {
          const png = await fetchPngBytes(chartUrl);
          const img = await pdfDoc.embedPng(png);
          const { x, y, w, h } = rectTLtoBL(pages[4], { ...L.p5.chart, h: L.p5.chart.h ?? 300, w: L.p5.chart.w ?? 300 });
          pages[4].drawImage(img, { x, y, width: w, height: h });
        } catch {}
      }
    }

    // p6
    if (pages[5] && L.p6?.themeExpl) {
      const packed = packTriptych({
        tldr: P["p6:tldr"],
        main: P["p6:seq"],
        act: P["p6:act"],
        mainTitle: "Sequence",
      });
      drawTextBox(pages[5], font, packed, L.p6.themeExpl, { maxLines: L.p6.themeExpl.maxLines });
    }

    // p7
    if (pages[6] && L.p7) {
      const packed = packTriptych({
        tldr: P["p7:tldr"],
        main: P["p7:theme"],
        act: P["p7:act"],
        mainTitle: "Themes",
      });
      if (L.p7.themesTop) drawTextBox(pages[6], font, packed, L.p7.themesTop, { maxLines: L.p7.themesTop.maxLines });
      if (L.p7.themesLow) drawTextBox(pages[6], font, P["p7:themesLow"], L.p7.themesLow, { maxLines: L.p7.themesLow.maxLines });
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
