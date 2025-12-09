/**
 * CTRL PoC Export Service · fill-template (Starter/PoC flow)
 * Place at: /api/fill-template.js  (ctrl-poc-service)
 */
export const config = { runtime: "nodejs" };

/* ───────────── imports ───────────── */
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

/* ───────────── utilities ───────────── */
const S = (v, fb = "") => (v == null ? String(fb) : String(v));
const N = (v, fb = 0) => (Number.isFinite(+v) ? +v : fb);
const clamp = (v, lo, hi) => (v < lo ? lo : v > hi ? hi : v);

const norm = (s) =>
  S(s || "")
    .replace(/\r\n?/g, "\n")
    .replace(/\t/g, " ")
    .replace(/[ \f\v]+/g, " ")
    .replace(/[ \t]+\n/g, "\n")
    .trim();

const splitToList = (s) =>
  norm(s)
    .split(/\n+/)
    .map((x) => x.trim())
    .filter(Boolean);

const cleanBullet = (s) =>
  S(s || "")
    .replace(/^\s*[-–—•·]\s*/u, "")
    .trim();

/* brand colour for CTRL magenta (#A64E8C) */
const BRAND = { r: 166 / 255, g: 78 / 255, b: 140 / 255 };

/* embed remote image (QuickChart etc.) */
async function embedRemoteImage(pdfDoc, url) {
  if (!url) return null;
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const buf = await res.arrayBuffer();
    const arr = new Uint8Array(buf);
    // naive mime sniff
    const sig = String.fromCharCode(arr[0], arr[1], arr[2], arr[3] || 0);
    if (sig.startsWith("\x89PNG")) {
      return await pdfDoc.embedPng(arr);
    }
    if (sig.startsWith("\xff\xd8")) {
      return await pdfDoc.embedJpg(arr);
    }
    // fallback: try both
    try {
      return await pdfDoc.embedPng(arr);
    } catch {
      return await pdfDoc.embedJpg(arr);
    }
  } catch {
    return null;
  }
}

/* text wrapping */
function drawTextBox(page, font, text, box, opts = {}) {
  if (!page || !font || !box) return;
  const raw = S(text || "");
  if (!raw) return;

  const {
    size = box.size || 12,
    align = box.align || "left",
    maxLines = opts.maxLines ?? box.maxLines ?? 99,
    lineGap = box.lineGap ?? 4,
    color = rgb(0, 0, 0),
  } = box;

  const txt = norm(raw);
  if (!txt) return;

  const pageH = page.getHeight();
  const x = N(box.x, 0);
  const yTop = N(box.y, 0);
  const w = N(box.w, 500);

  const words = txt.split(/\s+/);
  const lines = [];
  let current = "";

  const fontSize = size;
  const maxWidth = w;

  for (const word of words) {
    const testLine = current ? current + " " + word : word;
    const testWidth = font.widthOfTextAtSize(testLine, fontSize);
    if (testWidth <= maxWidth || !current) {
      current = testLine;
    } else {
      lines.push(current);
      current = word;
      if (lines.length >= maxLines) break;
    }
  }
  if (current && lines.length < maxLines) lines.push(current);

  const totalHeight = lines.length * fontSize + (lines.length - 1) * lineGap;
  let yStart = pageH - yTop - fontSize;
  if (box.valign === "middle" || box.valign === "center") {
    yStart = pageH - yTop - totalHeight / 2;
  } else if (box.valign === "bottom") {
    yStart = pageH - yTop - totalHeight;
  }

  lines.forEach((ln, idx) => {
    const lineWidth = font.widthOfTextAtSize(ln, fontSize);
    let drawX = x;
    if (align === "center") {
      drawX = x + (w - lineWidth) / 2;
    } else if (align === "right") {
      drawX = x + (w - lineWidth);
    }
    const drawY = yStart - idx * (fontSize + lineGap);
    page.drawText(ln, { x: drawX, y: drawY, size: fontSize, font, color });
  });
}

/* convert TL coords to BL rect */
const rectTLtoBL = (page, box, inset = 0) => {
  const pageH = page.getHeight();
  const x = N(box.x) + inset;
  const w = Math.max(0, N(box.w) - inset * 2);
  const h = Math.max(0, N(box.h) - inset * 2);
  const y = pageH - N(box.y) - N(box.h) + inset;
  return { x, y, w, h };
};

/* L-shaped magenta “shadow” under a card */
function drawShadowL(page, absBox, strength = 1) {
  if (!page || !absBox) return;

  // Convert TL box to BL coords once
  const rect = rectTLtoBL(page, absBox, 0);
  const x = rect.x;
  const y = rect.y;
  const w = rect.w;
  const h = rect.h;

  // Tune these numbers to match your template shadow thickness
  const sideWidth = 18 * strength; // right-hand bar
  const baseHeight = 18 * strength; // bottom bar

  // bottom bar
  page.drawRectangle({
    x,
    y,
    width: w,
    height: baseHeight,
    color: rgb(BRAND.r, BRAND.g, BRAND.b),
  });

  // right-hand bar (goes up the full card height and sits on the base)
  page.drawRectangle({
    x: x + w - sideWidth,
    y,
    width: sideWidth,
    height: h + baseHeight,
    color: rgb(BRAND.r, BRAND.g, BRAND.b),
  });
}

function resolveDomKey(dom, domChar, domDesc) {
  const d = S(dom || "").trim().charAt(0).toUpperCase();
  if (["C", "T", "R", "L"].includes(d)) return d;
  const s = S(domChar || domDesc || "").toLowerCase();
  if (/concealed/.test(s)) return "C";
  if (/triggered/.test(s)) return "T";
  if (/regulated/.test(s)) return "R";
  if (/lead/.test(s)) return "L";
  return "R";
}

/* ───────────── robust data parser ───────────── */

function parseDataParam(raw) {
  if (!raw) return {};

  // 1) Try direct JSON (Vercel often gives decoded query)
  try {
    return JSON.parse(String(raw));
  } catch {
    // ignore
  }

  // 2) Try JSON after decodeURIComponent
  try {
    const decoded = decodeURIComponent(String(raw));
    return JSON.parse(decoded);
  } catch {
    // ignore
  }

  // 3) Try base64-encoded JSON (legacy behaviour)
  try {
    let b64 = String(raw).replace(/-/g, "+").replace(/_/g, "/");
    while (b64.length % 4) b64 += "=";
    const txt = Buffer.from(b64, "base64").toString("utf8");
    return JSON.parse(txt);
  } catch {
    // ignore
  }

  return {};
}

/* GET/POST payload reader */
async function readPayload(req) {
  if (req.method === "POST") {
    const body = req.body || {};
    if (body.data) return parseDataParam(body.data);
    if (typeof body === "object" && !Array.isArray(body)) return body;
    return {};
  }

  // GET
  const q = req.query || {};
  if (q.data) return parseDataParam(q.data);
  return {};
}

/* ───────────── normalise new CTRL PoC payload ───────────── */
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
    d.preferredName ||
    d.name ||
    "";

  const dateLbl =
    d.dateLbl ||
    identity.dateLabel ||
    identity.dateLbl ||
    d.dateLabel ||
    d["p1:d"] ||
    "";

  // dominant + second state (accepts new ctrl.dominant)
  const domState =
    ctrl.dominant ||
    ctrl.dominantState ||
    d.domState ||
    d.dom ||
    d["p3:dom"] ||
    "";

  const dom2State =
    ctrl.secondState ||
    d.dom2 ||
    d["p3:dom2"] ||
    "";

  // counts (4 state): from ctrl.counts OR summary.counts/stateFrequency
  const counts =
    ctrl.counts ||
    summary.counts ||
    summary.stateFrequency ||
    d.counts ||
    d["p3:counts"] ||
    {};

  const order = ctrl.order || d.order || d["p3:order"] || "";

  const tldrRaw = text.tldr || d["p3:tldr"] || "";
  const tldrLines = splitToList(tldrRaw).map(cleanBullet).slice(0, 5);

  const actsIn = actionsObj.list || d.actions || d.actionsText || [];
  const actsList = Array.isArray(actsIn)
    ? actsIn.map((a) => cleanBullet(a.text || a))
    : splitToList(actsIn).map(cleanBullet);
  const act1 = actsList[0] || d["p9:action1"] || "";
  const act2 = actsList[1] || d["p9:action2"] || "";

  const chartUrl =
    chart.spiderUrl ||
    d.spiderChartUrl ||
    d.spiderChartURL ||
    d.chartUrl ||
    d["p5:chart"] ||
    "";

  const out = {
    raw: d,
    person: d.person || { fullName: nameCand || "" },
    name: nameCand || "",
    dateLbl,
    dom: domState || "",
    dom2: dom2State || "",
    domChar: d.domChar || "",
    domDesc: d.domDesc || "",
    domState,
    counts,
    order,
    tips: [],
    actions: actsList,
    chartUrl,
    layout: d.layout || null,

    // passthrough for any 12-band data if you want later
    bands: ctrl.bands || summary.bands || d.bands || {},

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

    "p5:freq": d["p5:freq"] || text.frequency || "",
    "p5:chart": d["p5:chart"] || chartUrl || "",

    "p6:seq": d["p6:seq"] || text.sequence || "",

    "p7:themesTop": d["p7:themesTop"] || text.themesTop || "",
    "p7:themesLow": d["p7:themesLow"] || text.themesLow || "",

    "p8:collabC": d["p8:collabC"] || workWith.concealed || "",
    "p8:collabT": d["p8:collabT"] || workWith.triggered || "",
    "p8:collabR": d["p8:collabR"] || workWith.regulated || "",
    "p8:collabL": d["p8:collabL"] || workWith.lead || "",

    "p9:action1": act1,
    "p9:action2": act2,
    "p9:closing": d["p9:closing"] || actionsObj.closingNote || "",
  };

  return out;
}

/* ───────────── template + asset loaders ───────────── */

async function loadTemplateBytesLocal(fname) {
  if (!fname.endsWith(".pdf"))
    throw new Error(`Invalid template filename: ${fname}`);

  const __file = fileURLToPath(import.meta.url);
  const __dir = path.dirname(__file);

  const candidates = [
    path.join(__dir, "..", "..", "public", fname),
    path.join(__dir, "..", "public", fname),
    path.join(__dir, "public", fname),
    path.join(process.cwd(), "public", fname),
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
    `Template not found in any known path for /public: ${fname} (${lastErr?.message || "no detail"})`
  );
}

/* generic asset loader (kept for future assets if needed) */
async function loadAssetBytes(fname) {
  const __file = fileURLToPath(import.meta.url);
  const __dir = path.dirname(__file);

  const candidates = [
    path.join(__dir, "..", "..", "public", fname),
    path.join(__dir, "..", "public", fname),
    path.join(__dir, "public", fname),
    path.join(process.cwd(), "public", fname),
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
    `Asset not found in any known path for /public: ${fname} (${lastErr?.message || "no detail"})`
  );
}

/* safe page accessors */
const pageOrNull = (pages, idx0) => pages[idx0] ?? null;

/* ───────────── handler ───────────── */
export default async function handler(req, res) {
  try {
    const q = req.method === "POST" ? (req.body || {}) : (req.query || {});
    const defaultTpl = "CTRL_PoC_Assessment_Profile_template.pdf";
    const tpl = S(q.tpl || defaultTpl).replace(/[^A-Za-z0-9._-]/g, "");

    const src = await readPayload(req);
    const P = normaliseInput(src);

    const pdfBytes = await loadTemplateBytesLocal(tpl);
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

    const pages = pdfDoc.getPages();
    const p1 = pageOrNull(pages, 0);
    const p2 = pageOrNull(pages, 1);
    const p3 = pageOrNull(pages, 2);
    const p4 = pageOrNull(pages, 3);
    const p5 = pageOrNull(pages, 4);
    const p6 = pageOrNull(pages, 5);
    const p7 = pageOrNull(pages, 6);
    const p8 = pageOrNull(pages, 7);
    const p9 = pageOrNull(pages, 8);
    const p10 = pageOrNull(pages, 9);
    const p11 = pageOrNull(pages, 10);
    const p12 = pageOrNull(pages, 11);

    // Layout anchors (tl coords in pt)
    const L = {
      p1: {
        name: { x: 7, y: 473, w: 500, size: 30, align: "center" },
        date: { x: 210, y: 600, w: 500, size: 25, align: "left" },
      },
      p3: {
        domChar: {
          x: 272,
          y: 640,
          w: 630,
          size: 23,
          align: "left",
          maxLines: 6,
        },
        domDesc: {
          x: 25,
          y: 685,
          w: 550,
          size: 18,
          align: "left",
          maxLines: 12,
        },
        state: {
          useAbsolute: true,
          absBoxes: {
            C: { x: 58, y: 258, w: 188, h: 156 },
            T: { x: 299, y: 258, w: 196, h: 156 },
            R: { x: 60, y: 433, w: 188, h: 158 },
            L: { x: 298, y: 430, w: 195, h: 173 },
          },
          labelSize: 12,
          labelOffsetTop: 18,
        },
      },
      p4: {
        spider: {
          x: 30,
          y: 585,
          w: 550,
          size: 16,
          align: "left",
          maxLines: 15,
        },
        chart: { x: 35, y: 235, w: 540, h: 260 },
      },
      p5: {
        seqpat: {
          x: 25,
          y: 250,
          w: 550,
          size: 18,
          align: "left",
          maxLines: 12,
        },
      },
      p6: {
        theme: null,
        themeExpl: {
          x: 25,
          y: 560,
          w: 550,
          size: 18,
          align: "left",
          maxLines: 12,
        },
      },
      p7: {
        colBoxes: [
          { x: 25, y: 330, w: 260, h: 120 },
          { x: 320, y: 330, w: 260, h: 120 },
          { x: 25, y: 595, w: 260, h: 120 },
          { x: 320, y: 595, w: 260, h: 120 },
        ],
        bodySize: 13,
        maxLines: 15,
      },
      p8: {
        colBoxes: [
          { x: 25, y: 330, w: 260, h: 120 },
          { x: 320, y: 330, w: 260, h: 120 },
          { x: 25, y: 595, w: 260, h: 120 },
          { x: 320, y: 595, w: 260, h: 120 },
        ],
        bodySize: 13,
        maxLines: 15,
      },
      p9: {
        ldrBoxes: [
          { x: 25, y: 330, w: 260, h: 120 },
          { x: 320, y: 330, w: 260, h: 120 },
          { x: 25, y: 595, w: 260, h: 120 },
        ],
        bodySize: 13,
        maxLines: 15,
      },
      p10: {
        ldrBoxes: [
          { x: 25, y: 330, w: 260, h: 120 },
          { x: 320, y: 330, w: 260, h: 120 },
          { x: 25, y: 595, w: 260, h: 120 },
        ],
        bodySize: 13,
        maxLines: 15,
      },
      p11: {
        lineGap: 6,
        itemGap: 6,
        bulletIndent: 18,
        tips1: {
          x: 30,
          y: 175,
          w: 530,
          h: 80,
          size: 18,
          align: "left",
          maxLines: 4,
        },
        tips2: {
          x: 30,
          y: 265,
          w: 530,
          h: 80,
          size: 18,
          align: "left",
          maxLines: 4,
        },
        acts1: {
          x: 30,
          y: 405,
          w: 530,
          h: 80,
          size: 18,
          align: "left",
          maxLines: 4,
        },
        acts2: {
          x: 30,
          y: 495,
          w: 530,
          h: 80,
          size: 18,
          align: "left",
          maxLines: 4,
        },
      },
    };

    /* p1 — cover (name + date) */
    if (p1 && L.p1) {
      const nameText =
        norm(P.name || (P.person && P.person.fullName) || P["p1:n"] || "");
      const dateText = norm(P.dateLbl || P.dateLabel || P["p1:d"] || "");

      if (nameText) {
        drawTextBox(p1, font, nameText, L.p1.name, { maxLines: 1 });
      }
      if (dateText) {
        drawTextBox(p1, font, dateText, L.p1.date, { maxLines: 1 });
      }
    }

    /* p3 — dominant / second state shadows + "YOU ARE HERE!" + Exec / TLDR / tip */
    (function drawPage3() {
      if (!p3 || !L.p3) return;

      const stateCfg = L.p3.state || {};
      const raw = P.raw || {};
      const ctrl = raw.ctrl || {};
      const summary = ctrl.summary || {};

      // 1) dominant + second state
      const domKey = resolveDomKey(
        P["p3:dom"] || P.dom || ctrl.dominant || ctrl.dominantState,
        P.domChar,
        P.domDesc
      );

      let secondKey = "";
      if (ctrl.secondState) {
        secondKey = S(ctrl.secondState).trim().charAt(0).toUpperCase();
      } else if (P.dom2) {
        secondKey = S(P.dom2).trim().charAt(0).toUpperCase();
      }

      if (!["C", "T", "R", "L"].includes(secondKey)) {
        const counts =
          P.counts ||
          ctrl.counts ||
          summary.counts ||
          summary.stateFrequency ||
          {};
        const keys = ["C", "T", "R", "L"];
        const ranked = keys
          .map((k) => ({ k, n: Number(counts[k] || 0) || 0 }))
          .sort((a, b) => b.n - a.n);
        const bestNonDom = ranked.find((r) => r.k !== domKey && r.n > 0);
        if (bestNonDom) secondKey = bestNonDom.k;
      }

      const absBoxes = stateCfg.absBoxes || {};

      // 2) 2nd place shadow (lighter / smaller)
      if (secondKey && secondKey !== domKey && absBoxes[secondKey]) {
        drawShadowL(p3, absBoxes[secondKey], 0.7);
      }

      // 3) dominant state shadow (full / thicker)
      if (domKey && absBoxes[domKey]) {
        drawShadowL(p3, absBoxes[domKey], 1.2);
      }

      // 4) "YOU ARE HERE!" label above dominant card
      if (domKey && absBoxes[domKey]) {
        const b = absBoxes[domKey];
        const labelBox = {
          x: b.x,
          y: b.y - (stateCfg.labelOffsetTop || 18),
          w: b.w,
          size: stateCfg.labelSize || 12,
          align: "center",
        };
        drawTextBox(p3, font, "YOU ARE HERE!", labelBox, { maxLines: 1 });
      }

      // 5) Exec summary + TLDR + tip
      const exec = norm(P["p3:exec"]);
      const tldrs = [
        norm(P["p3:tldr1"]),
        norm(P["p3:tldr2"]),
        norm(P["p3:tldr3"]),
        norm(P["p3:tldr4"]),
        norm(P["p3:tldr5"]),
      ].filter(Boolean);
      const tip = norm(P["p3:tip"]);

      const blocks = [];
      if (exec) blocks.push(exec);
      if (tldrs.length) blocks.push(tldrs.join("\n\n"));
      if (tip) blocks.push(tip);

      const body = blocks.join("\n\n\n");
      if (body && L.p3.domDesc) {
        drawTextBox(p3, font, body, L.p3.domDesc, {
          maxLines: L.p3.domDesc.maxLines,
        });
      }
    })();

    /* p4 — state / sub-state deep dive + spider chart */
    if (p4 && L.p4) {
      if (L.p4.spider && P["p4:stateDeep"]) {
        drawTextBox(p4, font, norm(P["p4:stateDeep"]), L.p4.spider, {
          maxLines: L.p4.spider.maxLines,
        });
      }

      // Spider chart on p4
      let chartUrl = norm(P["p5:chart"] || P.spiderChartUrl || P.chartUrl);
      if (chartUrl && L.p4.chart) {
        try {
          const u = new URL(chartUrl);
          u.searchParams.set("v", Date.now().toString(36));
          chartUrl = u.toString();
        } catch {
          // ignore URL parse errors, try as-is
        }
        const img = await embedRemoteImage(pdfDoc, chartUrl);
        if (img) {
          const H = p4.getHeight();
          const { x, y, w, h } = L.p4.chart;
          p4.drawImage(img, { x, y: H - y - h, width: w, height: h });
        }
      }
    }

    /* p5 — frequency narrative */
    if (p5 && L.p5?.seqpat && P["p5:freq"]) {
      drawTextBox(p5, font, norm(P["p5:freq"]), L.p5.seqpat, {
        maxLines: L.p5.seqpat.maxLines,
      });
    }

    /* p6 — sequence narrative */
    if (p6 && L.p6?.themeExpl && P["p6:seq"]) {
      drawTextBox(p6, font, norm(P["p6:seq"]), L.p6.themeExpl, {
        maxLines: L.p6.themeExpl.maxLines,
      });
    }

    /* p7 — themes (top + low) */
    if (p7 && Array.isArray(L.p7?.colBoxes) && L.p7.colBoxes.length >= 2) {
      const top = norm(P["p7:themesTop"]);
      const low = norm(P["p7:themesLow"]);

      if (top) {
        const box = L.p7.colBoxes[0];
        drawTextBox(
          p7,
          font,
          top,
          { x: box.x, y: box.y, w: box.w, size: L.p7.bodySize || 13, align: "left" },
          { maxLines: L.p7.maxLines || 15 }
        );
      }

      if (low) {
        const box = L.p7.colBoxes[1];
        drawTextBox(
          p7,
          font,
          low,
          { x: box.x, y: box.y, w: box.w, size: L.p7.bodySize || 13, align: "left" },
          { maxLines: L.p7.maxLines || 15 }
        );
      }
    }

    /* p8 — Work-with paragraphs (C/T/R/L) */
    if (p8 && Array.isArray(L.p8?.colBoxes) && L.p8.colBoxes.length >= 4) {
      const mapIdx = { C: 0, T: 1, R: 2, L: 3 };
      const txtByState = {
        C: norm(P["p8:collabC"]),
        T: norm(P["p8:collabT"]),
        R: norm(P["p8:collabR"]),
        L: norm(P["p8:collabL"]),
      };

      for (const key of ["C", "T", "R", "L"]) {
        const txt = txtByState[key];
        if (!txt) continue;
        const idx = mapIdx[key];
        const box = L.p8.colBoxes[idx];
        if (!box) continue;

        drawTextBox(
          p8,
          font,
          txt,
          { x: box.x, y: box.y, w: box.w, size: L.p8.bodySize || 13, align: "left" },
          { maxLines: L.p8.maxLines || 15 }
        );
      }
    }

    /* p9 — Actions + closing note */
    if (p9 && L.p11) {
      const tidy = (s) =>
        norm(String(s || ""))
          .replace(/^(?:[-–—•·]\s*)/i, "")
          .replace(/^\s*(tips?|tip)\s*:?\s*/i, "")
          .replace(/^\s*(actions?|next\s*action)\s*:?\s*/i, "")
          .trim();
      const good = (s) => s && s.length >= 3 && !/^tips?$|^actions?$/i.test(s);

      const actionsPacked = [tidy(P["p9:action1"]), tidy(P["p9:action2"])]
        .filter(good)
        .slice(0, 2);

      const closing = tidy(P["p9:closing"]);

      const drawBullet = (page, spec, text) => {
        if (!page || !spec || !text) return;
        const bullet = `• ${text}`;
        drawTextBox(page, font, bullet, spec, {
          maxLines: spec.maxLines || 4,
        });
      };

      const tipsSlots = [L.p11.tips1, L.p11.tips2];
      const actsSlots = [L.p11.acts1, L.p11.acts2];

      if (actionsPacked[0]) drawBullet(p9, actsSlots[0], actionsPacked[0]);
      if (actionsPacked[1]) drawBullet(p9, actsSlots[1], actionsPacked[1]);
      if (closing) drawBullet(p9, tipsSlots[0], closing);
    }

    /* footer label on p2..p12 */
    const footerLabel = norm(P.name);
    const putFooter = (page) => {
      if (!page || !footerLabel) return;
      drawTextBox(
        page,
        font,
        footerLabel,
        { x: 380, y: 51, w: 400, size: 13, align: "left" },
        { maxLines: 1 }
      );
    };
    [p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12].forEach(putFooter);

    // output
    const bytes = await pdfDoc.save();

    const safe = (value, fallback = "") =>
      String(value || fallback)
        .trim()
        .replace(/[^A-Za-z0-9]+/g, "_")
        .replace(/^_+|_+$/g, "");

    const namePart = safe(P.name || P.fullName || "Profile");
    const datePart = safe(P.dateLbl || P.dateLabel || P.date || "");

    const fileName = datePart
      ? `PoC_Profile_${namePart}_${datePart}.pdf`
      : `PoC_Profile_${namePart}.pdf`;

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${fileName}"`);
    res.send(Buffer.from(bytes));
  } catch (err) {
    console.error("PDF handler error:", err);
    res.status(500).json({
      error: "Failed to generate PDF",
      detail: err?.message || String(err),
    });
  }
}
