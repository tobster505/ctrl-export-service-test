/**
 * CTRL Coach Export Service · fill-template (Coach flow)
 * Place at: /pages/api/fill-template.js  (ctrl-coach-pdf-service)
 * TL-origin coordinates (pt), pages are 1-based.
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

/* safe JSON parse */
function safeJsonParse(str, fb = {}) {
  try {
    return JSON.parse(str);
  } catch {
    return fb;
  }
}

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
    size = 12,
    align = "left",
    maxLines = 99,
    lineGap = 4,
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

  const maxWidth = w;
  const fontSize = size;
  const spaceWidth = font.widthOfTextAtSize(" ", fontSize);

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

function paintStateHighlight(page3, dom, cfg = {}) {
  if (!page3) return null;
  const b = (cfg.absBoxes && cfg.absBoxes[dom]) || null;
  if (!b) return null;

  const pageH = page3.getHeight();
  const inset = N((cfg.styleByState && cfg.styleByState[dom]?.inset) ?? cfg.highlightInset ?? 4, 4);
  const radius = N((cfg.styleByState && cfg.styleByState[dom]?.radius) ?? cfg.highlightRadius ?? 20, 20);

  const x = N(b.x) + inset;
  const yTop = N(b.y) + inset;
  const w = Math.max(0, N(b.w) - inset * 2);
  const h = Math.max(0, N(b.h) - inset * 2);
  const y = pageH - yTop - h;

  const fillOpacity = clamp(N(cfg.fillOpacity, 0.4), 0, 1);
  const strokeOpacity = clamp(N(cfg.strokeOpacity, 0.9), 0, 1);
  const fillColor = cfg.fillColor || { r: 0.4, g: 0.6, b: 0.9 };
  const strokeColor = cfg.strokeColor || { r: 0.0, g: 0.2, b: 0.6 };

  const path = page3.getContentStream();

  // For simplicity, we just draw a rounded rectangle via rectangles+circles approximation
  page3.drawRectangle({
    x,
    y,
    width: w,
    height: h,
    borderColor: rgb(strokeColor.r, strokeColor.g, strokeColor.b),
    borderWidth: 1.2,
    color: rgb(fillColor.r, fillColor.g, fillColor.b),
    opacity: fillOpacity,
    borderOpacity: strokeOpacity,
  });

  // anchor where we can drop "YOU ARE HERE" label
  const labelX = x + w / 2 + (cfg.labelOffsetX || 0);
  const labelY = y + h + (cfg.labelOffsetY || 12);
  return { labelX, labelY };
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

function parseDataParam(b64ish) {
  if (!b64ish) return {};
  let s = String(b64ish);
  try {
    s = decodeURIComponent(s);
  } catch {}
  s = s.replace(/-/g, "+").replace(/_/g, "/");
  while (s.length % 4) s += "=";
  try {
    return JSON.parse(Buffer.from(s, "base64").toString("utf8"));
  } catch {
    return {};
  }
}

/* GET/POST payload reader (supports ?data= and JSON body) */
async function readPayload(req) {
  const q = req.method === "POST" ? (req.body || {}) : (req.query || {});
  if (q.data) return parseDataParam(q.data);
  if (req.method === "POST" && typeof req.body === "object") return req.body;
  return {};
}

/* normalise all the weird incoming shapes into a friendlier P object */
function normaliseInput(d = {}) {
  const wcol = Array.isArray(d.workwcol)
    ? d.workwcol.map((x) => ({ look: norm(x?.look || ""), work: norm(x?.work || "") }))
    : [];
  const wldr = Array.isArray(d.workwlead)
    ? d.workwlead.map((x) => ({ look: norm(x?.look || ""), work: norm(x?.work || "") }))
    : [];

  const tipsIn = d.tips ?? d.tipsText ?? (d.clientTipsActions && d.clientTipsActions.tips);
  const actsIn =
    d.actions ??
    d.actionsText ??
    (d.patternAction && d.patternAction.action) ??
    (d.clientTipsActions && d.clientTipsActions.actions);

  const tipsList = splitToList(tipsIn).map(cleanBullet).filter(Boolean).slice(0, 2);
  const actsList = splitToList(actsIn).map(cleanBullet).filter(Boolean).slice(0, 2);

  const nameCand =
    (d.person && d.person.fullName) ||
    d["p1:n"] ||
    d.fullName ||
    (d.person && d.person.preferredName) ||
    d.preferredName ||
    d.name;

  const out = {
    raw: d,
    person: d.person || { fullName: nameCand || "" },
    name: nameCand || "",
    dateLbl: d.dateLbl || d.dateLabel || d["p1:d"] || "",
    dom: d.dom || "",
    domChar: d.domChar || "",
    domDesc: d.domDesc || "",
    domState: d.domState || d.dom || "",
    tips: tipsList,
    actions: actsList,
    workwcol: wcol,
    workwlead: wldr,
    chartUrl: d.chartUrl || d["p5:chart"] || "",
    layout: d.layout || null,

    // PoC-specific keys will just flow through as-is; we reference them by
    // their explicit names later (p3:exec, p4:stateDeep, p5:freq, etc.)
    "p1:n": d["p1:n"] || nameCand || "",
    "p1:d": d["p1:d"] || d.dateLbl || "",
    "p3:dom": d["p3:dom"] || d.dom || "",
    "p3:exec": d["p3:exec"] || "",
    "p3:tldr1": d["p3:tldr1"] || "",
    "p3:tldr2": d["p3:tldr2"] || "",
    "p3:tldr3": d["p3:tldr3"] || "",
    "p3:tldr4": d["p3:tldr4"] || "",
    "p3:tldr5": d["p3:tldr5"] || "",
    "p3:tip": d["p3:tip"] || "",
    "p4:stateDeep": d["p4:stateDeep"] || "",
    "p5:freq": d["p5:freq"] || "",
    "p5:chart": d["p5:chart"] || d.chartUrl || "",
    "p6:seq": d["p6:seq"] || "",
    "p7:themesTop": d["p7:themesTop"] || "",
    "p7:themesLow": d["p7:themesLow"] || "",
    "p8:collabC": d["p8:collabC"] || "",
    "p8:collabT": d["p8:collabT"] || "",
    "p8:collabR": d["p8:collabR"] || "",
    "p8:collabL": d["p8:collabL"] || "",
    "p9:action1": d["p9:action1"] || "",
    "p9:action2": d["p9:action2"] || "",
    "p9:closing": d["p9:closing"] || "",
  };

  return out;
}

async function loadTemplateBytesLocal(fname) {
  if (!fname.endsWith(".pdf")) throw new Error(`Invalid template filename: ${fname}`);

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

/* safe page accessors */
const pageOrNull = (pages, idx0) => pages[idx0] ?? null;

/* ───────────── handler ───────────── */
export default async function handler(req, res) {
  try {
    const q = req.method === "POST" ? (req.body || {}) : (req.query || {});
    // Default to PoC template; still allow ?tpl= override from your URL
    const defaultTpl = "CTRL_Perspective_Assessment_Profile_template_PoC.pdf";
    const tpl = S(q.tpl || defaultTpl).replace(/[^A-Za-z0-9._-]/g, "");

    const src = await readPayload(req);
    const P = normaliseInput(src);

    // while stabilising, ignore remote layout (use locked layout in the template)
    const ALLOW_REMOTE_LAYOUT = false;

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

    // Minimal layout anchors (coach/PoC templates may differ; we keep generous guards)
    const L = {
      p1: {
        name: { x: 7, y: 473, w: 500, size: 30, align: "center" },
        date: { x: 210, y: 600, w: 500, size: 25, align: "left" },
      },
      p3: {
        domChar: { x: 272, y: 640, w: 630, size: 23, align: "left", maxLines: 6 },
        domDesc: { x: 25, y: 685, w: 550, size: 18, align: "left", maxLines: 12 },
        state: {
          useAbsolute: true,
          shape: "round",
          highlightInset: 6,
          highlightRadius: 28,
          fillOpacity: 0.45,
          styleByState: {
            C: { radius: 28, inset: 6 },
            T: { radius: 28, inset: 6 },
            R: { radius: 1000, inset: 1 },
            L: { radius: 28, inset: 6 },
          },
          labelByState: {
            C: { x: 60, y: 245 },
            T: { x: 290, y: 244 },
            R: { x: 60, y: 605 },
            L: { x: 290, y: 605 },
          },
          labelText: "YOU ARE HERE",
          labelSize: 10,
          labelColor: { r: 0, g: 0, b: 0 },
          labelOffsetX: 0,
          labelOffsetY: 0,
          labelPadTop: 12,
          labelPadBottom: 12,
          absBoxes: {
            C: { x: 58, y: 258, w: 188, h: 156 },
            T: { x: 299, y: 258, w: 196, h: 156 },
            R: { x: 60, y: 433, w: 188, h: 158 },
            L: { x: 298, y: 430, w: 195, h: 173 },
          },
        },
      },
      p4: {
        spider: { x: 30, y: 585, w: 550, size: 16, align: "left", maxLines: 15 },
        chart: { x: 35, y: 235, w: 540, h: 260 },
      },
      p5: {
        seqpat: { x: 25, y: 250, w: 550, size: 18, align: "left", maxLines: 12 },
      },
      p6: {
        theme: null,
        themeExpl: { x: 25, y: 560, w: 550, size: 18, align: "left", maxLines: 12 },
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
        tips1: { x: 30, y: 175, w: 530, h: 80, size: 18, align: "left", maxLines: 4 },
        tips2: { x: 30, y: 265, w: 530, h: 80, size: 18, align: "left", maxLines: 4 },
        acts1: { x: 30, y: 405, w: 530, h: 80, size: 18, align: "left", maxLines: 4 },
        acts2: { x: 30, y: 495, w: 530, h: 80, size: 18, align: "left", maxLines: 4 },
      },
    };

    /* p1 — cover (unchanged layout) */
    if (p1 && P.name) drawTextBox(p1, font, P.name, L.p1.name);
    if (p1 && P.dateLbl) drawTextBox(p1, font, P.dateLbl, L.p1.date);

    /* p3 — dominant state visual + Exec summary + TLDR + tip */
    (function drawPage3() {
      if (!p3 || !L.p3) return;

      // 1) Dominant highlight (4 boxes)
      const domKey = resolveDomKey(P["p3:dom"] || P.dom, P.domChar, P.domDesc);
      if (domKey && L.p3.state?.useAbsolute) {
        const anchor = paintStateHighlight(p3, domKey, L.p3.state);
        if (anchor && L.p3.state.labelText) {
          drawTextBox(
            p3,
            font,
            String(L.p3.state.labelText),
            {
              x: anchor.labelX,
              y: anchor.labelY,
              w: 180,
              size: L.p3.state.labelSize || 10,
              align: "center",
            },
            { maxLines: 1 }
          );
        }
      }

      // 2) Exec summary + TLDR + tip in main body
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
        drawTextBox(p3, font, body, L.p3.domDesc, { maxLines: L.p3.domDesc.maxLines });
      }
    })();

    /* p4 — deep state / sub-state interpretation (PoC_StateSub_Interpretation) */
    if (p4 && L.p4?.spider && P["p4:stateDeep"]) {
      drawTextBox(
        p4,
        font,
        norm(P["p4:stateDeep"]),
        L.p4.spider,
        { maxLines: L.p4.spider.maxLines }
      );
    }

    /* p5 — spider chart + frequency narrative
       PoC_Frequency_Text → P["p5:freq"]
       Spider chart URL   → P["p5:chart"] (QuickChart URL) */
    if (p5) {
      // Frequency narrative
      if (L.p5?.seqpat && P["p5:freq"]) {
        drawTextBox(
          p5,
          font,
          norm(P["p5:freq"]),
          L.p5.seqpat,
          { maxLines: L.p5.seqpat.maxLines }
        );
      }

      // Spider chart (reuse chart anchor from page 4 layout)
      let chartUrl = norm(P["p5:chart"] || P.spiderChartUrl || P.chartUrl);
      if (chartUrl && L.p4?.chart) {
        try {
          const u = new URL(chartUrl);
          u.searchParams.set("v", Date.now().toString(36));
          chartUrl = u.toString();
        } catch {}

        const img = await embedRemoteImage(pdfDoc, chartUrl);
        if (img) {
          const H = p5.getHeight();
          const { x, y, w, h } = L.p4.chart;
          p5.drawImage(img, { x, y: H - y - h, width: w, height: h });
        }
      }
    }

    /* p6 — sequence narrative
       PoC_Sequence_Text → P["p6:seq"] */
    if (p6 && L.p6?.themeExpl && P["p6:seq"]) {
      drawTextBox(
        p6,
        font,
        norm(P["p6:seq"]),
        L.p6.themeExpl,
        { maxLines: L.p6.themeExpl.maxLines }
      );
    }

    /* p7 — themes (top + low)
       PoC_Themes_Top → P["p7:themesTop"]
       PoC_Themes_Low → P["p7:themesLow"] */
    if (p7 && Array.isArray(L.p7?.colBoxes) && L.p7.colBoxes.length >= 2) {
      const top = norm(P["p7:themesTop"]);
      const low = norm(P["p7:themesLow"]);

      if (top) {
        const box = L.p7.colBoxes[0]; // left column
        drawTextBox(
          p7,
          font,
          top,
          { x: box.x, y: box.y, w: box.w, size: L.p7.bodySize || 13, align: "left" },
          { maxLines: L.p7.maxLines || 15 }
        );
      }

      if (low) {
        const box = L.p7.colBoxes[1]; // right column
        drawTextBox(
          p7,
          font,
          low,
          { x: box.x, y: box.y, w: box.w, size: L.p7.bodySize || 13, align: "left" },
          { maxLines: L.p7.maxLines || 15 }
        );
      }
    }

    /* p8 — Work-with paragraphs (C/T/R/L)
       PoC_WorkWith_* → P["p8:collabC/T/R/L"] */
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

    /* p9 — Actions + closing note
       PoC_Actions_List        → P["p9:action1"], P["p9:action2"]
       PoC_Actions_ClosingNote → P["p9:closing"] */
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
        drawTextBox(page, font, bullet, spec, { maxLines: spec.maxLines || 4 });
      };

      const tipsSlots = [L.p11.tips1, L.p11.tips2];
      const actsSlots = [L.p11.acts1, L.p11.acts2];

      if (actionsPacked[0]) drawBullet(p9, actsSlots[0], actionsPacked[0]);
      if (actionsPacked[1]) drawBullet(p9, actsSlots[1], actionsPacked[1]);
      if (closing) drawBullet(p9, tipsSlots[0], closing);
    }

    /* Optional: simple footer label on p2..p12 (existing behaviour) */
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
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", 'inline; filename="ctrl-poc-profile.pdf"');
    res.send(Buffer.from(bytes));
  } catch (err) {
    console.error("PDF handler error:", err);
    res.status(500).json({ error: "Failed to generate PDF", detail: err?.message || String(err) });
  }
}
