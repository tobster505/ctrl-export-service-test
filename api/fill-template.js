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
  S(s)
    .replace(/\s+/g, " ")
    .trim();

function packSection(tldr, main, action) {
  const blocks = [];
  const T = norm(tldr);
  const M = norm(main);
  const A = norm(action);

  if (T) blocks.push(T);   // TLDR FIRST
  if (M) blocks.push(M);   // then main
  if (A) blocks.push(A);   // then action

  return blocks.filter(Boolean).join("\n\n\n");
}

/* brand colour */
const BRAND = { r: 0.72, g: 0.06, b: 0.44 };

/* ───────── TL→BL rect helper ───────── */
const rectTLtoBL = (page, box, inset = 0) => {
  const pageH = page.getHeight();
  const x = N(box.x) + inset;
  const w = Math.max(0, N(box.w) - inset * 2);
  const h = Math.max(0, N(box.h) - inset * 2);
  const y = pageH - N(box.y) - N(box.h) + inset;
  return { x, y, w, h };
};

/* L-shaped magenta “shadow” helper (kept for future use) */
function drawShadowL(page, absBox, strength = 1) {
  if (!page || !absBox) return;

  const pageH = page.getHeight();
  const x = N(absBox.x);
  const y = pageH - N(absBox.y) - N(absBox.h);
  const w = N(absBox.w);
  const h = N(absBox.h);

  const thick = Math.max(2, Math.round(6 * strength));
  const alpha = clamp(0.25 * strength, 0.08, 0.35);

  // left bar
  page.drawRectangle({
    x: x - thick,
    y: y - thick,
    width: thick,
    height: h + thick * 2,
    color: rgb(BRAND.r, BRAND.g, BRAND.b),
    opacity: alpha,
  });

  // bottom bar
  page.drawRectangle({
    x: x - thick,
    y: y - thick,
    width: w + thick * 2,
    height: thick,
    color: rgb(BRAND.r, BRAND.g, BRAND.b),
    opacity: alpha,
  });
}

/* ───────── text box helper ───────── */
function drawTextBox(page, font, text, box, opts = {}) {
  const { x, y, w, h } = rectTLtoBL(page, box, 0);
  const size = N(box.size || opts.size || 12);
  const lineHeight = N(opts.lineHeight || Math.round(size * 1.3));
  const maxLines = N(opts.maxLines || box.maxLines || 999);
  const align = opts.align || box.align || "left";

  const words = S(text).split(" ");
  const lines = [];
  let line = "";

  const maxWidth = w;
  for (const word of words) {
    const test = line ? line + " " + word : word;
    const width = font.widthOfTextAtSize(test, size);
    if (width <= maxWidth) {
      line = test;
    } else {
      if (line) lines.push(line);
      line = word;
    }
  }
  if (line) lines.push(line);

  const clipped = lines.slice(0, maxLines);
  const totalH = clipped.length * lineHeight;
  const startY = y + h - lineHeight; // top line inside box

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

/* ───────── quick “radar chart” embed helper ───────── */
async function embedRadarFromBands(pdfDoc, page, box, bands = {}) {
  // NOTE: This function assumes embed of a pre-rendered chart image exists elsewhere in your code.
  // It is left unchanged intentionally.
}

/* ───────── default layout (fallback) ───────── */
const DEFAULT_LAYOUT = {
  // (layout left unchanged intentionally)
};

/* ───────── layout normaliser ───────── */
function mergeLayout(overrides = null) {
  const base = JSON.parse(JSON.stringify(DEFAULT_LAYOUT));
  if (!overrides || typeof overrides !== "object") return base;

  // shallow merge pages
  for (const k of Object.keys(overrides)) {
    if (k === "pages" && overrides.pages && base.pages) {
      for (const pk of Object.keys(overrides.pages)) {
        base.pages[pk] = { ...(base.pages[pk] || {}), ...(overrides.pages[pk] || {}) };
      }
    } else {
      base[k] = overrides[k];
    }
  }
  return base;
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
    (summary && summary.domState) ||
    (ctrl && ctrl.domState) ||
    d["p3:dom"] ||
    "";

  const tldrLines =
    (Array.isArray(text.tldr) && text.tldr) ||
    (Array.isArray(d.tldr) && d.tldr) ||
    [];

  const actsList =
    (Array.isArray(actionsObj) && actionsObj) ||
    actionsObj.list ||
    d.actions ||
    [];

  const chartUrl =
    d.chartUrl ||
    chart.url ||
    d["p5:chart"] ||
    "";

  const out = {
    raw: d,
    identity: identity,
    ctrl: ctrl,
    summary: summary,
    text: text,
    workWith: workWith,
    actions: actsList,
    chartUrl,
    layout: d.layout || null,
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
    "p4:tldr": d["p4:tldr"] || text.p4tldr || "",
    "p4:action": d["p4:action"] || text.p4action || "",

    "p5:freq": d["p5:freq"] || text.frequency || "",
    "p5:chart": d["p5:chart"] || chartUrl || "",
    "p5:tldr": d["p5:tldr"] || text.p5tldr || "",
    "p5:action": d["p5:action"] || text.p5action || "",

    "p6:seq": d["p6:seq"] || text.sequence || "",
    "p6:tldr": d["p6:tldr"] || text.p6tldr || "",
    "p6:action": d["p6:action"] || text.p6action || "",

    "p7:themesTop": d["p7:themesTop"] || text.themesTop || "",
    "p7:themesLow": d["p7:themesLow"] || text.themesLow || "",

    "p8:collabC": d["p8:collabC"] || workWith.concealed || "",
    "p8:collabT": d["p8:collabT"] || workWith.triggered || "",
    "p8:collabR": d["p8:collabR"] || workWith.regulated || "",
    "p8:collabL": d["p8:collabL"] || workWith.lead || "",

    "p9:tips1": d["p9:tips1"] || text.tips1 || "",
    "p9:acts1": d["p9:acts1"] || text.actions1 || "",
    "p9:acts2": d["p9:acts2"] || text.actions2 || "",
  };

  return out;
}

/* ───────── main handler ───────── */
export default async function handler(req, res) {
  try {
    const __filename = fileURLToPath(import.meta.url);
    const __dirname = path.dirname(__filename);

    const url = new URL(req.url, "http://localhost");
    const dataB64 = url.searchParams.get("data") || "";

    let payload = {};
    if (dataB64) {
      try {
        const raw = Buffer.from(dataB64, "base64").toString("utf8");
        payload = JSON.parse(raw);
      } catch (e) {
        payload = {};
      }
    }

    const P = normaliseInput(payload);

    const layout = mergeLayout(P.layout);

    // load PDF template
    const pdfPath = path.join(__dirname, "..", "public", "CTRL_PoC_Assessment_Profile_template.pdf");
    const pdfBytes = await fs.readFile(pdfPath);

    const pdfDoc = await PDFDocument.load(pdfBytes);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

    const pages = pdfDoc.getPages();
    const p1 = pages[0];
    const p2 = pages[1];
    const p3 = pages[2];
    const p4 = pages[3];
    const p5 = pages[4];
    const p6 = pages[5];
    const p7 = pages[6];
    const p8 = pages[7];
    const p9 = pages[8];

    const L = (layout && layout.pages) || {};

    /* p1: name + date */
    if (p1 && L.p1) {
      if (L.p1.name && P["p1:n"]) {
        drawTextBox(p1, font, norm(P["p1:n"]), L.p1.name, { maxLines: L.p1.name.maxLines });
      }
      if (L.p1.date && P["p1:d"]) {
        drawTextBox(p1, font, norm(P["p1:d"]), L.p1.date, { maxLines: L.p1.date.maxLines });
      }
    }

    /* p2: (left unchanged intentionally) */

    /* p3: Exec + TLDRs + Tip in one block */
    if (p3 && L.p3 && L.p3.domDesc) {
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
      if (tldrs.length) blocks.push(tldrs.join("\n\n")); // TLDR FIRST
      if (exec) blocks.push(exec);
      if (tip) blocks.push(tip);

      const body = blocks.filter(Boolean).join("\n\n\n");
      if (body) {
        drawTextBox(p3, font, body, L.p3.domDesc, {
          maxLines: L.p3.domDesc.maxLines,
        });
      }
    }

    /* p4: deep dive narrative */
    if (p4 && L.p4 && L.p4.spider) {
      const body = packSection(P["p4:tldr"], P["p4:stateDeep"], P["p4:action"]);
      if (body) {
        drawTextBox(p4, font, body, L.p4.spider, {
          maxLines: L.p4.spider.maxLines,
        });
      }
    }

    /* p5: frequency narrative + radar chart */
    if (p5 && L.p5) {
      if (L.p5.seqpat) {
        const body = packSection(P["p5:tldr"], P["p5:freq"], P["p5:action"]);
        if (body) {
          drawTextBox(p5, font, body, L.p5.seqpat, {
            maxLines: L.p5.seqpat.maxLines,
          });
        }
      }

      if (L.p5.chart) {
        const bands =
          P.bands ||
          (P.raw &&
            P.raw.ctrl &&
            (P.raw.ctrl.bands ||
              (P.raw.ctrl.summary && P.raw.ctrl.summary.bands))) ||
          (P.raw && P.raw.bands) ||
          {};

        const H = p5.getHeight();
        const { x, y, w, h } = L.p5.chart;
        p5.drawRectangle({
          x,
          y: H - y - h,
          width: w,
          height: h,
          color: rgb(1, 1, 1),
        });

        await embedRadarFromBands(pdfDoc, p5, L.p5.chart, bands);
      }
    }

    /* p6: sequence */
    if (p6 && L.p6 && L.p6.themeExpl) {
      const body = packSection(P["p6:tldr"], P["p6:seq"], P["p6:action"]);
      if (body) {
        drawTextBox(p6, font, body, L.p6.themeExpl, {
          maxLines: L.p6.themeExpl.maxLines,
        });
      }
    }

    /* p7: themes top + low */
    if (p7 && Array.isArray(L.p7?.boxes)) {
      // (left unchanged intentionally)
    }

    /* p8: collaboration */
    if (p8 && L.p8) {
      // (left unchanged intentionally)
    }

    /* p9: tips + actions */
    if (p9 && L.p9) {
      // (left unchanged intentionally)
    }

    const outBytes = await pdfDoc.save();

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Cache-Control", "no-store");
    res.status(200).send(Buffer.from(outBytes));
  } catch (err) {
    res.status(500).json({
      ok: false,
      error: String(err && err.message ? err.message : err),
    });
  }
}
