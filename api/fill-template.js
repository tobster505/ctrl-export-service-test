// V3
/**
 * CTRL PoC Export Service · fill-template (Starter/PoC flow)
 * Place at: /api/fill-template.js  (ctrl-poc-service)
 */
export const config = { runtime: "nodejs" };

import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { PDFDocument, StandardFonts } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";

/* ───────────── utilities ───────────── */
const S = (v, fb = "") => (v == null ? String(fb) : String(v));
const N = (v, fb = 0) => (Number.isFinite(+v) ? +v : fb);
const norm = (s) => S(s).replace(/\s+/g, " ").trim();

function safeJson(obj) {
  try { return JSON.parse(JSON.stringify(obj)); }
  catch { return {}; }
}

function isObj(o) { return o && typeof o === "object" && !Array.isArray(o); }

function getDeep(o, pathArr, fb) {
  try {
    let cur = o;
    for (const k of pathArr) cur = cur?.[k];
    return cur == null ? fb : cur;
  } catch { return fb; }
}

function asText(s) {
  const t = norm(s);
  return t;
}

// Safe underline: use "-" only (WinAnsi safe)
function underlineLine(title) {
  const t = norm(title || "");
  const len = Math.max(6, Math.min(32, t.length));
  return "-".repeat(len); // ASCII-safe for WinAnsi
}


// Convert "A • B • C" OR "A || B || C" OR newline lists into bullet lines
function toBulletLines(raw) {
  const s = S(raw, "").replace(/\r/g, "").trim();
  if (!s) return "";

  // Prefer explicit separators commonly produced by Gen cards
  let parts = [];
  if (s.includes(" • ")) parts = s.split(" • ");
  else if (s.includes("||")) parts = s.split("||");
  else if (s.includes("\n")) parts = s.split("\n");
  else {
    // Fallback: sentence-ish split
    parts = s.split(/(?<=\.)\s+/);
  }

  const cleaned = parts
    .map(p => norm(p).replace(/^[-•\u2022]\s*/, ""))
    .filter(Boolean);

  if (!cleaned.length) return "";

  return cleaned.map(p => `• ${p}`).join("\n");
}

// Merge layout objects shallowly at pages level
function mergeLayout(override) {
  const base = safeJson(DEFAULT_LAYOUT);
  if (!isObj(override)) return base;

  const out = safeJson(base);
  out.pages = out.pages || {};
  const op = override.pages || override; // allow passing {p3:{...}} directly

  for (const [k, v] of Object.entries(op || {})) {
    if (!isObj(v)) continue;
    out.pages[k] = { ...(out.pages[k] || {}), ...v };
  }
  return out;
}

/* ───────────── drawing helpers ───────────── */

function drawTextBox(page, font, text, box, opts = {}) {
  if (!text) return;

  const x = N(box.x);
  const y = N(box.y);
  const w = N(box.w);
  const h = N(box.h);
  const size = N(box.size || opts.size || 12);
  const align = (box.align || opts.align || "left").toLowerCase();
  const maxLines = N(opts.maxLines ?? box.maxLines ?? 50);

  const lineGap = N(opts.lineGap ?? box.gap ?? 2);
  const bulletIndent = N(opts.bulletIndent ?? box.bulletIndent ?? 12);

  // Split lines preserving bullets
  const rawLines = S(text).replace(/\r/g, "").split("\n");

  // Basic wrap by width (pdf-lib uses WinAnsi widths for standard fonts)
  const lines = [];
  for (let raw of rawLines) {
    raw = S(raw);

    // preserve explicit blank lines
    if (!raw.trim()) { lines.push(""); continue; }

    const isBullet = raw.trim().startsWith("•");
    const effectiveW = isBullet ? (w - bulletIndent) : w;

    const words = raw.split(/\s+/);
    let cur = "";
    for (const word of words) {
      const test = cur ? (cur + " " + word) : word;
      const testWidth = font.widthOfTextAtSize(test, size);
      if (testWidth <= effectiveW) {
        cur = test;
      } else {
        if (cur) lines.push(cur);
        cur = word;
      }
    }
    if (cur) lines.push(cur);
  }

  const finalLines = lines.slice(0, maxLines);

  // Start from top
  const lineHeight = size + lineGap;
  let cursorY = y;

  for (let i = 0; i < finalLines.length; i++) {
    const line = finalLines[i] ?? "";

    // Bullet indent handling (only if the line begins with bullet)
    const isBullet = line.trim().startsWith("•");
    const drawX = isBullet ? (x + bulletIndent) : x;

    const lineW = font.widthOfTextAtSize(line, size);
    let tx = drawX;

    if (align === "center") {
      tx = x + (w - lineW) / 2;
    } else if (align === "right") {
      tx = x + w - lineW;
    }

    // Stop if out of box height
    if (h && (cursorY - (i * lineHeight) < y - h)) break;

    page.drawText(line, { x: tx, y: cursorY - (i * lineHeight), size, font });
  }
}

function writeSection(page, font, title, content, boxTitle, boxBody, mode = "paragraph", opts = {}) {
  // Title (underlined)
  const titleText = `${title}\n${underlineLine(title)}`;
  drawTextBox(page, font, titleText, boxTitle, { maxLines: 2 });

  // Body
  let bodyText = content || "";
  if (mode === "bullets") bodyText = toBulletLines(content);

  drawTextBox(page, font, bodyText, boxBody, opts);
}

/* ───────────── layout defaults ───────────── */

const DEFAULT_LAYOUT = {
  pages: {
    // Page numbers are 1-based keys, but we address pdfDoc pages by index (0-based)
    // p1: cover
    p1: {
      name:   { x: 45, y: 453, w: 500, h: 60, size: 30, align: "center", maxLines: 1 },
      email:  { x: 45, y: 424, w: 500, h: 40, size: 14, align: "center", maxLines: 1 },
      date:   { x: 45, y: 405, w: 500, h: 40, size: 14, align: "center", maxLines: 1 },
    },

    // p3: Exec summary
    p3: {
      hdrName: { x: 44, y: 788, w: 530, h: 20, size: 12, align: "left", maxLines: 1 },

      // Title slots
      tldrTitle: { x: 55, y: 680, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      tldrBody:  { x: 55, y: 650, w: 520, h: 130, size: 12, align: "left", maxLines: 10 },

      execTitle: { x: 55, y: 510, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      execBody:  { x: 55, y: 480, w: 520, h: 160, size: 12, align: "left", maxLines: 12 },

      actTitle:  { x: 55, y: 280, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      actBody:   { x: 55, y: 250, w: 520, h: 80,  size: 12, align: "left", maxLines: 6 },
    },

    // p4: state deep dive
    p4: {
      hdrName: { x: 44, y: 788, w: 530, h: 20, size: 12, align: "left", maxLines: 1 },

      tldrTitle: { x: 55, y: 690, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      tldrBody:  { x: 55, y: 660, w: 520, h: 110, size: 12, align: "left", maxLines: 9 },

      execTitle: { x: 55, y: 530, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      execBody:  { x: 55, y: 500, w: 520, h: 150, size: 12, align: "left", maxLines: 11 },

      actTitle:  { x: 55, y: 295, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      actBody:   { x: 55, y: 265, w: 520, h: 80,  size: 12, align: "left", maxLines: 6 },
    },

    // p5: frequency
    p5: {
      hdrName: { x: 44, y: 788, w: 530, h: 20, size: 12, align: "left", maxLines: 1 },

      tldrTitle: { x: 55, y: 690, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      tldrBody:  { x: 55, y: 660, w: 520, h: 110, size: 12, align: "left", maxLines: 9 },

      execTitle: { x: 55, y: 530, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      execBody:  { x: 55, y: 500, w: 520, h: 200, size: 12, align: "left", maxLines: 14 },
    },

    // p6: sequence
    p6: {
      hdrName: { x: 44, y: 788, w: 530, h: 20, size: 12, align: "left", maxLines: 1 },

      tldrTitle: { x: 55, y: 690, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      tldrBody:  { x: 55, y: 660, w: 520, h: 110, size: 12, align: "left", maxLines: 9 },

      execTitle: { x: 55, y: 530, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      execBody:  { x: 55, y: 500, w: 520, h: 150, size: 12, align: "left", maxLines: 11 },

      actTitle:  { x: 55, y: 295, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      actBody:   { x: 55, y: 265, w: 520, h: 80,  size: 12, align: "left", maxLines: 6 },
    },

    // p7: themes
    p7: {
      hdrName: { x: 44, y: 788, w: 530, h: 20, size: 12, align: "left", maxLines: 1 },

      tldrTitle: { x: 55, y: 690, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      tldrBody:  { x: 55, y: 660, w: 520, h: 110, size: 12, align: "left", maxLines: 9 },

      execTitle: { x: 55, y: 530, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      execBody:  { x: 55, y: 500, w: 520, h: 150, size: 12, align: "left", maxLines: 11 },

      actTitle:  { x: 55, y: 295, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      actBody:   { x: 55, y: 265, w: 520, h: 80,  size: 12, align: "left", maxLines: 6 },
    },

    // p8: workWith
    p8: {
      hdrName: { x: 44, y: 788, w: 530, h: 20, size: 12, align: "left", maxLines: 1 },

      // Left column
      colL_title: { x: 60, y: 545, w: 240, h: 20, size: 12, align: "left", maxLines: 2 },
      colL_body:  { x: 60, y: 510, w: 240, h: 250, size: 11, align: "left", maxLines: 16, bulletIndent: 12, gap: 2 },

      // Right column
      colR_title: { x: 330, y: 545, w: 240, h: 20, size: 12, align: "left", maxLines: 2 },
      colR_body:  { x: 330, y: 510, w: 240, h: 250, size: 11, align: "left", maxLines: 16, bulletIndent: 12, gap: 2 },
    },

    // p9: actions
    p9: {
      hdrName: { x: 44, y: 788, w: 530, h: 20, size: 12, align: "left", maxLines: 1 },

      tldrTitle: { x: 55, y: 690, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      tldrBody:  { x: 55, y: 660, w: 520, h: 110, size: 12, align: "left", maxLines: 9 },

      execTitle: { x: 55, y: 530, w: 520, h: 18, size: 14, align: "left", maxLines: 2 },
      execBody:  { x: 55, y: 500, w: 520, h: 150, size: 12, align: "left", maxLines: 11 },
    }
  }
};

/* ───────────── layout override via URL ───────────── */

function applyLayoutOverridesFromUrl(layout, url) {
  // Supports:
  // &p3_tldrTitle_x=... &p3_tldrTitle_y=... etc
  // Keys are: p{N}_{box}_{prop}
  // props: x y w h size align maxLines gap bulletIndent
  const out = safeJson(layout);
  out.pages = out.pages || {};

  for (const [k, v] of url.searchParams.entries()) {
    if (!k.startsWith("p")) continue;
    const m = k.match(/^(p\d+)_(\w+)_(x|y|w|h|size|maxLines|gap|bulletIndent|align)$/);
    if (!m) continue;

    const pageKey = m[1];
    const boxKey = m[2];
    const prop = m[3];

    out.pages[pageKey] = out.pages[pageKey] || {};
    out.pages[pageKey][boxKey] = out.pages[pageKey][boxKey] || {};

    if (prop === "align") {
      out.pages[pageKey][boxKey][prop] = S(v);
    } else {
      out.pages[pageKey][boxKey][prop] = N(v);
    }
  }

  return out;
}

/* ───────────── template loading ───────────── */

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

async function loadFontBytesLocal(fontRelPath) {
  // Example: "fonts/OpenSans-Regular.ttf"
  const safe = String(fontRelPath || "").replace(/^\/+/, "");
  if (!safe) throw new Error("No font path provided");
  if (!/\.(ttf|otf)$/i.test(safe)) throw new Error(`Invalid font file: ${safe}`);

  const __file = fileURLToPath(import.meta.url);
  const __dir = path.dirname(__file);

  const candidates = [
    path.join(process.cwd(), "public", safe),
    path.join(__dir, "..", "public", safe),
    path.join(__dir, "public", safe),
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
    `Font not found: ${safe}. Tried: ${candidates.join(" | ")}. Last: ${lastErr?.message || "no detail"}`
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

/* ───────────── payload normalisation ───────────── */

function normalisePayload(P) {
  const payload = isObj(P) ? safeJson(P) : {};

  // Flatten common shapes:
  // payload.text.execSummary_tldr -> P["p3:tldr"]
  const t = payload.text || {};
  const ww = payload.workWith || {};

  const out = {};

  // Cover
  out["p1:n"] = norm(getDeep(payload, ["identity", "fullName"], ""));
  out["p1:e"] = norm(getDeep(payload, ["identity", "email"], ""));
  out["p1:d"] = norm(getDeep(payload, ["identity", "dateLabel"], getDeep(payload, ["dateLbl"], "")));

  // Page 3
  out["p3:tldr"] = asText(t.execSummary_tldr);
  out["p3:exec"] = asText(t.execSummary);
  out["p3:tip"]  = asText(t.execSummary_tipact);

  // Page 4
  out["p4:tldr"]     = asText(t.state_tldr);
  out["p4:stateDeep"] = asText(t.domState);
  out["p4:action"]    = asText(t.state_tipact);
  out["p4:bottom"]    = asText(t.bottomState);

  // Page 5
  out["p5:tldr"] = asText(t.frequency_tldr);
  out["p5:freq"] = asText(t.frequency);

  // Page 6
  out["p6:tldr"] = asText(t.sequence_tldr);
  out["p6:seq"]  = asText(t.sequence);
  out["p6:tip"]  = asText(t.sequence_tipact);

  // Page 7
  out["p7:tldr"] = asText(t.theme_tldr);
  out["p7:theme"] = asText(t.theme);
  out["p7:tip"] = asText(t.theme_tipact);

  // Page 8 workWith columns
  out["p8:colL_title"] = "Concealed";
  out["p8:colL_body"]  = asText(ww.concealed);

  out["p8:colR_title"] = "Triggered";
  out["p8:colR_body"]  = asText(ww.triggered);

  // Page 9
  out["p9:tldr"] = asText(t.act_anchor);
  out["p9:exec"] = asText(t.execSummary);
  out["p9:tip"]  = asText(t.execSummary_tipact);

  // Keep ctrl info
  out.ctrl = payload.ctrl || {};
  out.layout = payload.layout || {};

  return out;
}

/* ───────────── dom/2nd resolution ───────────── */

function computeDomAndSecondKeys(P) {
  // Prefer keys directly
  const domKey = norm(getDeep(P, ["ctrl", "dominantKey"], "")).toUpperCase();
  const secondKey = norm(getDeep(P, ["ctrl", "secondKey"], "")).toUpperCase();
  if (domKey && secondKey) return { domKey, secondKey };

  // Fallback: infer from labels (Concealed/Triggered/Regulated/Lead)
  const toKey = (label) => {
    const s = norm(label).toLowerCase();
    if (s.startsWith("con")) return "C";
    if (s.startsWith("tri")) return "T";
    if (s.startsWith("reg")) return "R";
    if (s.startsWith("lea")) return "L";
    return "";
  };

  const dom = toKey(getDeep(P, ["ctrl", "dominant"], ""));
  const snd = toKey(getDeep(P, ["ctrl", "secondState"], ""));
  return { domKey: dom || "C", secondKey: snd || "T" };
}

/* ───────────── main handler ───────────── */

export default async function handler(req, res) {
  try {
    const url = new URL(req.url, "http://localhost");

    // Optional debug probe
    const debug = url.searchParams.get("debug") === "1";

const payload = await readPayload(req);
const P = normalisePayload(payload);


    if (debug) {
      return res.status(200).json({
        ok: true,
        where: "fill-template:v2:after_normaliseInput",
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
          p6_act:  (P["p6:tip"] || "").length,

          p7_main: (P["p7:theme"] || "").length,
          p7_tldr: (P["p7:tldr"] || "").length,
          p7_act:  (P["p7:tip"] || "").length,

          p8_L: (P["p8:colL_body"] || "").length,
          p8_R: (P["p8:colR_body"] || "").length,

          p9_tldr: (P["p9:tldr"] || "").length
        }
      });
    }

    const { domKey, secondKey } = computeDomAndSecondKeys(P);
    const combo = `${domKey}${secondKey}`;

    const validCombos = new Set(["CT","CL","CR","TC","TR","TL","RC","RT","RL","LC","LR","LT"]);
    const safeCombo = validCombos.has(combo) ? combo : "CT";
    const tpl = `CTRL_PoC_Assessment_Profile_template_${safeCombo}.pdf`;

    const pdfBytes = await loadTemplateBytesLocal(tpl);
    const pdfDoc = await PDFDocument.load(pdfBytes);
    // Font: allow custom TTF/OTF via /public/fonts (URL override: &font=fonts/OpenSans-Regular.ttf)
    pdfDoc.registerFontkit(fontkit);
    const fontPath = url.searchParams.get("font") || "fonts/OpenSans-Regular.ttf";

    let font;
    try {
      const fontBytes = await loadFontBytesLocal(fontPath);
      font = await pdfDoc.embedFont(fontBytes, { subset: true });
    } catch (e) {
      console.warn("[fill-template] Custom font failed, falling back to Helvetica:", e?.message || String(e));
      font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    }

    // Layout: default + payload overrides + URL overrides (URL overrides already cover ALL boxes)
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

    // --- Page 1: cover ---
    if (pages[0] && L.p1) {
      if (L.p1.name)  drawTextBox(pages[0], font, P["p1:n"], L.p1.name,  { maxLines: L.p1.name.maxLines ?? 1 });
      if (L.p1.email) drawTextBox(pages[0], font, P["p1:e"], L.p1.email, { maxLines: L.p1.email.maxLines ?? 1 });
      if (L.p1.date)  drawTextBox(pages[0], font, P["p1:d"], L.p1.date,  { maxLines: L.p1.date.maxLines ?? 1 });
    }

    // --- Page 3: Exec Summary ---
    const p3 = pages[2];
    if (p3 && L.p3) {
      // TLDR (bullets)
      writeSection(
        p3, font,
        "TLDR",
        P["p3:tldr"],
        L.p3.tldrTitle,
        L.p3.tldrBody,
        "bullets",
        { maxLines: L.p3.tldrBody?.maxLines ?? 10 }
      );

      // Executive Summary (paragraph)
      writeSection(
        p3, font,
        "Executive Summary",
        P["p3:exec"],
        L.p3.execTitle,
        L.p3.execBody,
        "paragraph",
        { maxLines: L.p3.execBody?.maxLines ?? 12 }
      );

      // Key Action (short)
      writeSection(
        p3, font,
        "Key Action",
        P["p3:tip"],
        L.p3.actTitle,
        L.p3.actBody,
        "paragraph",
        { maxLines: L.p3.actBody?.maxLines ?? 6 }
      );
    }

    // --- Page 4: State Deep Dive ---
    const p4 = pages[3];
    if (p4 && L.p4) {
      writeSection(
        p4, font,
        "TLDR",
        P["p4:tldr"],
        L.p4.tldrTitle,
        L.p4.tldrBody,
        "bullets",
        { maxLines: L.p4.tldrBody?.maxLines ?? 9 }
      );

      // Executive Summary: include dom + bottom if present
      const domText = norm(P["p4:stateDeep"]);
      const bottomText = norm(P["p4:bottom"]);
      const p4Exec = [domText, bottomText].filter(Boolean).join("\n\n");
      writeSection(
        p4, font,
        "Executive Summary",
        p4Exec,
        L.p4.execTitle,
        L.p4.execBody,
        "paragraph",
        { maxLines: L.p4.execBody?.maxLines ?? 11 }
      );

      writeSection(
        p4, font,
        "Key Action",
        P["p4:action"],
        L.p4.actTitle,
        L.p4.actBody,
        "paragraph",
        { maxLines: L.p4.actBody?.maxLines ?? 6 }
      );
    }

    // --- Page 5: Frequency ---
    const p5 = pages[4];
    if (p5 && L.p5) {
      writeSection(
        p5, font,
        "TLDR",
        P["p5:tldr"],
        L.p5.tldrTitle,
        L.p5.tldrBody,
        "bullets",
        { maxLines: L.p5.tldrBody?.maxLines ?? 9 }
      );

      writeSection(
        p5, font,
        "Executive Summary",
        P["p5:freq"],
        L.p5.execTitle,
        L.p5.execBody,
        "paragraph",
        { maxLines: L.p5.execBody?.maxLines ?? 14 }
      );
    }

    // --- Page 6: Sequence ---
    const p6 = pages[5];
    if (p6 && L.p6) {
      writeSection(
        p6, font,
        "TLDR",
        P["p6:tldr"],
        L.p6.tldrTitle,
        L.p6.tldrBody,
        "bullets",
        { maxLines: L.p6.tldrBody?.maxLines ?? 9 }
      );

      writeSection(
        p6, font,
        "Executive Summary",
        P["p6:seq"],
        L.p6.execTitle,
        L.p6.execBody,
        "paragraph",
        { maxLines: L.p6.execBody?.maxLines ?? 11 }
      );

      writeSection(
        p6, font,
        "Key Action",
        P["p6:tip"],
        L.p6.actTitle,
        L.p6.actBody,
        "paragraph",
        { maxLines: L.p6.actBody?.maxLines ?? 6 }
      );
    }

    // --- Page 7: Themes ---
    const p7 = pages[6];
    if (p7 && L.p7) {
      writeSection(
        p7, font,
        "TLDR",
        P["p7:tldr"],
        L.p7.tldrTitle,
        L.p7.tldrBody,
        "bullets",
        { maxLines: L.p7.tldrBody?.maxLines ?? 9 }
      );

      writeSection(
        p7, font,
        "Executive Summary",
        P["p7:theme"],
        L.p7.execTitle,
        L.p7.execBody,
        "paragraph",
        { maxLines: L.p7.execBody?.maxLines ?? 11 }
      );

      writeSection(
        p7, font,
        "Key Action",
        P["p7:tip"],
        L.p7.actTitle,
        L.p7.actBody,
        "paragraph",
        { maxLines: L.p7.actBody?.maxLines ?? 6 }
      );
    }

    // --- Page 8: WorkWith ---
    const p8 = pages[7];
    if (p8 && L.p8) {
      if (L.p8.colL_title) drawTextBox(p8, font, P["p8:colL_title"], L.p8.colL_title, { maxLines: L.p8.colL_title.maxLines ?? 2 });
      if (L.p8.colL_body)  drawTextBox(p8, font, P["p8:colL_body"],  L.p8.colL_body,  { maxLines: L.p8.colL_body.maxLines ?? 16 });

      if (L.p8.colR_title) drawTextBox(p8, font, P["p8:colR_title"], L.p8.colR_title, { maxLines: L.p8.colR_title.maxLines ?? 2 });
      if (L.p8.colR_body)  drawTextBox(p8, font, P["p8:colR_body"],  L.p8.colR_body,  { maxLines: L.p8.colR_body.maxLines ?? 16 });
    }

    // --- Page 9: Actions ---
    const p9 = pages[8];
    if (p9 && L.p9) {
      writeSection(
        p9, font,
        "TLDR",
        P["p9:tldr"],
        L.p9.tldrTitle,
        L.p9.tldrBody,
        "bullets",
        { maxLines: L.p9.tldrBody?.maxLines ?? 9 }
      );

      writeSection(
        p9, font,
        "Executive Summary",
        P["p9:exec"],
        L.p9.execTitle,
        L.p9.execBody,
        "paragraph",
        { maxLines: L.p9.execBody?.maxLines ?? 11 }
      );
    }

    const outBytes = await pdfDoc.save();

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="CTRL_PoC_${safeCombo}.pdf"`);
    return res.status(200).send(Buffer.from(outBytes));
  } catch (err) {
    return res.status(500).json({
      ok: false,
      error: err?.message || String(err),
      stack: err?.stack || ""
    });
  }
}
