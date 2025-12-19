/**
 * CTRL PoC Export Service · fill-template (Starter/PoC flow)
 * Place at: /api/fill-template.js  (ctrl-poc-service)
 *
 * VERSION: V7
 * - Keeps V4 layout + rendering + master probe
 * - Adds V6.1-style payload aliasing + auto-heal for:
 *   ctrl.summary.dominantKey, ctrl.summary.secondKey, ctrl.summary.templateKey
 * - Normalises state keys (full words -> C/T/R/L)
 * - Derives templateKey from dom+second when missing
 * - Derives dom+second from templateKey when those are missing
 * - Validates templateKey against 12 combos; safe default CT
 */
export const config = { runtime: "nodejs" };

import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { PDFDocument, StandardFonts } from "pdf-lib";

// ────────────────────────────────────────────────────────────────
// Helpers: base64url decode, safe JSON parse, simple guards
// ────────────────────────────────────────────────────────────────
function b64urlToUtf8(b64url) {
  const b64 = String(b64url || "")
    .replace(/-/g, "+")
    .replace(/_/g, "/")
    .padEnd(Math.ceil(String(b64url || "").length / 4) * 4, "=");

  return Buffer.from(b64, "base64").toString("utf8");
}

function safeJsonParse(str) {
  try {
    return JSON.parse(str);
  } catch (e) {
    return null;
  }
}

function normStr(v) {
  if (v === null || v === undefined) return "";
  return String(v);
}

function pickFirst(obj, keys, fallback = "") {
  for (const k of keys) {
    const v = obj?.[k];
    if (v !== null && v !== undefined && String(v).trim().length) return v;
  }
  return fallback;
}

// Safe getter for nested paths like "ctrl.summary.dominantKey"
function getPath(obj, dottedPath) {
  if (!obj || !dottedPath) return undefined;
  const parts = String(dottedPath).split(".");
  let cur = obj;
  for (const p of parts) {
    if (!cur || typeof cur !== "object") return undefined;
    cur = cur[p];
  }
  return cur;
}

function pickFirstPath(obj, paths, fallback = "") {
  for (const p of paths) {
    const v = getPath(obj, p);
    if (v !== null && v !== undefined && String(v).trim().length) return v;
  }
  return fallback;
}

// ────────────────────────────────────────────────────────────────
// State key normalisation + templateKey validation
// ────────────────────────────────────────────────────────────────
const VALID_TEMPLATE_KEYS = new Set([
  "CT","CR","CL",
  "TC","TR","TL",
  "RC","RT","RL",
  "LC","LT","LR",
]);

function toStateLetter(v) {
  const s = normStr(v).trim();
  if (!s) return "";
  const up = s.toUpperCase();

  // Already a letter
  if (up === "C" || up === "T" || up === "R" || up === "L") return up;

  // Full words (user-facing)
  const map = {
    CONCEALED: "C",
    TRIGGERED: "T",
    REGULATED: "R",
    LEAD: "L",
  };
  if (map[up]) return map[up];

  // Sometimes people send "Concealed (Emerging)" etc.
  const cleaned = up.replace(/[^A-Z]/g, "");
  if (map[cleaned]) return map[cleaned];

  // Sometimes keys like "domKey":"Triggered"
  if (cleaned.includes("CONCEALED")) return "C";
  if (cleaned.includes("TRIGGERED")) return "T";
  if (cleaned.includes("REGULATED")) return "R";
  if (cleaned.includes("LEAD")) return "L";

  return "";
}

function normaliseTemplateKey(v) {
  const s = normStr(v).trim().toUpperCase();
  if (!s) return "";
  const cleaned = s.replace(/[^A-Z]/g, "");
  if (cleaned.length >= 2) {
    const tk = cleaned.slice(0, 2);
    if (VALID_TEMPLATE_KEYS.has(tk)) return tk;
  }
  return "";
}

function buildTemplateKey(domLetter, secondLetter) {
  const d = toStateLetter(domLetter);
  const s = toStateLetter(secondLetter);
  const tk = `${d}${s}`;
  return VALID_TEMPLATE_KEYS.has(tk) ? tk : "";
}

function deriveDomSecondFromTemplateKey(templateKey) {
  const tk = normaliseTemplateKey(templateKey);
  if (!tk) return { domKey: "", secondKey: "" };
  return { domKey: tk[0], secondKey: tk[1] };
}

// ────────────────────────────────────────────────────────────────
// Layout: coordinates (existing V4 system)
// ────────────────────────────────────────────────────────────────
const LAYOUT = {
  p3: {
    exec: { x: 55, y: 285, w: 950, size: 18, maxLines: 20, align: "left" },
    execTLDR: { x: 55, y: 215, w: 950, size: 18, maxLines: 10, align: "left" },
    tipAct: { x: 55, y: 730, w: 950, size: 18, maxLines: 10, align: "left" },
  },
  p4: {
    main: { x: 55, y: 250, w: 950, size: 18, maxLines: 28, align: "left" },
    tldr: { x: 55, y: 190, w: 950, size: 18, maxLines: 10, align: "left" },
    act: { x: 55, y: 735, w: 950, size: 18, maxLines: 10, align: "left" },
  },
  p5: {
    main: { x: 55, y: 250, w: 950, size: 18, maxLines: 22, align: "left" },
    tldr: { x: 55, y: 190, w: 950, size: 18, maxLines: 10, align: "left" },
  },
  p6: {
    main: { x: 55, y: 250, w: 950, size: 18, maxLines: 22, align: "left" },
    tldr: { x: 55, y: 190, w: 950, size: 18, maxLines: 10, align: "left" },
    act: { x: 55, y: 735, w: 950, size: 18, maxLines: 10, align: "left" },
  },
  p7: {
    top: { x: 55, y: 250, w: 950, size: 18, maxLines: 25, align: "left" },
    topTLDR: { x: 55, y: 190, w: 950, size: 18, maxLines: 10, align: "left" },
    low: { x: 55, y: 520, w: 950, size: 18, maxLines: 18, align: "left" },
    lowTLDR: { x: 55, y: 470, w: 950, size: 18, maxLines: 8, align: "left" },
    tip: { x: 55, y: 740, w: 950, size: 18, maxLines: 9, align: "left" },
  },
  p9: {
    anchor: { x: 55, y: 240, w: 950, size: 18, maxLines: 14, align: "left" },
  },
};

// ────────────────────────────────────────────────────────────────
// PDF text drawing (simple, stable)
// ────────────────────────────────────────────────────────────────
function splitToLines(text, maxCharsPerLine) {
  const t = normStr(text).replace(/\r/g, "");
  if (!t.trim()) return [];
  const words = t.split(/\s+/);
  const lines = [];
  let line = "";

  for (const w of words) {
    const cand = line ? `${line} ${w}` : w;
    if (cand.length <= maxCharsPerLine) {
      line = cand;
    } else {
      if (line) lines.push(line);
      line = w;
    }
  }
  if (line) lines.push(line);
  return lines;
}

function drawTextBox(page, font, text, box) {
  const { x, y, w, size, maxLines, align } = box;
  const t = normStr(text);
  if (!t.trim()) return;

  const maxChars = Math.max(18, Math.floor(w / (size * 0.55)));
  const lines = splitToLines(t, maxChars).slice(0, maxLines);

  const lineHeight = size * 1.25;
  let cursorY = y;

  for (const line of lines) {
    let tx = x;
    if (align === "center") {
      const tw = font.widthOfTextAtSize(line, size);
      tx = x + (w - tw) / 2;
    } else if (align === "right") {
      const tw = font.widthOfTextAtSize(line, size);
      tx = x + (w - tw);
    }

    page.drawText(line, { x: tx, y: cursorY, size, font });
    cursorY -= lineHeight;
  }
}

// ────────────────────────────────────────────────────────────────
// Normalise payload: V7 aliasing + self-heal for dom/second/template
// ────────────────────────────────────────────────────────────────
function normaliseInput(payload) {
  // Identity (same spirit as V4, but a bit more forgiving)
  const identity =
    payload?.identity ||
    payload?.ctrl?.summary?.identity ||
    payload?.ctrl?.identity ||
    payload?.summary?.identity ||
    {};

  const fullName = pickFirst(identity, ["fullName", "FullName", "name", "Name"], "");
  const email = pickFirst(identity, ["email", "Email"], "");
  const dateLabel = pickFirst(
    identity,
    ["dateLabel", "dateLbl", "date", "Date"],
    payload?.dateLbl || ""
  );

  // Pull ctrl-ish roots (multiple possible “homes”)
  const ctrl = payload?.ctrl || payload?.CTRL || {};
  const summary =
    ctrl?.summary ||
    payload?.ctrlSummary ||
    payload?.summary ||
    payload?.ctrl?.Summary ||
    {};

  // V7: read dom/second/template from MANY aliases (V6.1-style)
  // - accepts: ctrl.summary.*, summary.*, domSecond.*, top-level, older keys
  let dominantRaw = pickFirstPath(payload, [
    "ctrl.summary.dominantKey",
    "ctrl.summary.domKey",
    "ctrl.summary.dominant",
    "ctrl.summary.dominantState",
    "ctrl.summary.domState",
    "ctrl.dominantKey",
    "ctrl.domKey",
    "ctrl.dominant",
    "ctrl.dominantState",
    "ctrlSummary.dominantKey",
    "ctrlSummary.domKey",
    "summary.dominantKey",
    "summary.domKey",
    "domSecond.domKey",
    "domSecond.dominantKey",
    "dominantKey",
    "domKey",
    "dominant",
    "dominantState",
    "domState",
  ], "");

  let secondRaw = pickFirstPath(payload, [
    "ctrl.summary.secondKey",
    "ctrl.summary.secondState",
    "ctrl.summary.second",
    "ctrl.secondKey",
    "ctrl.secondState",
    "ctrlSummary.secondKey",
    "ctrlSummary.secondState",
    "summary.secondKey",
    "summary.secondState",
    "domSecond.secondKey",
    "domSecond.secondState",
    "secondKey",
    "secondState",
    "second",
  ], "");

  let templateRaw = pickFirstPath(payload, [
    "ctrl.summary.templateKey",
    "ctrl.summary.tplKey",
    "ctrl.summary.template",
    "ctrl.templateKey",
    "ctrl.tplKey",
    "ctrlSummary.templateKey",
    "summary.templateKey",
    "domSecond.templateKey",
    "templateKey",
    "tplKey",
    "template",
  ], "");

  // Normalise to letters
  let dominantKey = toStateLetter(dominantRaw);
  let secondKey = toStateLetter(secondRaw);
  let templateKey = normaliseTemplateKey(templateRaw);

  // Self-heal logic:
  // 1) If templateKey exists but dom/second missing -> derive dom/second
  if (templateKey && (!dominantKey || !secondKey)) {
    const d = deriveDomSecondFromTemplateKey(templateKey);
    dominantKey = dominantKey || d.domKey;
    secondKey = secondKey || d.secondKey;
  }

  // 2) If dom+second exists but templateKey missing -> build templateKey
  if (!templateKey && dominantKey && secondKey) {
    templateKey = buildTemplateKey(dominantKey, secondKey);
  }

  // 3) Final safety: if templateKey still invalid -> default CT
  if (!templateKey) templateKey = "CT";

  // 12-band scoring
  const bands =
    ctrl?.bands ||
    payload?.bands ||
    payload?.ctrl12 ||
    payload?.ctrl?.ctrl12 ||
    payload?.ctrl?.bands ||
    {};

  // Text sections (Gen output)
  const text = payload?.text || payload?.gen || payload?.copy || {};

  // Work-with lens (accept more aliases)
  const workWith =
    payload?.workWith ||
    payload?.workwith ||
    payload?.work_with ||
    payload?.collab ||
    payload?.lens ||
    {};

  // Questions meta
  const questions =
    payload?.questions ||
    payload?.ctrl?.questions ||
    payload?.ctrl?.summary?.questions ||
    payload?.summary?.questions ||
    [];

  return {
    raw: payload,
    identity: { fullName, email, dateLabel },
    ctrl: {
      summary: {
        dominantKey: dominantKey || null,
        secondKey: secondKey || null,
        templateKey: templateKey || null,
        dominant: summary?.dominant || null,
        secondState: summary?.secondState || summary?.second || null,
        dominantSubState: summary?.dominantSubState || null,
      },
      bands,
    },
    text,
    workWith,
    questions,
  };
}

// ────────────────────────────────────────────────────────────────
// Choose template PDF file by templateKey
// ────────────────────────────────────────────────────────────────
function pickTemplateInfo(P) {
  let tk = normaliseTemplateKey(P?.ctrl?.summary?.templateKey);
  if (!tk) tk = "CT";
  const filename = `CTRL_PoC_Assessment_Profile_template_${tk}.pdf`;
  return { templateKey: tk, filename };
}

// ────────────────────────────────────────────────────────────────
// MAIN HANDLER (unchanged V4 structure)
// ────────────────────────────────────────────────────────────────
export default async function handler(req, res) {
  try {
    const __filename = fileURLToPath(import.meta.url);
    const __dirname = path.dirname(__filename);

    const url = new URL(req.url, "https://example.local");
    const dataParam = url.searchParams.get("data");
    if (!dataParam) return res.status(400).json({ ok: false, error: "Missing ?data=" });

    const jsonStr = b64urlToUtf8(dataParam);
    const payload = safeJsonParse(jsonStr);
    if (!payload) return res.status(400).json({ ok: false, error: "Invalid JSON in ?data=" });

    const P = normaliseInput(payload);
    const info = pickTemplateInfo(P);

    const pdfPath = path.join(__dirname, "..", "public", info.filename);
    const pdfBytes = await fs.readFile(pdfPath);

    const pdfDoc = await PDFDocument.load(pdfBytes);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const pages = pdfDoc.getPages();

    const reportDate = P.identity.dateLabel || payload?.dateLbl || "";

    // ────────────────────────────────────────────────────────────────
    // MASTER DEBUG (V7 keeps V4 probe style)
    // debug=1 returns summary JSON
    // debug=2 returns full values JSON
    // ────────────────────────────────────────────────────────────────
    const debug = String(url.searchParams.get("debug") || "").trim();
    const wantDebug = debug === "1" || debug === "2";

    const trunc = (v, n = 140) => {
      const s = (v === null || v === undefined) ? "" : String(v);
      if (s.length <= n) return s;
      return s.slice(0, n) + "…";
    };

    const strInfo = (v) => {
      if (v === null || v === undefined) return { has: false, len: 0, preview: "" };
      const s = String(v);
      return { has: s.length > 0, len: s.length, preview: trunc(s) };
    };

    const expectedTextKeys = [
      "execSummary_tldr",
      "execSummary",
      "execSummary_tipact",
      "state_tldr",
      "domState",
      "bottomState",
      "state_tipact",
      "frequency_tldr",
      "frequency",
      "sequence_tldr",
      "sequence",
      "sequence_tipact",
      "theme_tldr",
      "theme",
      "theme_tipact",
      "act_anchor",
    ];

    const expectedWorkWithKeys = ["concealed", "triggered", "regulated", "lead"];

    const buildMasterProbe = () => {
      const text = P.text || {};
      const workWith = P.workWith || {};
      const bands = (P.ctrl && P.ctrl.bands) || (P.bands || null);

      const textInfo = {};
      expectedTextKeys.forEach((k) => (textInfo[k] = strInfo(text[k])));

      const freqTldrExtras = Object.keys(text)
        .filter((k) => k.toLowerCase().startsWith("frequency_tldr") && k !== "frequency_tldr")
        .sort();

      const workWithInfo = {};
      expectedWorkWithKeys.forEach((k) => (workWithInfo[k] = strInfo(workWith[k])));

      const bandsKeys = bands ? Object.keys(bands) : [];
      const bandsPresent12 =
        ["C_low","C_mid","C_high","T_low","T_mid","T_high","R_low","R_mid","R_high","L_low","L_mid","L_high"]
          .filter((k) => bands && typeof bands[k] !== "undefined").length;

      const missing = { identity: [], ctrl: [], text: [], workWith: [] };

      if (!P.identity?.fullName) missing.identity.push("identity.fullName");
      if (!P.identity?.email) missing.identity.push("identity.email");
      if (!P.identity?.dateLabel) missing.identity.push("identity.dateLabel");

      if (!P.ctrl?.summary?.dominantKey) missing.ctrl.push("ctrl.summary.dominantKey");
      if (!P.ctrl?.summary?.secondKey) missing.ctrl.push("ctrl.summary.secondKey");
      if (!P.ctrl?.summary?.templateKey) missing.ctrl.push("ctrl.summary.templateKey");
      if (!bands || bandsPresent12 !== 12) missing.ctrl.push("ctrl.bands (12/12)");

      expectedTextKeys.forEach((k) => {
        if (!textInfo[k].has && k !== "bottomState") missing.text.push(`text.${k}`);
      });

      expectedWorkWithKeys.forEach((k) => {
        if (!workWithInfo[k].has) missing.workWith.push(`workWith.${k}`);
      });

      const lengths = {
        p3_exec: (text.execSummary || "").length,
        p3_tldr: (text.execSummary_tldr || "").length,
        p3_tip: (text.execSummary_tipact || "").length,
        p4_main: (text.domState || "").length + (text.bottomState || "").length,
        p4_tldr: (text.state_tldr || "").length,
        p4_act: (text.state_tipact || "").length,
        p5_main: (text.frequency || "").length,
        p5_tldr: (text.frequency_tldr || "").length,
        p6_main: (text.sequence || "").length,
        p6_tldr: (text.sequence_tldr || "").length,
        p6_act: (text.sequence_tipact || "").length,
        p7_top: (text.theme || "").length,
        p7_top_tldr: (text.theme_tldr || "").length,
        p7_low: 0,
        p7_low_tldr: 0,
        p7_top_tip: (text.theme_tipact || "").length,
        p9_anchor: (text.act_anchor || "").length,
        bandsKeys: bandsKeys.length,
        bandsPresent12,
        freqTldrExtraKeys: freqTldrExtras.length,
      };

      const summary = {
        ok: true,
        where: "fill-template:v7:master_probe:summary",
        domSecond: {
          domKey: P.ctrl?.summary?.dominantKey || null,
          secondKey: P.ctrl?.summary?.secondKey || null,
          templateKey: P.ctrl?.summary?.templateKey || null,
        },
        identity: {
          fullName: strInfo(P.identity?.fullName),
          email: strInfo(P.identity?.email),
          dateLabel: strInfo(P.identity?.dateLabel),
        },
        counts: {
          questions: Array.isArray(P.questions) ? P.questions.length : 0,
          bandsKeys: bandsKeys.length,
          bandsPresent12,
          textKeys: Object.keys(text).length,
          workWithKeys: Object.keys(workWith).length,
          freqTldrExtraKeys: freqTldrExtras,
        },
        lengths,
        missing,
        previews: {
          execSummary_tldr: trunc(text.execSummary_tldr),
          execSummary: trunc(text.execSummary),
          domState: trunc(text.domState),
          frequency_tldr: trunc(text.frequency_tldr),
          frequency: trunc(text.frequency),
          sequence_tldr: trunc(text.sequence_tldr),
          theme_tldr: trunc(text.theme_tldr),
          act_anchor: trunc(text.act_anchor),
          workWith_triggered: trunc(workWith.triggered),
        },
      };

      const full = {
        ok: true,
        where: "fill-template:v7:master_probe:full",
        identity: P.identity || null,
        ctrlSummary: P.ctrl?.summary || null,
        bands: bands || null,
        questions: P.questions || null,
        text: text || null,
        workWith: workWith || null,
        missing,
        lengths,
      };

      return { summary, full };
    };

    const MASTER = buildMasterProbe();

    try {
      console.log("[fill-template] MASTER_PROBE", JSON.stringify(MASTER.summary));
    } catch (_) {
      console.log("[fill-template] MASTER_PROBE (stringify failed)");
    }

    if (wantDebug) {
      return res.status(200).json(debug === "2" ? MASTER.full : MASTER.summary);
    }

    // Page 3: Exec Summary
    const p3 = pages[2];
    if (p3 && LAYOUT.p3) {
      drawTextBox(p3, font, P.text.execSummary_tldr, LAYOUT.p3.execTLDR);
      drawTextBox(p3, font, P.text.execSummary, LAYOUT.p3.exec);
      drawTextBox(p3, font, P.text.execSummary_tipact, LAYOUT.p3.tipAct);
    }

    // Page 4: Dominant / State narrative
    const p4 = pages[3];
    if (p4 && LAYOUT.p4) {
      drawTextBox(p4, font, P.text.state_tldr, LAYOUT.p4.tldr);
      const domAndBottom = `${normStr(P.text.domState)}${P.text.bottomState ? "\n\n" + normStr(P.text.bottomState) : ""}`;
      drawTextBox(p4, font, domAndBottom, LAYOUT.p4.main);
      drawTextBox(p4, font, P.text.state_tipact, LAYOUT.p4.act);
    }

    // Page 5: Frequency
    const p5 = pages[4];
    if (p5 && LAYOUT.p5) {
      drawTextBox(p5, font, P.text.frequency_tldr, LAYOUT.p5.tldr);
      drawTextBox(p5, font, P.text.frequency, LAYOUT.p5.main);
    }

    // Page 6: Sequence
    const p6 = pages[5];
    if (p6 && LAYOUT.p6) {
      drawTextBox(p6, font, P.text.sequence_tldr, LAYOUT.p6.tldr);
      drawTextBox(p6, font, P.text.sequence, LAYOUT.p6.main);
      drawTextBox(p6, font, P.text.sequence_tipact, LAYOUT.p6.act);
    }

    // Page 7: Themes
    const p7 = pages[6];
    if (p7 && LAYOUT.p7) {
      drawTextBox(p7, font, P.text.theme_tldr, LAYOUT.p7.topTLDR);
      drawTextBox(p7, font, P.text.theme, LAYOUT.p7.top);
      drawTextBox(p7, font, P.text.theme_tipact, LAYOUT.p7.tip);
    }

    // Page 9: Action Anchor
    const p9 = pages[8];
    if (p9 && LAYOUT.p9) {
      drawTextBox(p9, font, P.text.act_anchor, LAYOUT.p9.anchor);
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
