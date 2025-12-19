/**
 * CTRL PoC Export Service · fill-template (Starter/PoC flow)
 * Place at: /api/fill-template.js  (ctrl-poc-service)
 *
 * VERSION: V4 (adds split Frequency page blocks for TLDR/main/action)
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
  // Convert base64url → base64
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

function clamp(n, lo, hi) {
  const x = Number(n);
  if (Number.isNaN(x)) return lo;
  return Math.max(lo, Math.min(hi, x));
}

function pickFirst(obj, keys, fallback = "") {
  for (const k of keys) {
    const v = obj?.[k];
    if (v !== null && v !== undefined && String(v).trim().length) return v;
  }
  return fallback;
}

// ────────────────────────────────────────────────────────────────
// Layout: coordinates (existing system)
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

  // crude but predictable: estimate characters per line from width and font size
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

    page.drawText(line, {
      x: tx,
      y: cursorY,
      size,
      font,
    });
    cursorY -= lineHeight;
  }
}

// ────────────────────────────────────────────────────────────────
// Normalise payload: accept your current Botpress/Vercel schema
// ────────────────────────────────────────────────────────────────
function normaliseInput(payload) {
  // Expect: { identity, ctrl: { summary, bands }, text, workWith, questions }
  const identity = payload?.identity || payload?.ctrl?.summary?.identity || payload?.ctrl?.summary?.identity || {};

  // allow alt casing
  const fullName = pickFirst(identity, ["fullName", "FullName", "name", "Name"], "");
  const email = pickFirst(identity, ["email", "Email"], "");
  const dateLabel = pickFirst(identity, ["dateLabel", "dateLbl", "date", "Date"], payload?.dateLbl || "");

  // ctrl summary bits (keys + labels)
  const ctrl = payload?.ctrl || {};
  const summary = ctrl?.summary || {};
  const dominantKey = pickFirst(summary, ["dominantKey", "dominant", "dominant_state", "domKey"], "");
  const secondKey = pickFirst(summary, ["secondKey", "secondState", "second_state", "second"], "");
  const templateKey = pickFirst(summary, ["templateKey", "tplKey", "template"], "");

  // 12-band scoring
  const bands = ctrl?.bands || payload?.bands || payload?.ctrl12 || payload?.ctrl?.ctrl12 || {};

  // text sections (Gen output)
  const text = payload?.text || payload?.gen || payload?.copy || {};

  // Work-with lens
  const workWith = payload?.workWith || payload?.workwith || payload?.collab || payload?.lens || {};

  // Questions meta
  const questions = payload?.questions || payload?.ctrl?.questions || payload?.ctrl?.summary?.questions || [];

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
// Choose template PDF file by templateKey (TL/RT/etc)
// (Keep your existing mapping; extend if needed)
// ────────────────────────────────────────────────────────────────
function pickTemplateInfo(P) {
  // Your repo uses /public/*.pdf
  // Example key: "TL" → CTRL_PoC_Assessment_Profile_template_TL.pdf
  const templateKey = P?.ctrl?.summary?.templateKey || "CT";
  const filename = `CTRL_PoC_Assessment_Profile_template_${templateKey}.pdf`;
  return { templateKey, filename };
}

// ────────────────────────────────────────────────────────────────
// MAIN HANDLER
// ────────────────────────────────────────────────────────────────
export default async function handler(req, res) {
  try {
    const __filename = fileURLToPath(import.meta.url);
    const __dirname = path.dirname(__filename);

    // Parse URL safely (Next/Vercel gives req.url relative)
    const url = new URL(req.url, "https://example.local");

    // Payload can arrive via ?data=base64
    const dataParam = url.searchParams.get("data");
    if (!dataParam) {
      return res.status(400).json({ ok: false, error: "Missing ?data=" });
    }

    const jsonStr = b64urlToUtf8(dataParam);
    const payload = safeJsonParse(jsonStr);
    if (!payload) {
      return res.status(400).json({ ok: false, error: "Invalid JSON in ?data=" });
    }

    const P = normaliseInput(payload);
    const info = pickTemplateInfo(P);

    // Resolve PDF template path
    const pdfPath = path.join(__dirname, "..", "public", info.filename);
    const pdfBytes = await fs.readFile(pdfPath);

    const pdfDoc = await PDFDocument.load(pdfBytes);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const pages = pdfDoc.getPages();

    // Derive the report date label
    const reportDate = P.identity.dateLabel || payload?.dateLbl || "";

    // ────────────────────────────────────────────────────────────────
    // MASTER DEBUG (V4)
    // - Always logs summary to Vercel logs
    // - debug=1 returns summary JSON
    // - debug=2 returns full values JSON
    // ────────────────────────────────────────────────────────────────
    const debug = String(url.searchParams.get("debug") || "").trim(); // "" | "1" (summary) | "2" (full)
    const wantDebug = debug === "1" || debug === "2";

    // ────────────────────────────────────────────────────────────────
    // Master probe builder (keeps you sane when payloads go feral)
    // ────────────────────────────────────────────────────────────────
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
      "frequency_tldr",        // p5 TLDR (and any extra freq TLDRs will be listed separately)
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

      // Capture any additional p5/frequency TLDR variants, if you add them later (e.g., frequency_tldr2)
      const freqTldrExtras = Object.keys(text)
        .filter((k) => k.toLowerCase().startsWith("frequency_tldr") && k !== "frequency_tldr")
        .sort();

      const workWithInfo = {};
      expectedWorkWithKeys.forEach((k) => (workWithInfo[k] = strInfo(workWith[k])));

      const bandsKeys = bands ? Object.keys(bands) : [];
      const bandsPresent12 =
        ["C_low","C_mid","C_high","T_low","T_mid","T_high","R_low","R_mid","R_high","L_low","L_mid","L_high"]
          .filter((k) => bands && typeof bands[k] !== "undefined").length;

      const missing = {
        identity: [],
        ctrl: [],
        text: [],
        workWith: [],
      };

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
        // page-aligned lengths (keeps parity with your debug snippets)
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
        where: "fill-template:v4:master_probe:summary",
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
        // high signal previews (so you can eyeball quickly)
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
        where: "fill-template:v4:master_probe:full",
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

    // Always log a compact summary (never the full text blobs) to Vercel logs.
    try {
      console.log("[fill-template] MASTER_PROBE", JSON.stringify(MASTER.summary));
    } catch (_) {
      console.log("[fill-template] MASTER_PROBE (stringify failed)");
    }

    // Optional debug response (so you can open the URL and inspect quickly)
    if (wantDebug) {
      return res.status(200).json(debug === "2" ? MASTER.full : MASTER.summary);
    }

    // ────────────────────────────────────────────────────────────────
    // Page 1 overlays (name/date/email)
    // (Your template likely already has these printed; keep minimal)
    // ────────────────────────────────────────────────────────────────
    // If you do have placeholders, add them here.
    // (Leaving as-is: your earlier versions were template-driven.)

    // ────────────────────────────────────────────────────────────────
    // Page 3: Exec Summary
    // ────────────────────────────────────────────────────────────────
    const p3 = pages[2];
    if (p3 && LAYOUT.p3) {
      drawTextBox(p3, font, P.text.execSummary_tldr, LAYOUT.p3.execTLDR);
      drawTextBox(p3, font, P.text.execSummary, LAYOUT.p3.exec);
      drawTextBox(p3, font, P.text.execSummary_tipact, LAYOUT.p3.tipAct);
    }

    // ────────────────────────────────────────────────────────────────
    // Page 4: Dominant / State narrative
    // ────────────────────────────────────────────────────────────────
    const p4 = pages[3];
    if (p4 && LAYOUT.p4) {
      drawTextBox(p4, font, P.text.state_tldr, LAYOUT.p4.tldr);
      const domAndBottom = `${normStr(P.text.domState)}${P.text.bottomState ? "\n\n" + normStr(P.text.bottomState) : ""}`;
      drawTextBox(p4, font, domAndBottom, LAYOUT.p4.main);
      drawTextBox(p4, font, P.text.state_tipact, LAYOUT.p4.act);
    }

    // ────────────────────────────────────────────────────────────────
    // Page 5: Frequency
    // ────────────────────────────────────────────────────────────────
    const p5 = pages[4];
    if (p5 && LAYOUT.p5) {
      drawTextBox(p5, font, P.text.frequency_tldr, LAYOUT.p5.tldr);
      drawTextBox(p5, font, P.text.frequency, LAYOUT.p5.main);
    }

    // ────────────────────────────────────────────────────────────────
    // Page 6: Sequence
    // ────────────────────────────────────────────────────────────────
    const p6 = pages[5];
    if (p6 && LAYOUT.p6) {
      drawTextBox(p6, font, P.text.sequence_tldr, LAYOUT.p6.tldr);
      drawTextBox(p6, font, P.text.sequence, LAYOUT.p6.main);
      drawTextBox(p6, font, P.text.sequence_tipact, LAYOUT.p6.act);
    }

    // ────────────────────────────────────────────────────────────────
    // Page 7: Themes
    // ────────────────────────────────────────────────────────────────
    const p7 = pages[6];
    if (p7 && LAYOUT.p7) {
      drawTextBox(p7, font, P.text.theme_tldr, LAYOUT.p7.topTLDR);
      drawTextBox(p7, font, P.text.theme, LAYOUT.p7.top);
      // optional low-theme blocks (kept zero-length in your current payload)
      drawTextBox(p7, font, P.text.theme_tipact, LAYOUT.p7.tip);
    }

    // ────────────────────────────────────────────────────────────────
    // Page 9: Action Anchor
    // ────────────────────────────────────────────────────────────────
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
