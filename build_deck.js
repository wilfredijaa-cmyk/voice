// Generates cascaded-voice-ai.pptx — run with: node build_deck.js
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const md = require("react-icons/md");

// Palette: deep plum (dominant), amber "signal" accent, soft lavender surfaces
const C = {
  plum: "24163A",
  plumMid: "3A2757",
  plumSoft: "5B4580",
  amber: "FFB22E",
  coral: "F2545B",
  teal: "2BB3A3",
  lav: "F4F1F9",
  lavMid: "E4DDF0",
  ink: "1E1A28",
  muted: "6B6480",
  white: "FFFFFF",
};
const HEAD = "Cambria";
const BODY = "Calibri";

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const shadow = () => ({ type: "outer", color: "000000", opacity: 0.18, blur: 6, offset: 2, angle: 90 });

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.title = "Cascaded Voice AI";

  // ---------- Slide 1: Architecture ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.plum };
    s.addText("VOICE AI ARCHITECTURES", {
      x: 0.5, y: 0.35, w: 9, h: 0.3, fontFace: BODY, fontSize: 12, bold: true,
      color: C.amber, charSpacing: 4, margin: 0, isTextBox: true,
    });
    s.addText("The Cascaded Pipeline", {
      x: 0.5, y: 0.65, w: 9, h: 0.7, fontFace: HEAD, fontSize: 38, bold: true,
      color: C.white, margin: 0, isTextBox: true,
    });
    s.addText("Three specialist models, chained through text — the “Unix philosophy” of voice AI.", {
      x: 0.5, y: 1.35, w: 9, h: 0.4, fontFace: BODY, fontSize: 15, italic: true,
      color: C.lavMid, margin: 0, isTextBox: true,
    });

    // Pipeline: pill, model, pill, model, pill, model, pill
    const pillW = 0.75, modelW = 1.4, gap = 0.3, y = 2.35, modelH = 1.35;
    const nodes = [
      { kind: "pill", label: "Audio", ic: fa.FaMicrophone },
      { kind: "model", stage: "STT", name: "Whisper", ic: md.MdHearing, color: C.teal },
      { kind: "pill", label: "Text", ic: md.MdTextFields },
      { kind: "model", stage: "LLM", name: "Llama 3", ic: fa.FaBrain, color: C.amber },
      { kind: "pill", label: "Text", ic: md.MdTextFields },
      { kind: "model", stage: "TTS", name: "VoxCPM", ic: md.MdRecordVoiceOver, color: C.coral },
      { kind: "pill", label: "Audio", ic: fa.FaVolumeUp },
    ];
    let x = 0.5;
    for (let i = 0; i < nodes.length; i++) {
      const n = nodes[i];
      if (n.kind === "pill") {
        const ph = 0.95, py = y + (modelH - ph) / 2;
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
          x, y: py, w: pillW, h: ph, rectRadius: 0.15,
          fill: { color: C.plumMid }, line: { color: C.plumSoft, width: 1 },
        });
        s.addImage({ data: await icon(n.ic, C.lavMid), x: x + (pillW - 0.34) / 2, y: py + 0.14, w: 0.34, h: 0.34 });
        s.addText(n.label, {
          x, y: py + 0.55, w: pillW, h: 0.3, fontFace: BODY, fontSize: 12,
          color: C.lavMid, align: "center", margin: 0, isTextBox: true,
        });
        x += pillW;
      } else {
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
          x, y, w: modelW, h: modelH, rectRadius: 0.12,
          fill: { color: C.white }, line: { color: n.color, width: 2.5 }, shadow: shadow(),
        });
        s.addShape(pres.shapes.OVAL, {
          x: x + (modelW - 0.52) / 2, y: y + 0.13, w: 0.52, h: 0.52, fill: { color: n.color }, line: { color: n.color },
        });
        s.addImage({ data: await icon(n.ic, C.white), x: x + (modelW - 0.3) / 2, y: y + 0.24, w: 0.3, h: 0.3 });
        s.addText(n.stage, {
          x, y: y + 0.7, w: modelW, h: 0.3, fontFace: BODY, fontSize: 15, bold: true,
          color: C.ink, align: "center", margin: 0, isTextBox: true,
        });
        s.addText(n.name, {
          x, y: y + 0.98, w: modelW, h: 0.26, fontFace: BODY, fontSize: 12,
          color: C.muted, align: "center", margin: 0, isTextBox: true,
        });
        x += modelW;
      }
      if (i < nodes.length - 1) {
        s.addShape(pres.shapes.LINE, {
          x: x + 0.04, y: y + modelH / 2, w: gap - 0.08, h: 0,
          line: { color: C.amber, width: 2, endArrowType: "triangle" },
        });
        x += gap;
      }
    }

    // Stage captions under each model
    const caps = ["Speech → transcript", "Reasoning · RAG · tools", "Transcript → speech"];
    const modelXs = [0.5 + pillW + gap, 0.5 + 2 * (pillW + gap) + modelW + gap, 0.5 + 3 * (pillW + gap) + 2 * (modelW + gap)];
    modelXs.forEach((mx, i) => {
      s.addText(caps[i], {
        x: mx - 0.2, y: y + modelH + 0.12, w: modelW + 0.4, h: 0.3, fontFace: BODY, fontSize: 11,
        color: C.lavMid, align: "center", margin: 0, isTextBox: true,
      });
    });

    s.addText([
      { text: "Key idea: ", options: { bold: true, color: C.amber } },
      { text: "text is the contract between every stage — so each model can be tuned, swapped, inspected and scaled on its own.", options: { color: C.white } },
    ], { x: 0.5, y: 4.65, w: 9, h: 0.5, fontFace: BODY, fontSize: 14, margin: 0, isTextBox: true });

    s.addNotes("The cascaded architecture chains three specialist models. Whisper transcribes the user's speech, Llama 3 reasons over the transcript and produces a text reply, and VoxCPM synthesises that reply back into speech. Because the hand-off between stages is plain text, each component can be optimised or replaced independently.");
  }

  // ---------- Slide 2: Strengths (2x2 grid) ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    s.addText("Why Cascaded Works", {
      x: 0.5, y: 0.4, w: 9, h: 0.65, fontFace: HEAD, fontSize: 34, bold: true, color: C.plum, margin: 0, isTextBox: true,
    });
    s.addText("Text in the middle buys control, clarity and flexibility.", {
      x: 0.5, y: 1.02, w: 9, h: 0.35, fontFace: BODY, fontSize: 15, color: C.muted, margin: 0, isTextBox: true,
    });
    const items = [
      { ic: fa.FaBrain, t: "Best-in-class reasoning & RAG", d: "Text LLMs lead on logic, JSON output and RAG — 100k tokens of manuals is easy in text, painful in an audio KV cache." },
      { ic: fa.FaExchangeAlt, t: "Modular & swappable", d: "Need a new voice? Swap only the TTS model. STT and LLM stay untouched; each stage upgrades on its own schedule." },
      { ic: fa.FaShieldAlt, t: "Easy guardrails", d: "Run regex, NeMo Guardrails or Llama Guard on the intermediate text and block unsafe output before it ever reaches TTS." },
      { ic: fa.FaSearch, t: "Full observability", d: "An exact transcript of what the user said and what the AI replied makes debugging, QA and auditing straightforward." },
    ];
    const cw = 4.35, ch = 1.7, gx = 0.3, gy = 0.3, x0 = 0.5, y0 = 1.62;
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      const cx = x0 + (i % 2) * (cw + gx), cy = y0 + Math.floor(i / 2) * (ch + gy);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: cx, y: cy, w: cw, h: ch, rectRadius: 0.1, fill: { color: C.lav }, line: { color: C.lavMid, width: 1 },
      });
      s.addShape(pres.shapes.OVAL, { x: cx + 0.25, y: cy + 0.25, w: 0.55, h: 0.55, fill: { color: C.plum }, line: { color: C.plum } });
      s.addImage({ data: await icon(it.ic, C.amber), x: cx + 0.39, y: cy + 0.39, w: 0.27, h: 0.27 });
      s.addText(it.t, {
        x: cx + 0.95, y: cy + 0.25, w: cw - 1.15, h: 0.55, fontFace: BODY, fontSize: 16, bold: true,
        color: C.plum, valign: "middle", margin: 0, isTextBox: true,
      });
      s.addText(it.d, {
        x: cx + 0.25, y: cy + 0.9, w: cw - 0.5, h: 0.7, fontFace: BODY, fontSize: 12,
        color: C.ink, valign: "top", margin: 0, isTextBox: true,
      });
    }
    s.addNotes("The core advantage is that text sits between every stage. That unlocks mature text-LLM reasoning and RAG, independent component swaps, simple text-level moderation, and a complete transcript for debugging.");
  }

  // ---------- Slide 3: Weaknesses + latency stack ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    s.addText("The Trade-offs", {
      x: 0.5, y: 0.4, w: 9, h: 0.65, fontFace: HEAD, fontSize: 34, bold: true, color: C.plum, margin: 0, isTextBox: true,
    });
    s.addText("What gets lost — and what gets expensive — when speech becomes text.", {
      x: 0.5, y: 1.02, w: 9, h: 0.35, fontFace: BODY, fontSize: 15, color: C.muted, margin: 0, isTextBox: true,
    });

    const rows = [
      { ic: md.MdSentimentDissatisfied, t: "Paralinguistic blind spot", d: "Text strips tone. A heavy sigh plus “Yeah, sure, whatever” is transcribed as agreement — and answered cheerfully." },
      { ic: fa.FaStopwatch, t: "The latency tax", d: "Time-to-first-audio is the sum of every stage. Sub-800 ms takes heroic engineering." },
      { ic: fa.FaHandPaper, t: "Barge-in complexity", d: "Interruptions need a separate VAD, instant TTS cancellation, buffer flushes and race-free state handling." },
    ];
    const ry0 = 1.62, rh = 1.1;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i], ry = ry0 + i * rh;
      s.addShape(pres.shapes.OVAL, { x: 0.5, y: ry + 0.05, w: 0.55, h: 0.55, fill: { color: C.coral }, line: { color: C.coral } });
      s.addImage({ data: await icon(r.ic, C.white), x: 0.64, y: ry + 0.19, w: 0.27, h: 0.27 });
      s.addText(r.t, {
        x: 1.25, y: ry, w: 3.9, h: 0.32, fontFace: BODY, fontSize: 15, bold: true, color: C.plum, margin: 0, isTextBox: true,
      });
      s.addText(r.d, {
        x: 1.25, y: ry + 0.33, w: 3.9, h: 0.66, fontFace: BODY, fontSize: 12, color: C.ink, valign: "top", margin: 0, isTextBox: true,
      });
    }

    // Latency budget panel
    const px = 5.5, py = 1.55, pw = 4.0, ph = 3.55;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: px, y: py, w: pw, h: ph, rectRadius: 0.1, fill: { color: C.plum }, line: { color: C.plum },
    });
    s.addText("Where time-to-first-audio goes", {
      x: px + 0.25, y: py + 0.2, w: pw - 0.5, h: 0.32, fontFace: BODY, fontSize: 14, bold: true, color: C.white, margin: 0, isTextBox: true,
    });
    s.addText("Illustrative budget for an ~800 ms target", {
      x: px + 0.25, y: py + 0.5, w: pw - 0.5, h: 0.26, fontFace: BODY, fontSize: 10, italic: true, color: C.lavMid, margin: 0, isTextBox: true,
    });
    const steps = [
      ["VAD tail-padding", 200, C.lavMid],
      ["STT inference", 150, C.teal],
      ["LLM prefill + 1st token", 250, C.amber],
      ["TTS prefill + 1st chunk", 200, C.coral],
    ];
    const barX = px + 0.25, barMaxW = pw - 0.5, total = 800;
    let by = py + 0.95;
    for (const [label, ms, col] of steps) {
      s.addText([
        { text: label, options: { color: C.white } },
        { text: `  ${ms} ms`, options: { color: col, bold: true } },
      ], { x: barX, y: by, w: barMaxW, h: 0.26, fontFace: BODY, fontSize: 11, margin: 0, isTextBox: true });
      s.addShape(pres.shapes.RECTANGLE, { x: barX, y: by + 0.3, w: barMaxW, h: 0.16, fill: { color: C.plumMid }, line: { color: C.plumMid } });
      s.addShape(pres.shapes.RECTANGLE, { x: barX, y: by + 0.3, w: barMaxW * ms / total, h: 0.16, fill: { color: col }, line: { color: col } });
      by += 0.6;
    }
    s.addNotes("Converting speech to text discards tone and emotion. Latency is additive across VAD, STT, LLM and TTS — the budget shown is illustrative, not measured. Handling interruptions requires an external VAD and a careful cancellation state machine.");
  }

  // ---------- Slide 4: Comparison table ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    s.addText("Cascaded vs. Omni (Native Audio)", {
      x: 0.5, y: 0.35, w: 9, h: 0.6, fontFace: HEAD, fontSize: 30, bold: true, color: C.plum, margin: 0, isTextBox: true,
    });
    const hdr = (t, fill, color) => ({ text: t, options: { bold: true, fill: { color: fill }, color, fontSize: 12 } });
    const good = (t) => ({ text: t, options: { color: "1E7A6F", bold: true } });
    const bad = (t) => ({ text: t, options: { color: "B8323A" } });
    const plain = (t) => ({ text: t, options: { color: C.ink } });
    const feat = (t) => ({ text: t, options: { bold: true, color: C.plum } });
    const data = [
      [hdr("Feature", C.plum, C.white), hdr("Cascaded (STT → LLM → TTS)", C.plum, C.amber), hdr("Omni (audio ↔ audio)", C.plum, C.white)],
      [feat("Latency (TTFA)"), bad("~800 ms – 1.5 s; heavy optimisation"), good("~300 – 500 ms; near human")],
      [feat("Emotion & tone"), bad("Lost in transcription"), good("Preserved and mirrored")],
      [feat("Barge-in"), bad("External VAD + state machine"), good("Native — model stops generating")],
      [feat("Reasoning / RAG"), good("Excellent — mature text LLMs"), bad("Poor / unreliable")],
      [feat("Guardrails"), good("Easy — filter text before TTS"), bad("Hard — must filter audio")],
      [feat("Hallucinations"), good("Textual, easy to catch"), bad("Acoustic artefacts, odd voices")],
      [feat("Voice changes"), good("Swap the TTS model"), bad("Baked into weights; fine-tune")],
      [feat("Serving"), plain("Multi-model routing"), plain("Huge audio KV-cache")],
      [feat("Cost / minute"), good("Moderate; scale stages independently"), bad("High; large unified GPU memory")],
    ];
    s.addTable(data, {
      x: 0.5, y: 1.1, w: 9, colW: [2.0, 3.5, 3.5], rowH: 0.38,
      fontFace: BODY, fontSize: 11, valign: "middle", margin: [0, 0.1, 0, 0.1],
      border: { type: "solid", pt: 0.75, color: C.lavMid },
      fill: { color: C.white },
    });
    s.addText("Green = advantage · Red = limitation", {
      x: 0.5, y: 5.05, w: 9, h: 0.25, fontFace: BODY, fontSize: 10, italic: true, color: C.muted, margin: 0, isTextBox: true,
    });
    s.addNotes("Omni models process audio tokens directly, winning on latency, emotion and interruptions. Cascaded pipelines win on reasoning, RAG, guardrails, voice flexibility and cost. Latency figures are typical ranges, not benchmarks.");
  }

  // ---------- Slide 5: When to choose ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.plum };
    s.addText("Choosing an Architecture", {
      x: 0.5, y: 0.45, w: 9, h: 0.65, fontFace: HEAD, fontSize: 34, bold: true, color: C.white, margin: 0, isTextBox: true,
    });
    const cols = [
      {
        title: "Cascaded fits when…", color: C.amber, ic: fa.FaProjectDiagram,
        pts: ["Answers depend on documents, tools or strict JSON", "Compliance needs auditable transcripts and text guardrails", "Brand voices must be swappable", "Cost and independent scaling matter"],
      },
      {
        title: "Omni fits when…", color: C.teal, ic: fa.FaWaveSquare,
        pts: ["Natural, sub-500 ms conversation is the product", "Tone and emotion must be heard and mirrored", "Users interrupt often", "Reasoning depth is modest"],
      },
    ];
    for (let i = 0; i < 2; i++) {
      const c = cols[i], cx = 0.5 + i * 4.65, cy = 1.35, cw = 4.35, ch = 2.85;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: cx, y: cy, w: cw, h: ch, rectRadius: 0.1, fill: { color: C.plumMid }, line: { color: C.plumSoft, width: 1 },
      });
      s.addShape(pres.shapes.OVAL, { x: cx + 0.25, y: cy + 0.25, w: 0.5, h: 0.5, fill: { color: c.color }, line: { color: c.color } });
      s.addImage({ data: await icon(c.ic, C.plum), x: cx + 0.37, y: cy + 0.37, w: 0.26, h: 0.26 });
      s.addText(c.title, {
        x: cx + 0.9, y: cy + 0.25, w: cw - 1.1, h: 0.5, fontFace: BODY, fontSize: 17, bold: true, color: c.color, valign: "middle", margin: 0, isTextBox: true,
      });
      s.addText(c.pts.map((p, j) => ({ text: p, options: { bullet: true, breakLine: j < c.pts.length - 1 } })), {
        x: cx + 0.25, y: cy + 0.95, w: cw - 0.5, h: ch - 1.15, fontFace: BODY, fontSize: 13, color: C.white,
        paraSpaceAfter: 6, valign: "top", margin: 0, isTextBox: true,
      });
    }
    s.addText("Cascaded trades latency and emotional nuance for control, accuracy and flexibility.", {
      x: 0.5, y: 4.7, w: 9, h: 0.4, fontFace: HEAD, fontSize: 16, italic: true, color: C.lavMid, margin: 0, isTextBox: true,
    });
    s.addNotes("Use cascaded pipelines for knowledge-heavy, regulated or brand-sensitive use cases. Use omni models where conversational feel is the product. Hybrid designs — omni front-end with text-LLM tool calls — are an emerging middle ground.");
  }

  await pres.writeFile({ fileName: "cascaded-voice-ai.pptx" });
  console.log("wrote cascaded-voice-ai.pptx");
}

main().catch((e) => { console.error(e); process.exit(1); });
