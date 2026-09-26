// Generates cascaded-voice-ai.pptx — run with: node build_deck.js
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const md = require("react-icons/md");

// Palette: deep forest green (dominant), amber "signal" accent, pale mint surfaces
const C = {
  green: "0E3B2E",
  greenMid: "185443",
  greenSoft: "2F7A60",
  amber: "FFB22E",
  coral: "F2545B",
  sky: "3FA7D6",
  mint: "E6F2EA",
  mintMid: "CFE6D6",
  ink: "15241E",
  muted: "4F6559",
  white: "FFFFFF",
  good: "1E7A4F",
  bad: "B8323A",
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

const shadow = () => ({ type: "outer", color: "000000", opacity: 0.2, blur: 6, offset: 2, angle: 90 });

function title(s, text, sub, dark = false) {
  s.addText(text, {
    x: 0.5, y: 0.4, w: 9, h: 0.65, fontFace: HEAD, fontSize: 32, bold: true,
    color: dark ? C.white : C.green, margin: 0, isTextBox: true,
  });
  if (sub) {
    s.addText(sub, {
      x: 0.5, y: 1.02, w: 9, h: 0.35, fontFace: BODY, fontSize: 15,
      color: dark ? C.mintMid : C.muted, margin: 0, isTextBox: true,
    });
  }
}

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.title = "Cascaded Voice AI";

  // ---------- Slide 1: Architecture ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.green };
    s.addText("VOICE AI ARCHITECTURES", {
      x: 0.5, y: 0.35, w: 9, h: 0.3, fontFace: BODY, fontSize: 12, bold: true,
      color: C.amber, charSpacing: 4, margin: 0, isTextBox: true,
    });
    s.addText("The Cascaded Pipeline", {
      x: 0.5, y: 0.65, w: 9, h: 0.7, fontFace: HEAD, fontSize: 36, bold: true,
      color: C.white, margin: 0, isTextBox: true,
    });
    s.addText("Three specialist models, chained through text — the “Unix philosophy” of voice AI.", {
      x: 0.5, y: 1.33, w: 9, h: 0.4, fontFace: BODY, fontSize: 15, italic: true,
      color: C.mintMid, margin: 0, isTextBox: true,
    });

    // Pipeline: pill, model, pill, model, pill, model, pill
    const pillW = 0.75, modelW = 1.4, gap = 0.3, y = 1.95, modelH = 2.05;
    const nodes = [
      { kind: "pill", label: "Audio", ic: fa.FaMicrophone },
      { kind: "model", stage: "STT", tools: ["Whisper", "NVIDIA Parakeet", "Moonshine"], ic: md.MdHearing, color: C.sky },
      { kind: "pill", label: "Text", ic: md.MdTextFields },
      { kind: "model", stage: "LLM", tools: ["Llama 3", "Qwen3", "Gemma 3"], ic: fa.FaBrain, color: C.amber },
      { kind: "pill", label: "Text", ic: md.MdTextFields },
      { kind: "model", stage: "TTS", tools: ["VoxCPM", "Kokoro", "Chatterbox"], ic: md.MdRecordVoiceOver, color: C.coral },
      { kind: "pill", label: "Audio", ic: fa.FaVolumeUp },
    ];
    let x = 0.5;
    for (let i = 0; i < nodes.length; i++) {
      const n = nodes[i];
      if (n.kind === "pill") {
        const ph = 0.95, py = y + (modelH - ph) / 2;
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
          x, y: py, w: pillW, h: ph, rectRadius: 0.15,
          fill: { color: C.greenMid }, line: { color: C.greenSoft, width: 1 },
        });
        s.addImage({ data: await icon(n.ic, C.mintMid), x: x + (pillW - 0.34) / 2, y: py + 0.14, w: 0.34, h: 0.34 });
        s.addText(n.label, {
          x, y: py + 0.55, w: pillW, h: 0.3, fontFace: BODY, fontSize: 12,
          color: C.mintMid, align: "center", margin: 0, isTextBox: true,
        });
        x += pillW;
      } else {
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
          x, y, w: modelW, h: modelH, rectRadius: 0.12,
          fill: { color: C.white }, line: { color: n.color, width: 2.5 }, shadow: shadow(),
        });
        s.addShape(pres.shapes.OVAL, {
          x: x + (modelW - 0.5) / 2, y: y + 0.12, w: 0.5, h: 0.5, fill: { color: n.color }, line: { color: n.color },
        });
        s.addImage({ data: await icon(n.ic, C.white), x: x + (modelW - 0.28) / 2, y: y + 0.23, w: 0.28, h: 0.28 });
        s.addText(n.stage, {
          x, y: y + 0.68, w: modelW, h: 0.3, fontFace: BODY, fontSize: 15, bold: true,
          color: C.ink, align: "center", margin: 0, isTextBox: true,
        });
        s.addShape(pres.shapes.LINE, {
          x: x + 0.3, y: y + 1.04, w: modelW - 0.6, h: 0, line: { color: C.mintMid, width: 1 },
        });
        s.addText(n.tools.map((t, j) => ({
          text: t,
          options: { bold: j === 0, color: j === 0 ? C.ink : C.muted, breakLine: j < n.tools.length - 1 },
        })), {
          x, y: y + 1.12, w: modelW, h: 0.82, fontFace: BODY, fontSize: 11,
          align: "center", valign: "top", paraSpaceAfter: 2, margin: 0, isTextBox: true,
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
    s.addText("Bold = reference stack · others are open-source drop-in alternatives", {
      x: 0.5, y: y + modelH + 0.12, w: 9, h: 0.26, fontFace: BODY, fontSize: 10, italic: true,
      color: C.mintMid, align: "center", margin: 0, isTextBox: true,
    });

    s.addText([
      { text: "Key idea: ", options: { bold: true, color: C.amber } },
      { text: "text is the contract between every stage — so each model can be tuned, swapped, inspected and scaled on its own.", options: { color: C.white } },
    ], { x: 0.5, y: 4.7, w: 9, h: 0.5, fontFace: BODY, fontSize: 14, margin: 0, isTextBox: true });

    s.addNotes("The cascaded architecture chains three specialist models. Speech-to-text (Whisper, or open alternatives such as NVIDIA Parakeet and Moonshine) transcribes the user; a text LLM (Llama 3, Qwen3, Gemma 3) reasons and replies; text-to-speech (VoxCPM, Kokoro, Chatterbox) speaks the answer. Because the hand-off between stages is plain text, each component can be optimised or replaced independently.");
  }

  // ---------- Slide 2: Real-time plumbing ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.mint };
    title(s, "Real-Time Plumbing", "What turns three models into a conversation you can actually talk to.");

    // Flow row
    const chips = [
      ["User device", 1.35, C.white, C.green],
      ["Transport", 1.35, C.sky, C.white],
      ["VAD", 1.35, C.amber, C.ink],
      ["Turn detection", 1.6, C.coral, C.white],
      ["STT → LLM → TTS", 1.85, C.green, C.white],
    ];
    const cgap = (9 - chips.reduce((a, c) => a + c[1], 0)) / (chips.length - 1);
    let cx = 0.5;
    const cy = 1.55;
    chips.forEach(([label, w, fill, color], i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: cx, y: cy, w, h: 0.45, rectRadius: 0.22, fill: { color: fill }, line: { color: fill === C.white ? C.mintMid : fill, width: 1 },
      });
      s.addText(label, {
        x: cx, y: cy, w, h: 0.45, fontFace: BODY, fontSize: 12, bold: true, color, align: "center", valign: "middle", margin: 0, isTextBox: true,
      });
      cx += w;
      if (i < chips.length - 1) {
        s.addShape(pres.shapes.LINE, {
          x: cx + 0.05, y: cy + 0.225, w: cgap - 0.1, h: 0,
          line: { color: C.greenSoft, width: 1.5, endArrowType: "triangle", beginArrowType: i === 0 ? "triangle" : undefined },
        });
        cx += cgap;
      }
    });

    const cards = [
      { ic: fa.FaBroadcastTower, color: C.sky, t: "Transport",
        d: "WebRTC for browser & mobile: UDP, Opus, jitter buffer, echo cancellation. WebSocket for servers; SIP for phones.",
        tools: ["LiveKit", "mediasoup", "Janus Gateway"] },
      { ic: fa.FaWaveSquare, color: C.amber, t: "VAD",
        d: "Labels each ~30 ms frame as speech or silence. Gates STT and fires barge-in when the user talks over the bot.",
        tools: ["Silero VAD", "WebRTC VAD", "TEN VAD"] },
      { ic: fa.FaExchangeAlt, color: C.coral, t: "Turn Detection",
        d: "Decides whether the user has finished or is only pausing. Semantic models beat fixed silence timeouts.",
        tools: ["LiveKit Turn Detector", "Pipecat Smart Turn", "TEN Turn Detection"] },
      { ic: fa.FaProjectDiagram, color: C.green, t: "Orchestration",
        d: "Streams audio and text frames between stages; handles interruptions, TTS cancellation and conversation context.",
        tools: ["Pipecat", "LiveKit Agents", "TEN Framework"] },
    ];
    const cw = 2.1, cg = 0.2, top = 2.2, ch = 2.95;
    for (let i = 0; i < cards.length; i++) {
      const c = cards[i], x = 0.5 + i * (cw + cg);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: top, w: cw, h: ch, rectRadius: 0.1, fill: { color: C.white }, line: { color: C.mintMid, width: 1 },
      });
      s.addShape(pres.shapes.OVAL, { x: x + 0.18, y: top + 0.18, w: 0.42, h: 0.42, fill: { color: c.color }, line: { color: c.color } });
      s.addImage({ data: await icon(c.ic, c.color === C.amber ? C.ink : C.white), x: x + 0.28, y: top + 0.28, w: 0.22, h: 0.22 });
      s.addText(c.t, {
        x: x + 0.18, y: top + 0.7, w: cw - 0.36, h: 0.3, fontFace: BODY, fontSize: 14, bold: true,
        color: C.green, valign: "middle", margin: 0, isTextBox: true,
      });
      s.addText(c.d, {
        x: x + 0.18, y: top + 1.05, w: cw - 0.36, h: 1.0, fontFace: BODY, fontSize: 10,
        color: C.ink, valign: "top", margin: 0, isTextBox: true,
      });
      s.addText("OPEN-SOURCE TOOLS", {
        x: x + 0.18, y: top + 2.1, w: cw - 0.36, h: 0.2, fontFace: BODY, fontSize: 8, bold: true,
        color: C.muted, charSpacing: 1, margin: 0, isTextBox: true,
      });
      s.addText(c.tools.map((t, j) => ({ text: t, options: { breakLine: j < c.tools.length - 1 } })), {
        x: x + 0.18, y: top + 2.3, w: cw - 0.36, h: 0.55, fontFace: BODY, fontSize: 10, bold: true,
        color: C.greenSoft, valign: "top", margin: 0, isTextBox: true,
      });
    }
    s.addNotes("The models are only half the system. WebRTC is the default transport for low-latency browser and mobile audio (UDP, Opus, jitter buffering, echo cancellation); WebSockets suit server-to-server links and SIP connects phone calls. VAD such as Silero classifies short frames as speech or silence and triggers barge-in. Turn-detection models (LiveKit's turn detector, Pipecat Smart Turn, TEN Turn Detection) judge whether the user is actually done, avoiding cut-offs from fixed silence timers. Frameworks like Pipecat, LiveKit Agents and TEN Framework wire it all together and manage interruptions and context.");
  }

  // ---------- Slide 3: Strengths (2x2 grid) ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.mint };
    title(s, "Why Cascaded Works", "Text in the middle buys control, clarity and flexibility.");
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
        x: cx, y: cy, w: cw, h: ch, rectRadius: 0.1, fill: { color: C.white }, line: { color: C.mintMid, width: 1 },
      });
      s.addShape(pres.shapes.OVAL, { x: cx + 0.25, y: cy + 0.25, w: 0.55, h: 0.55, fill: { color: C.green }, line: { color: C.green } });
      s.addImage({ data: await icon(it.ic, C.amber), x: cx + 0.39, y: cy + 0.39, w: 0.27, h: 0.27 });
      s.addText(it.t, {
        x: cx + 0.95, y: cy + 0.25, w: cw - 1.15, h: 0.55, fontFace: BODY, fontSize: 16, bold: true,
        color: C.green, valign: "middle", margin: 0, isTextBox: true,
      });
      s.addText(it.d, {
        x: cx + 0.25, y: cy + 0.9, w: cw - 0.5, h: 0.7, fontFace: BODY, fontSize: 12,
        color: C.ink, valign: "top", margin: 0, isTextBox: true,
      });
    }
    s.addNotes("The core advantage is that text sits between every stage. That unlocks mature text-LLM reasoning and RAG, independent component swaps, simple text-level moderation, and a complete transcript for debugging.");
  }

  // ---------- Slide 4: Weaknesses + latency stack ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.mint };
    title(s, "The Trade-offs", "What gets lost — and what gets expensive — when speech becomes text.");

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
        x: 1.25, y: ry, w: 3.9, h: 0.32, fontFace: BODY, fontSize: 15, bold: true, color: C.green, margin: 0, isTextBox: true,
      });
      s.addText(r.d, {
        x: 1.25, y: ry + 0.33, w: 3.9, h: 0.66, fontFace: BODY, fontSize: 12, color: C.ink, valign: "top", margin: 0, isTextBox: true,
      });
    }

    // Latency budget panel
    const px = 5.5, py = 1.55, pw = 4.0, ph = 3.55;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: px, y: py, w: pw, h: ph, rectRadius: 0.1, fill: { color: C.green }, line: { color: C.green },
    });
    s.addText("Where time-to-first-audio goes", {
      x: px + 0.25, y: py + 0.2, w: pw - 0.5, h: 0.32, fontFace: BODY, fontSize: 14, bold: true, color: C.white, margin: 0, isTextBox: true,
    });
    s.addText("Illustrative budget for an ~800 ms target", {
      x: px + 0.25, y: py + 0.5, w: pw - 0.5, h: 0.26, fontFace: BODY, fontSize: 10, italic: true, color: C.mintMid, margin: 0, isTextBox: true,
    });
    const steps = [
      ["VAD tail-padding", 200, C.mintMid],
      ["STT inference", 150, C.sky],
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
      s.addShape(pres.shapes.RECTANGLE, { x: barX, y: by + 0.3, w: barMaxW, h: 0.16, fill: { color: C.greenMid }, line: { color: C.greenMid } });
      s.addShape(pres.shapes.RECTANGLE, { x: barX, y: by + 0.3, w: barMaxW * ms / total, h: 0.16, fill: { color: col }, line: { color: col } });
      by += 0.6;
    }
    s.addNotes("Converting speech to text discards tone and emotion. Latency is additive across VAD, STT, LLM and TTS — the budget shown is illustrative, not measured. Handling interruptions requires an external VAD and a careful cancellation state machine.");
  }

  // ---------- Slide 5: Comparison table ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.mint };
    s.addText("Cascaded vs. Omni (Native Audio)", {
      x: 0.5, y: 0.35, w: 9, h: 0.6, fontFace: HEAD, fontSize: 30, bold: true, color: C.green, margin: 0, isTextBox: true,
    });
    const hdr = (t, color) => ({ text: t, options: { bold: true, fill: { color: C.green }, color, fontSize: 12 } });
    const good = (t) => ({ text: t, options: { color: C.good, bold: true } });
    const bad = (t) => ({ text: t, options: { color: C.bad } });
    const plain = (t) => ({ text: t, options: { color: C.ink } });
    const feat = (t) => ({ text: t, options: { bold: true, color: C.green } });
    const data = [
      [hdr("Feature", C.white), hdr("Cascaded (STT → LLM → TTS)", C.amber), hdr("Omni (audio ↔ audio)", C.white)],
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
      border: { type: "solid", pt: 0.75, color: C.mintMid },
      fill: { color: C.white },
    });
    s.addText("Green = advantage · Red = limitation", {
      x: 0.5, y: 5.05, w: 9, h: 0.25, fontFace: BODY, fontSize: 10, italic: true, color: C.muted, margin: 0, isTextBox: true,
    });
    s.addNotes("Omni models process audio tokens directly, winning on latency, emotion and interruptions. Cascaded pipelines win on reasoning, RAG, guardrails, voice flexibility and cost. Latency figures are typical ranges, not benchmarks.");
  }

  // ---------- Slide 6: When to choose + omni examples ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.green };
    title(s, "Choosing an Architecture", "Cascaded trades latency and emotional nuance for control, accuracy and flexibility.", true);
    const cols = [
      {
        title: "Cascaded fits when…", color: C.amber, ic: fa.FaProjectDiagram,
        pts: ["Answers depend on documents, tools or strict JSON", "Compliance needs transcripts and text guardrails", "Brand voices must be swappable"],
      },
      {
        title: "Omni fits when…", color: C.sky, ic: fa.FaWaveSquare,
        pts: ["Natural, sub-500 ms conversation is the product", "Tone and emotion must be heard and mirrored", "Users interrupt often; reasoning is modest"],
      },
    ];
    for (let i = 0; i < 2; i++) {
      const c = cols[i], cx = 0.5 + i * 4.65, cy = 1.55, cw = 4.35, ch = 1.95;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: cx, y: cy, w: cw, h: ch, rectRadius: 0.1, fill: { color: C.greenMid }, line: { color: C.greenSoft, width: 1 },
      });
      s.addShape(pres.shapes.OVAL, { x: cx + 0.22, y: cy + 0.18, w: 0.44, h: 0.44, fill: { color: c.color }, line: { color: c.color } });
      s.addImage({ data: await icon(c.ic, C.green), x: cx + 0.33, y: cy + 0.29, w: 0.22, h: 0.22 });
      s.addText(c.title, {
        x: cx + 0.8, y: cy + 0.18, w: cw - 1.0, h: 0.44, fontFace: BODY, fontSize: 16, bold: true, color: c.color, valign: "middle", margin: 0, isTextBox: true,
      });
      s.addText(c.pts.map((p, j) => ({ text: p, options: { bullet: true, breakLine: j < c.pts.length - 1 } })), {
        x: cx + 0.22, y: cy + 0.75, w: cw - 0.44, h: ch - 0.9, fontFace: BODY, fontSize: 12, color: C.white,
        paraSpaceAfter: 4, valign: "top", margin: 0, isTextBox: true,
      });
    }

    s.addText("EXAMPLE OMNI / SPEECH-TO-SPEECH MODELS", {
      x: 0.5, y: 3.75, w: 9, h: 0.25, fontFace: BODY, fontSize: 11, bold: true, color: C.amber, charSpacing: 2, margin: 0, isTextBox: true,
    });
    const models = [
      ["Qwen3-Omni", "Alibaba · open"],
      ["Moshi", "Kyutai · open"],
      ["MiniCPM-o", "OpenBMB · open"],
      ["GLM-4-Voice", "Zhipu AI · open"],
      ["GPT Realtime", "OpenAI · API"],
      ["Gemini Live", "Google · API"],
    ];
    const mw = 1.4, mg = (9 - 6 * mw) / 5;
    models.forEach(([name, org], i) => {
      const mx = 0.5 + i * (mw + mg), my = 4.1;
      const open = org.endsWith("open");
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: mx, y: my, w: mw, h: 0.85, rectRadius: 0.1,
        fill: { color: open ? C.white : C.greenMid }, line: { color: open ? C.white : C.greenSoft, width: 1 },
      });
      s.addText(name, {
        x: mx, y: my + 0.14, w: mw, h: 0.32, fontFace: BODY, fontSize: 13, bold: true,
        color: open ? C.green : C.white, align: "center", margin: 0, isTextBox: true,
      });
      s.addText(org, {
        x: mx, y: my + 0.47, w: mw, h: 0.26, fontFace: BODY, fontSize: 10,
        color: open ? C.muted : C.mintMid, align: "center", margin: 0, isTextBox: true,
      });
    });
    s.addNotes("Use cascaded pipelines for knowledge-heavy, regulated or brand-sensitive use cases; use omni models where conversational feel is the product. Open-weight speech-to-speech examples: Qwen3-Omni (Alibaba), Moshi (Kyutai, full-duplex), MiniCPM-o (OpenBMB), GLM-4-Voice (Zhipu AI). Hosted APIs: OpenAI's GPT Realtime and Google's Gemini Live native-audio models. Hybrid designs — an omni front-end that calls a text LLM for tools and RAG — are an emerging middle ground.");
  }

  await pres.writeFile({ fileName: "cascaded-voice-ai.pptx" });
  console.log("wrote cascaded-voice-ai.pptx");
}

main().catch((e) => { console.error(e); process.exit(1); });
