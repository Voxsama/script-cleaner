// cleaner.js - browser port of hindi_extract.py. Same rules, same output.
// Works in the browser and in Node (for tests). No libraries.
(function (root) {
  "use strict";

  const SENTENCES_PER_BLOCK = 2;
  const BLANK_LINE_BETWEEN_PARAGRAPHS = true;
  const BREAK_TAG = '<break time="1s" />';

  const DEV_WORD = "[\\u0900-\\u097F]+";
  const ARROW = "(?:\\bto\\b|->|=>|→)";
  const DEVANAGARI = /[ऀ-ॿ]/;
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  const PAIR_RE = () => new RegExp(`(${DEV_WORD})\\s*${ARROW}\\s*(${DEV_WORD})`, "gu");
  const VERSE_EXAMPLE_RE = new RegExp(
    `(${DEV_WORD})\\s+(\\d+)\\s*:\\s*(\\d+)\\s*[-–—]\\s*(\\d+)\\s*${ARROW}\\s*(.{0,80})`, "u");
  const PAUSE_RE = () => /^[ \t]*pause\s+(\d+)\s*(?:s|sec|secs|seconds?)?\s+(after|before)\s*:?\s*(.+?)\s*$/gimu;

  // ---------------- 1. Understand the prompt ----------------
  class Rules {
    constructor() { this.swaps = []; this.verseRange = null; this.verseSingle = null; this.pauses = []; }

    static fromText(text) {
      const r = new Rules();
      const pauseLines = [...text.matchAll(PAUSE_RE())].map((m) => [m[1], m[2], m[3]]);
      text = text.replace(PAUSE_RE(), "");
      const m = VERSE_EXAMPLE_RE.exec(text);
      if (m) {
        const [, book, ch, v1, v2, rhs] = m;
        const nums = [[ch, "{ch}"], [v1, "{v1}"], [v2, "{v2}"]];
        let pos = 0, end = 0;
        for (const [num] of nums) {
          const i = rhs.indexOf(num, pos);
          if (i < 0) break;
          pos = end = i + num.length;
        }
        if (end) {
          let template = rhs.slice(0, end);
          for (const [num, ph] of [...nums].reverse()) {
            template = template.replace(new RegExp(`(?<!\\d)${num}(?!\\d)`), ph);
          }
          template = template.replace(book, "{book}");
          r.verseRange = template;
          r.verseSingle = template.slice(0, template.indexOf("{v1}") + 4);
        }
        const rhsStart = m.index + m[0].length - rhs.length;
        text = text.slice(0, m.index) + text.slice(rhsStart + end);
      }
      for (const [, oldW, newW] of text.matchAll(PAIR_RE())) {
        if (oldW !== newW && !r.swaps.some(([a, b]) => a === oldW && b === newW)) r.swaps.push([oldW, newW]);
      }
      for (const [secs, where, phrase] of pauseLines) {
        r.pauses.push([parseInt(secs, 10), where.toLowerCase(), r.apply(phrase.trim())]);
      }
      return r;
    }

    describe() {
      const out = [];
      const fill = (t, o) => t.replace(/\{(\w+)\}/g, (_, k) => o[k]);
      if (this.verseRange) {
        out.push("Scripture refs: <book> 6:28-40 -> " + fill(this.verseRange, { book: "<book>", ch: 6, v1: 28, v2: 40 }));
        out.push("                <book> 6:28    -> " + fill(this.verseSingle, { book: "<book>", ch: 6, v1: 28 }));
      }
      for (const [a, b] of this.swaps) out.push(`${a} -> ${b}`);
      for (const [s, w, p] of this.pauses) out.push(`${s}s pause ${w}: ${p}`);
      return out.join("\n");
    }

    apply(text) {
      const fill = (t, o) => t.replace(/\{(\w+)\}/g, (_, k) => o[k]);
      if (this.verseRange) {
        text = text.replace(new RegExp(`(${DEV_WORD})\\s+(\\d+)\\s*:\\s*(\\d+)\\s*[-–—]\\s*(\\d+)`, "gu"),
          (_, book, ch, v1, v2) => fill(this.verseRange, { book, ch, v1, v2 }));
        text = text.replace(new RegExp(`(${DEV_WORD})\\s+(\\d+)\\s*:\\s*(\\d+)(?![\\d:])`, "gu"),
          (_, book, ch, v1) => fill(this.verseSingle, { book, ch, v1 }));
      }
      for (const [a, b] of this.swaps) text = text.split(a).join(b);
      return text;
    }

    addPauses(sentence) {
      for (const [secs, where, phrase] of this.pauses) {
        if (sentence.includes(phrase)) {
          const tags = Array(secs).fill(BREAK_TAG).join(" ");
          sentence = where === "before" ? `${tags} ${sentence}` : `${sentence} ${tags}`;
        }
      }
      return sentence;
    }
  }

  // ---------------- 2. Read the .docx ----------------
  async function inflateRaw(bytes) {
    const ds = new DecompressionStream("deflate-raw");
    const stream = new Blob([bytes]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function readZipEntry(buf, wanted) {
    const u8 = new Uint8Array(buf), dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("Not a Word (.docx) file");
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    for (let n = 0; n < count; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), elen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      const name = new TextDecoder().decode(u8.subarray(p + 46, p + 46 + nlen));
      if (name === wanted) {
        const lnlen = dv.getUint16(off + 26, true), lelen = dv.getUint16(off + 28, true);
        const data = u8.subarray(off + 30 + lnlen + lelen, off + 30 + lnlen + lelen + csize);
        const raw = method === 0 ? data : await inflateRaw(data);
        return new TextDecoder("utf-8").decode(raw);
      }
      p += 46 + nlen + elen + clen;
    }
    throw new Error("Not a Word (.docx) file (no document inside)");
  }

  // Walk the XML in document order. Uses DOMParser in the browser; a tiny
  // tag scanner otherwise (Node tests), which is enough for document.xml.
  function paragraphTexts(xml) {
    const decode = (s) => s.replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-f]+);/gi, (_, e) =>
      ({ lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" }[e.toLowerCase()] ??
        String.fromCodePoint(e[1].toLowerCase() === "x" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10))));
    const paras = [];
    // stack of open paragraph buffers (paragraphs can nest via text boxes)
    const stack = [];
    const re = /<(\/?)w:(p|t|tab|br|cr)\b[^>]*?(\/?)>|([^<]+)|<[^>]*>/g;
    let inT = 0, m;
    while ((m = re.exec(xml))) {
      const [, close, tag, selfClose, text] = m;
      if (text !== undefined) { if (inT && stack.length) stack.forEach((b) => b.push(decode(text))); continue; }
      if (!tag) continue;
      if (tag === "p") {
        if (close) { const b = stack.pop(); if (b) paras.push(b.join("")); }
        else if (selfClose) paras.push("");
        else stack.push([]);
      } else if (tag === "t") {
        if (close) inT = 0; else if (!selfClose) inT = 1;
      } else if (!close && stack.length) {
        const ch = tag === "tab" ? " " : "\n";
        stack.forEach((b) => b.push(ch));
      }
    }
    return paras;
  }

  async function readDocxBlocks(buf) {
    const xml = await readZipEntry(buf, "word/document.xml");
    const blocks = [];
    let current = [];
    for (const text of paragraphTexts(xml)) {
      if (text.trim()) current.push(text);
      else if (current.length) { blocks.push(current); current = []; }
    }
    if (current.length) blocks.push(current);
    return blocks;
  }

  // ---------------- 3. Clean, split, group ----------------
  const isHindi = (t) => DEVANAGARI.test(t);
  const HAS_LETTER = /\p{L}/u;
  const NO_SPACE_BEFORE = /^[ऀ-ःऺ-ॏ॑-ॗॢॣ\s,.;:!?।॥)\]}”’'"-]/u;
  const NO_SPACE_AFTER = /[\s(\[{“‘'"-]$/u;

  function joinPieces(pieces) {
    let out = "";
    for (const piece of pieces) {
      if (out && !NO_SPACE_AFTER.test(out) && !NO_SPACE_BEFORE.test(piece)) out += " ";
      out += piece;
    }
    return out;
  }

  function hindiParagraphs(block) {
    const runs = [];
    for (const piece of block) {
      const lang = !HAS_LETTER.test(piece) && runs.length ? runs[runs.length - 1][0] : isHindi(piece) ? "hi" : "en";
      if (runs.length && runs[runs.length - 1][0] === lang) runs[runs.length - 1][1].push(piece);
      else runs.push([lang, [piece]]);
    }
    return runs.filter(([l]) => l === "hi").map(([, p]) => joinPieces(p));
  }

  function splitSentences(text) {
    const out = [];
    for (let chunk of text.split("\n")) {
      chunk = chunk.replace(/\s+/gu, " ").trim();
      for (const s of chunk.split(/(?<=[।॥.?!])\s+/u)) if (s.trim()) out.push(s.trim());
    }
    return out;
  }

  function groupSentences(sentences, size = SENTENCES_PER_BLOCK) {
    if (size <= 1) return sentences.map((s) => [s]);
    const blocks = [];
    for (let i = 0; i < sentences.length; i += size) blocks.push(sentences.slice(i, i + size));
    if (blocks.length > 1 && blocks[blocks.length - 1].length < size) blocks[blocks.length - 2].push(...blocks.pop());
    return blocks;
  }

  function linesFromParagraphs(paragraphs, rules) {
    const lines = [];
    for (const para of paragraphs) {
      const sentences = splitSentences(rules.apply(para)).map((s) => rules.addPauses(s));
      if (!sentences.length) continue;
      for (const group of groupSentences(sentences)) {
        if (lines.length && BLANK_LINE_BETWEEN_PARAGRAPHS) lines.push("");
        lines.push(group.join(" "));
      }
    }
    return lines;
  }

  async function processDocx(buf, rules) {
    const paras = [];
    for (const block of await readDocxBlocks(buf)) paras.push(...hindiParagraphs(block));
    return linesFromParagraphs(paras, rules);
  }

  // ---------------- XLIFF ----------------
  // Hindi paragraphs = the <target> of every translation unit, in order.
  // Works for XLIFF 1.2 (trans-unit) and 2.x (unit/segment). Units without a
  // Hindi target are skipped, never replaced by the English source.
  function readXliffParagraphs(xml) {
    const decode = (s) => s.replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-f]+);/gi, (_, e) =>
      ({ lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" }[e.toLowerCase()] ??
        String.fromCodePoint(e[1].toLowerCase() === "x" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10))));
    const textOf = (inner) => inner
      .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, (_, c) => c.replace(/&/g, "&amp;").replace(/</g, "&lt;"))
      .replace(/<[^>]*>/g, "");
    xml = xml.replace(/<!--[\s\S]*?-->/g, "");
    const paras = [];
    const unitRe = /<(?:[\w-]+:)?(trans-unit|unit)\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:[\w-]+:)?\1>)/g;
    for (const u of xml.matchAll(unitRe)) {
      const body = u[2] || "";
      const targets = [...body.matchAll(/<(?:[\w-]+:)?target\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:[\w-]+:)?target>)/g)]
        .map((t) => decode(textOf(t[1] || "")));
      const text = targets.join(" ").trim();
      if (text && isHindi(text)) paras.push(text);
    }
    return paras;
  }

  async function processXliff(buf, rules) {
    const xml = new TextDecoder("utf-8").decode(buf);
    return linesFromParagraphs(readXliffParagraphs(xml), rules);
  }

  const SUPPORTED = { docx: processDocx, xliff: processXliff, xlf: processXliff };

  async function processFile(buf, name, rules) {
    const ext = (name.match(/\.([^.]+)$/) || [, ""])[1].toLowerCase();
    if (!SUPPORTED[ext]) throw new Error("Not a .docx, .xliff or .xlf file");
    return SUPPORTED[ext](buf, rules);
  }

  root.HindiCleaner = { Rules, processDocx, processXliff, processFile, readXliffParagraphs };
})(typeof window !== "undefined" ? window : globalThis);
