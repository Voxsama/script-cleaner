// cleaner.js - browser port of hindi_extract.py. Same rules, same output.
// Works in the browser and in Node (for tests). No libraries.
(function (root) {
  "use strict";

  const SENTENCES_PER_BLOCK = 2;
  const BLANK_LINE_BETWEEN_PARAGRAPHS = true;
  // Start every output with the file's passage reference on its own line.
  const ADD_REFERENCE_HEADER = true;
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
      const sortedPauses = [...this.pauses].sort((a, b) => b[2].length - a[2].length);
      const appliedAfter = new Set();
      const appliedBefore = new Set();
      for (const [secs, where, phrase] of sortedPauses) {
        if (sentence.includes(phrase)) {
          if (where === "after") {
            if ([...appliedAfter].some((prev) => prev.includes(phrase))) continue;
            appliedAfter.add(phrase);
            const tags = Array(secs).fill(BREAK_TAG).join(" ");
            sentence = `${sentence} ${tags}`;
          } else if (where === "before") {
            if ([...appliedBefore].some((prev) => prev.includes(phrase))) continue;
            appliedBefore.add(phrase);
            const tags = Array(secs).fill(BREAK_TAG).join(" ");
            sentence = `${tags} ${sentence}`;
          }
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

  function quoteSpans(text) {
    const spans = [];
    let opening = null, start = null;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (opening === null) {
        if (["\"", "“", "‘", "`"].includes(ch)) { opening = ch; start = i; }
      } else if ((opening === "\"" && ["\"", "”"].includes(ch)) ||
                 (opening === "“" && ["”", "\""].includes(ch)) ||
                 (opening === "‘" && ["’", "'"].includes(ch)) ||
                 (opening === "`" && ch === "`")) {
        spans.push([start, i + 1]); opening = null; start = null;
      }
    }
    return spans;
  }

  function splitQuoteExamples(text) {
    const spans = quoteSpans(text);
    if (spans.length < 2) return [text];
    const chains = [];
    let chainStart = null, chainEnd = null;
    for (let i = 1; i < spans.length; i++) {
      const between = text.slice(spans[i - 1][1], spans[i][0]);
      if (/^\s*या(?:\s+[\u0900-\u097f]+){0,3}\s*$/u.test(between)) {
        if (chainStart === null) chainStart = i - 1;
        chainEnd = i;
      } else if (chainStart !== null) {
        chains.push([chainStart, chainEnd]); chainStart = chainEnd = null;
      }
    }
    if (chainStart !== null) chains.push([chainStart, chainEnd]);
    if (!chains.length) return [text];
    const [first, last] = chains.reduce((best, current) =>
      current[1] - current[0] > best[1] - best[0] ? current : best);
    if (last - first < 1) return [text];
    const intro = text.slice(0, spans[first][0]).trim();
    let end = spans[last][1];
    while (end < text.length && "।॥.,;:!?".includes(text[end])) end++;
    const alternatives = text.slice(spans[first][0], end).trim();
    const remainder = text.slice(end).trim();
    return [intro, alternatives, remainder].filter(Boolean);
  }

  // Quote marks: [opening, closing]. A straight " opens and closes, the curly
  // pair \u201c/\u201d and the single curly \u2018/\u2019 match each other.
  const QUOTE_PAIRS = [['"', '"'], ["\u201c", "\u201d"], ["\u2018", "\u2019"], ["`", "`"]];
  const SENTENCE_CHARS = "\u0964\u0965.?!\uff01\uff1f";

  // Split into sentences. A sentence end inside an open quotation is not a
  // break, so quoted speech stays in one piece and the text after it flows.
  function splitSentences(text) {
    const out = [];
    for (let chunk of text.split("\n")) {
      chunk = chunk.replace(/\s+/gu, " ").trim();
      if (!chunk) continue;
      let start = 0, closer = null, i = 0;
      while (i < chunk.length) {
        const ch = chunk[i];
        if (closer === null) {
          for (const [o, c] of QUOTE_PAIRS) if (ch === o) { closer = c; break; }
        } else if (ch === closer) {
          closer = null;
        }
        if (closer === null && SENTENCE_CHARS.includes(ch) && chunk[i + 1] === " ") {
          let j = i + 1;
          while (j < chunk.length && chunk[j] === " ") j++;
          const piece = chunk.slice(start, j).trim();
          if (piece) out.push(piece);
          start = j; i = j; continue;
        }
        i++;
      }
      const piece = chunk.slice(start).trim();
      if (piece) out.push(piece);
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

  // Book names of more than one word; a leading 1/2/3 (1 शमूएल) is handled too.
  const MULTI_WORD_BOOKS = ["भजन संहिता", "प्रेरितों के काम", "श्रेष्ठ गीत"];
  const reEscape = (t) => t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const REFERENCE_RE = new RegExp(
    "(?:[123]\\s+)?(?:" + MULTI_WORD_BOOKS.map(reEscape).join("|") + `|${DEV_WORD})` +
    "\\s+\\d+\\s*:\\s*\\d+(?:\\s*[-–—]\\s*\\d+)?", "u");

  function findReference(paragraphs, rules) {
    for (const para of paragraphs) {
      const m = REFERENCE_RE.exec(para);
      if (m) return rules.apply(m[0]).trim();
    }
    return null;
  }

  function withHeader(ref, body) {
    if (!(ADD_REFERENCE_HEADER && ref)) return [...body];
    return body.length ? [ref, "", ...body] : [ref];
  }

  function linesFromParagraphs(paragraphs, rules) {
    return withHeader(findReference(paragraphs, rules), bodyLines(paragraphs, rules));
  }

  function mergeYaParagraphs(paragraphs) {
    const merged = [];
    for (const p of paragraphs) {
      const pStrip = p.trim();
      if (!pStrip) continue;
      const endsWithYa = merged.length && /(?:[;:"”'’`\-—–]\s*)?या\s*$/u.test(merged[merged.length - 1].trim());
      const startsWithYa = merged.length && /^(?:[;:"“'’`\-—–]\s*)?या(?:\s|$)/u.test(pStrip);
      if (endsWithYa || startsWithYa) {
        merged[merged.length - 1] = merged[merged.length - 1].trimEnd() + " " + pStrip.trimStart();
      } else {
        merged.push(pStrip);
      }
    }
    return merged;
  }

  function hasYaAlternatives(text) {
    return /(?:[;"”'’`]\s*|^)या\s+/u.test(text) || (text.split(" या ").length > 1 && (text.includes(";") || text.includes('"') || text.includes('”')));
  }

  // The cleaned text of a file, without the reference header.
  function bodyLines(paragraphs, rules) {
    const lines = [];
    for (const para of mergeYaParagraphs(paragraphs)) {
      for (const piece of splitQuoteExamples(rules.apply(para))) {
        const sentences = splitSentences(piece).map((s) => rules.addPauses(s));
        if (!sentences.length) continue;
        if (hasYaAlternatives(piece)) {
          if (lines.length && BLANK_LINE_BETWEEN_PARAGRAPHS) lines.push("");
          lines.push(sentences.join(" "));
          continue;
        }
        if (SENTENCES_PER_BLOCK <= 1) {
          if (lines.length && BLANK_LINE_BETWEEN_PARAGRAPHS) lines.push("");
          lines.push(...sentences);
          continue;
        }
        for (const group of groupSentences(sentences)) {
          if (lines.length && BLANK_LINE_BETWEEN_PARAGRAPHS) lines.push("");
          lines.push(group.join(" "));
        }
      }
    }
    return lines;
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

  // Hindi paragraphs of a .docx or .xliff file.
  async function readParagraphs(buf, name) {
    const ext = (name.match(/\.([^.]+)$/) || [, ""])[1].toLowerCase();
    if (ext === "docx") {
      const paras = [];
      for (const block of await readDocxBlocks(buf)) paras.push(...hindiParagraphs(block));
      return paras;
    }
    if (ext === "xliff" || ext === "xlf") return readXliffParagraphs(new TextDecoder("utf-8").decode(buf));
    throw new Error("Not a .docx, .xliff or .xlf file");
  }

  // Convert one file on its own (no borrowing from other files).
  async function processFile(buf, name, rules) {
    return linesFromParagraphs(await readParagraphs(buf, name), rules);
  }

  // {ref: own reference or null, body: lines} for one file.
  async function fileParts(buf, name, rules) {
    const paras = await readParagraphs(buf, name);
    return { ref: findReference(paras, rules), body: bodyLines(paras, rules) };
  }

  // ---------------- Borrowing a missing reference ----------------
  // Files of the same passage share a name apart from the section part:
  // exo-p76-s2-v1... and exo-p76-s3-v1... -> "exo-p76-v1". null if no "-s<number>".
  function passageKey(name) {
    const stem = name.replace(/^.*[\\/]/, "").split(".")[0];
    const re = /-s\d+(?=-|$)/i;
    return re.test(stem) ? stem.replace(re, "").toLowerCase() : null;
  }

  // items: [{name, ref}] -> adds header and headerFrom (null = own reference,
  // a file name = copied from that file). Copies only from the same passage;
  // for names without a section part, only when all references agree.
  function fillReferences(items) {
    const own = items.filter((it) => it.ref).sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0));
    for (const it of items) {
      it.header = it.ref || null; it.headerFrom = null;
      if (it.header) continue;
      const key = passageKey(it.name);
      const donor = key ? own.find((d) => passageKey(d.name) === key)
        : own.length && new Set(own.map((d) => d.ref)).size === 1 ? own[0] : null;
      if (donor) { it.header = donor.ref; it.headerFrom = donor.name; }
    }
    return items;
  }

  root.HindiCleaner = { Rules, processFile, fileParts, withHeader, passageKey, fillReferences, splitQuoteExamples,
    readXliffParagraphs, findReference, ADD_REFERENCE_HEADER };
})(typeof window !== "undefined" ? window : globalThis);
