import MarkdownIt from "markdown-it";
import type Token from "markdown-it/lib/token.mjs";

export type Run = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  code?: boolean;
  link?: string;
};

export type Block =
  | { kind: "paragraph"; runs: Run[] }
  | { kind: "heading"; level: 1 | 2 | 3 | 4 | 5 | 6; runs: Run[] }
  | {
      kind: "listItem";
      ordered: boolean;
      depth: number;
      listId: number;
      runs: Run[];
    };

const md = new MarkdownIt({ html: false, linkify: false, typographer: false });

export function parseMarkdown(source: string): Block[] {
  const tokens = md.parse(source, {});
  const blocks: Block[] = [];
  // Blockquotes aren't supported yet; suppress their contents until they
  // ship rather than emitting bare paragraphs that lose the quote semantics.
  let blockquoteDepth = 0;
  // listStack tracks ordered-ness per nesting level. listId is a synthetic
  // counter that groups every item in one top-level Markdown list (including
  // its nested sub-lists) so the applier knows when to start a fresh
  // Word.List vs. continue an existing one.
  const listStack: Array<{ ordered: boolean }> = [];
  let listIdCounter = 0;
  let currentListId = 0;

  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];

    if (t.type === "blockquote_open") {
      blockquoteDepth++;
      continue;
    }
    if (t.type === "blockquote_close") {
      blockquoteDepth--;
      continue;
    }
    if (blockquoteDepth > 0) continue;

    if (t.type === "bullet_list_open" || t.type === "ordered_list_open") {
      if (listStack.length === 0) {
        listIdCounter++;
        currentListId = listIdCounter;
      }
      listStack.push({ ordered: t.type === "ordered_list_open" });
      continue;
    }
    if (t.type === "bullet_list_close" || t.type === "ordered_list_close") {
      listStack.pop();
      continue;
    }
    if (t.type === "list_item_open" || t.type === "list_item_close") continue;

    if (t.type === "heading_open") {
      const runs = flattenInline(tokens[i + 1]);
      if (runs.length > 0) {
        blocks.push({ kind: "heading", level: headingLevel(t.tag), runs });
      }
      i += 2;
    } else if (t.type === "paragraph_open") {
      const runs = flattenInline(tokens[i + 1]);
      if (runs.length > 0) {
        if (listStack.length > 0) {
          const top = listStack[listStack.length - 1];
          blocks.push({
            kind: "listItem",
            ordered: top.ordered,
            depth: listStack.length - 1,
            listId: currentListId,
            runs,
          });
        } else {
          blocks.push({ kind: "paragraph", runs });
        }
      }
      i += 2;
    }
  }
  return blocks;
}

function flattenInline(token: Token | undefined): Run[] {
  if (!token || token.type !== "inline") return [];
  const out: Run[] = [];
  let bold = 0;
  let italic = 0;
  let strike = 0;
  let link: string | undefined;

  for (const child of token.children ?? []) {
    switch (child.type) {
      case "text":
        if (child.content) {
          pushRun(
            out,
            makeRun(child.content, bold, italic, strike, false, link),
          );
        }
        break;
      case "code_inline":
        pushRun(out, makeRun(child.content, bold, italic, strike, true, link));
        break;
      case "strong_open":
        bold++;
        break;
      case "strong_close":
        bold--;
        break;
      case "em_open":
        italic++;
        break;
      case "em_close":
        italic--;
        break;
      case "s_open":
        strike++;
        break;
      case "s_close":
        strike--;
        break;
      case "link_open":
        link = child.attrGet("href") ?? undefined;
        break;
      case "link_close":
        link = undefined;
        break;
      case "softbreak":
        pushRun(out, makeRun(" ", bold, italic, strike, false, link));
        break;
      case "hardbreak":
        // U+000B is Word's in-paragraph line break.
        pushRun(out, makeRun("\v", bold, italic, strike, false, link));
        break;
    }
  }
  return out;
}

function makeRun(
  text: string,
  bold: number,
  italic: number,
  strike: number,
  code: boolean,
  link: string | undefined,
): Run {
  const run: Run = { text };
  if (bold > 0) run.bold = true;
  if (italic > 0) run.italic = true;
  if (strike > 0) run.strike = true;
  if (code) run.code = true;
  if (link) run.link = link;
  return run;
}

function pushRun(out: Run[], run: Run): void {
  const prev = out[out.length - 1];
  if (prev && sameFormat(prev, run)) {
    prev.text += run.text;
  } else {
    out.push(run);
  }
}

function sameFormat(a: Run, b: Run): boolean {
  return (
    !!a.bold === !!b.bold &&
    !!a.italic === !!b.italic &&
    !!a.strike === !!b.strike &&
    !!a.code === !!b.code &&
    a.link === b.link
  );
}

function headingLevel(tag: string): 1 | 2 | 3 | 4 | 5 | 6 {
  // markdown-it heading_open tokens always carry tag "h1".."h6".
  return Number(tag[1]) as 1 | 2 | 3 | 4 | 5 | 6;
}
