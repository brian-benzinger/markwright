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
  | { kind: "paragraph"; runs: Run[]; quoteDepth?: number }
  | { kind: "heading"; level: 1 | 2 | 3 | 4 | 5 | 6; runs: Run[] }
  | {
      kind: "listItem";
      ordered: boolean;
      depth: number;
      listId: number;
      runs: Run[];
    }
  | { kind: "codeBlock"; content: string; language?: string };

const md = new MarkdownIt({ html: false, linkify: false, typographer: false });

export function parseMarkdown(source: string): Block[] {
  const tokens = md.parse(source, {});
  const blocks: Block[] = [];
  // Tracks current blockquote nesting depth; > 0 turns inner blocks into
  // quote-styled paragraphs (lossy for headings and list items — they
  // collapse to flat quoted paragraphs, which is the conventional
  // rendering and keeps the AST simple).
  let blockquoteDepth = 0;
  // listStack tracks ordered-ness per nesting level; listId is a synthetic
  // counter that groups every item in one top-level Markdown list scope
  // (including nested sub-lists) so the applier knows when to start a fresh
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
    if (blockquoteDepth > 0) {
      // Inside a blockquote: emit any paragraph- or heading-bearing
      // content as a quote-styled paragraph at the current depth. List
      // and other container tokens are skipped so list semantics don't
      // collide with quote styling.
      if (t.type === "heading_open" || t.type === "paragraph_open") {
        const runs = flattenInline(tokens[i + 1]);
        i += 2;
        if (runs.length > 0) {
          blocks.push({ kind: "paragraph", runs, quoteDepth: blockquoteDepth });
        }
      }
      continue;
    }

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

    if (t.type === "fence" || t.type === "code_block") {
      const language =
        t.type === "fence" && t.info.trim() ? t.info.trim() : undefined;
      blocks.push({ kind: "codeBlock", content: t.content, language });
      continue;
    }

    if (t.type === "heading_open" || t.type === "paragraph_open") {
      const runs = flattenInline(tokens[i + 1]);
      // Skip the inline token (i+1) and the matching _close (i+2); the
      // for-loop's i++ takes us past the close.
      i += 2;
      if (runs.length === 0) continue;
      if (t.type === "heading_open") {
        blocks.push({ kind: "heading", level: headingLevel(t.tag), runs });
      } else if (listStack.length > 0) {
        blocks.push(listItemBlock(listStack, currentListId, runs));
      } else {
        blocks.push({ kind: "paragraph", runs });
      }
    }
  }
  return blocks;
}

function listItemBlock(
  stack: Array<{ ordered: boolean }>,
  listId: number,
  runs: Run[],
): Block {
  return {
    kind: "listItem",
    ordered: stack[stack.length - 1].ordered,
    depth: stack.length - 1,
    listId,
    runs,
  };
}

type InlineState = {
  boldDepth: number;
  italicDepth: number;
  strikeDepth: number;
  link: string | undefined;
};

function flattenInline(token: Token | undefined): Run[] {
  if (!token || token.type !== "inline") return [];
  const out: Run[] = [];
  const s: InlineState = {
    boldDepth: 0,
    italicDepth: 0,
    strikeDepth: 0,
    link: undefined,
  };

  for (const child of token.children ?? []) {
    switch (child.type) {
      case "text":
        if (child.content) pushRun(out, makeRun(child.content, s));
        break;
      case "code_inline":
        pushRun(out, makeRun(child.content, s, true));
        break;
      case "strong_open":
        s.boldDepth++;
        break;
      case "strong_close":
        s.boldDepth--;
        break;
      case "em_open":
        s.italicDepth++;
        break;
      case "em_close":
        s.italicDepth--;
        break;
      case "s_open":
        s.strikeDepth++;
        break;
      case "s_close":
        s.strikeDepth--;
        break;
      case "link_open":
        s.link = child.attrGet("href") ?? undefined;
        break;
      case "link_close":
        s.link = undefined;
        break;
      case "softbreak":
        pushRun(out, makeRun(" ", s));
        break;
      case "hardbreak":
        // U+000B is Word's in-paragraph line break.
        pushRun(out, makeRun("\v", s));
        break;
    }
  }
  return out;
}

function makeRun(text: string, s: InlineState, code = false): Run {
  const run: Run = { text };
  if (s.boldDepth > 0) run.bold = true;
  if (s.italicDepth > 0) run.italic = true;
  if (s.strikeDepth > 0) run.strike = true;
  if (code) run.code = true;
  if (s.link) run.link = s.link;
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
