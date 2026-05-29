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

export type Image = {
  src: string;
  alt: string;
  title?: string;
};

// Inline content within a block. Runs carry text + marks; images are
// atomic. Discriminated structurally — only Image carries `src`.
export type Inline = Run | Image;

export type Alignment = "left" | "center" | "right";

export type Block =
  | { kind: "paragraph"; runs: Inline[]; quoteDepth?: number }
  | { kind: "heading"; level: 1 | 2 | 3 | 4 | 5 | 6; runs: Inline[] }
  | {
      kind: "listItem";
      ordered: boolean;
      depth: number;
      listId: number;
      runs: Inline[];
      checked?: boolean;
    }
  | { kind: "codeBlock"; content: string; language?: string }
  | { kind: "thematicBreak" }
  | {
      kind: "table";
      header: Inline[][];
      rows: Inline[][][];
      alignments: Alignment[];
    };

// linkify converts bare URLs (https://example.com) into link tokens with
// the same shape as [label](href), so the existing flattenInline link
// handling covers them with no extra logic.
const md = new MarkdownIt({ html: false, linkify: true, typographer: false });

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
  // Tables don't nest in GFM and we only support them at the top level
  // (inside a list or blockquote they're dropped). One mutable context
  // is enough to accumulate cells across the table_open..table_close
  // span.
  let tableCtx: TableContext | null = null;

  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];

    if (tableCtx) {
      // While inside a table, only table-internal tokens are meaningful;
      // everything else is skipped until table_close.
      if (t.type === "table_close") {
        blocks.push({
          kind: "table",
          header: tableCtx.header,
          rows: tableCtx.rows,
          alignments: tableCtx.alignments,
        });
        tableCtx = null;
        continue;
      }
      if (t.type === "thead_open") {
        tableCtx.inHeader = true;
      } else if (t.type === "thead_close") {
        tableCtx.inHeader = false;
      } else if (t.type === "tr_open") {
        tableCtx.currentRow = [];
      } else if (t.type === "tr_close") {
        if (tableCtx.inHeader) {
          tableCtx.header = tableCtx.currentRow;
        } else {
          tableCtx.rows.push(tableCtx.currentRow);
        }
      } else if (t.type === "th_open" || t.type === "td_open") {
        if (tableCtx.inHeader) {
          tableCtx.alignments.push(cellAlignment(t));
        }
      } else if (t.type === "inline") {
        tableCtx.currentRow.push(flattenInline(t));
      }
      // tbody_open, tbody_close, th_close, td_close — fall through.
      continue;
    }

    if (t.type === "table_open" && blockquoteDepth === 0 && listStack.length === 0) {
      tableCtx = {
        header: [],
        rows: [],
        alignments: [],
        inHeader: false,
        currentRow: [],
      };
      continue;
    }

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
      const language = t.type === "fence" && t.info.trim() ? t.info.trim() : undefined;
      blocks.push({ kind: "codeBlock", content: t.content, language });
      continue;
    }

    if (t.type === "hr") {
      blocks.push({ kind: "thematicBreak" });
      continue;
    }

    if (t.type === "heading_open" || t.type === "paragraph_open") {
      // GFM task lists ride on a `[ ]` / `[x]` prefix in the first text
      // child. Detect (and strip) it before flattening so the prefix
      // doesn't leak into the runs.
      const checked =
        t.type === "paragraph_open" && listStack.length > 0
          ? consumeTaskPrefix(tokens[i + 1])
          : undefined;
      const runs = flattenInline(tokens[i + 1]);
      // Skip the inline token (i+1) and the matching _close (i+2); the
      // for-loop's i++ takes us past the close.
      i += 2;
      if (runs.length === 0 && checked === undefined) continue;
      if (t.type === "heading_open") {
        blocks.push({ kind: "heading", level: headingLevel(t.tag), runs });
      } else if (listStack.length > 0) {
        blocks.push(listItemBlock(listStack, currentListId, runs, checked));
      } else {
        blocks.push({ kind: "paragraph", runs });
      }
    }
  }
  return blocks;
}

type TableContext = {
  header: Inline[][];
  rows: Inline[][][];
  alignments: Alignment[];
  inHeader: boolean;
  currentRow: Inline[][];
};

function cellAlignment(token: Token): Alignment {
  // markdown-it serialises divider colons as a `text-align:...` style
  // attribute on each th/td. Default (no colon, no style) is left.
  const style = token.attrGet("style");
  if (style?.includes("text-align:center")) return "center";
  if (style?.includes("text-align:right")) return "right";
  return "left";
}

function listItemBlock(
  stack: Array<{ ordered: boolean }>,
  listId: number,
  runs: Inline[],
  checked: boolean | undefined
): Block {
  return {
    kind: "listItem",
    ordered: stack[stack.length - 1].ordered,
    depth: stack.length - 1,
    listId,
    runs,
    ...(checked !== undefined ? { checked } : {}),
  };
}

// Mutates the inline token's first text child to strip a leading
// `[ ]` / `[x]` / `[X]` task marker. Returns true for checked, false
// for unchecked, undefined when no marker is present.
function consumeTaskPrefix(token: Token | undefined): boolean | undefined {
  if (!token || token.type !== "inline") return undefined;
  const children = token.children ?? [];
  const first = children[0];
  if (!first || first.type !== "text") return undefined;
  const match = /^\[([ xX])\] ?/.exec(first.content);
  if (!match) return undefined;
  first.content = first.content.slice(match[0].length);
  return match[1] !== " ";
}

type InlineState = {
  boldDepth: number;
  italicDepth: number;
  strikeDepth: number;
  link: string | undefined;
};

function flattenInline(token: Token | undefined): Inline[] {
  if (!token || token.type !== "inline") return [];
  const out: Inline[] = [];
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
      case "image": {
        // Atomic — markdown-it's image token carries the src as an attr
        // and the alt text in `content`. Title is optional.
        const image: Image = {
          src: child.attrGet("src") ?? "",
          alt: child.content,
        };
        const title = child.attrGet("title");
        if (title) image.title = title;
        out.push(image);
        break;
      }
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

function pushRun(out: Inline[], run: Run): void {
  const prev = out[out.length - 1];
  if (prev && !("src" in prev) && sameFormat(prev, run)) {
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
