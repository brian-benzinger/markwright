import MarkdownIt from "markdown-it";
import type Token from "markdown-it/lib/token.mjs";

export type Block =
  | { kind: "paragraph"; text: string }
  | { kind: "heading"; level: 1 | 2 | 3 | 4 | 5 | 6; text: string };

const md = new MarkdownIt({ html: false, linkify: false, typographer: false });

export function parseMarkdown(source: string): Block[] {
  const tokens = md.parse(source, {});
  const blocks: Block[] = [];
  // Skip block tokens nested inside list items / blockquotes until M3
  // handles those container types — otherwise their inner paragraphs
  // would surface as flat top-level paragraphs.
  let containerDepth = 0;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (
      t.type === "bullet_list_open" ||
      t.type === "ordered_list_open" ||
      t.type === "blockquote_open" ||
      t.type === "list_item_open"
    ) {
      containerDepth++;
      continue;
    }
    if (
      t.type === "bullet_list_close" ||
      t.type === "ordered_list_close" ||
      t.type === "blockquote_close" ||
      t.type === "list_item_close"
    ) {
      containerDepth--;
      continue;
    }
    if (containerDepth > 0) continue;

    if (t.type === "heading_open") {
      const level = headingLevel(t.tag);
      blocks.push({ kind: "heading", level, text: inlineText(tokens[i + 1]) });
      i += 2;
    } else if (t.type === "paragraph_open") {
      blocks.push({ kind: "paragraph", text: inlineText(tokens[i + 1]) });
      i += 2;
    }
    // M3 will add: bullet/ordered lists, blockquotes, fenced code, tables, marks.
  }
  return blocks;
}

function inlineText(token: Token | undefined): string {
  if (!token || token.type !== "inline") return "";
  const children = token.children ?? [];
  return children
    .filter((c) => c.type === "text")
    .map((c) => c.content)
    .join("");
}

function headingLevel(tag: string): 1 | 2 | 3 | 4 | 5 | 6 {
  switch (tag) {
    case "h1":
      return 1;
    case "h2":
      return 2;
    case "h3":
      return 3;
    case "h4":
      return 4;
    case "h5":
      return 5;
    case "h6":
      return 6;
    default:
      return 1;
  }
}
