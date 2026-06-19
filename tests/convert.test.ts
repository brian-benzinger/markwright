import { describe, it, expect } from "vitest";
import { parseMarkdown } from "../src/convert";

describe("parseMarkdown — block shape", () => {
  it("returns an empty array for empty input", () => {
    expect(parseMarkdown("")).toEqual([]);
  });

  it("returns an empty array for whitespace-only input", () => {
    expect(parseMarkdown("   \n  \n")).toEqual([]);
  });

  it("parses a single paragraph as one plain run", () => {
    expect(parseMarkdown("hello world")).toEqual([
      { kind: "paragraph", runs: [{ text: "hello world" }] },
    ]);
  });

  it.each([
    [1, "# heading text"],
    [2, "## heading text"],
    [3, "### heading text"],
    [4, "#### heading text"],
    [5, "##### heading text"],
    [6, "###### heading text"],
  ])("parses heading level %i", (level, source) => {
    expect(parseMarkdown(source)).toEqual([
      { kind: "heading", level, runs: [{ text: "heading text" }] },
    ]);
  });

  it("preserves block order across mixed input", () => {
    expect(parseMarkdown("# H\n\np1\n\n## S\n\np2")).toEqual([
      { kind: "heading", level: 1, runs: [{ text: "H" }] },
      { kind: "paragraph", runs: [{ text: "p1" }] },
      { kind: "heading", level: 2, runs: [{ text: "S" }] },
      { kind: "paragraph", runs: [{ text: "p2" }] },
    ]);
  });

  it("preserves XML special characters verbatim", () => {
    // The Office.js applier takes raw strings, so we don't escape `&`, `<`,
    // `>` in the parser. If we ever ship an OOXML emitter (for tables /
    // images / footnotes), that path will need to escape on its own.
    expect(parseMarkdown("a & b < c > d")).toEqual([
      { kind: "paragraph", runs: [{ text: "a & b < c > d" }] },
    ]);
  });

  it("drops an empty ATX heading with no text after the marker", () => {
    // `#` alone produces an inline token with no children; runs.length === 0
    // and checked === undefined, so the parser skips it rather than emitting
    // an empty heading block.
    expect(parseMarkdown("#")).toEqual([]);
  });

  it("drops a top-level paragraph whose inline content resolves to no runs", () => {
    // [](url) produces link_open/link_close with no text child, so
    // flattenInline returns []; with checked === undefined the guard at the
    // paragraph_open branch silently skips it.
    expect(parseMarkdown("[](url)")).toEqual([]);
  });
});

describe("parseMarkdown — inline marks", () => {
  it("parses bold", () => {
    expect(parseMarkdown("a **bold** b")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "a " }, { text: "bold", bold: true }, { text: " b" }],
      },
    ]);
  });

  it("parses italic from asterisks", () => {
    expect(parseMarkdown("a *it* b")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "a " }, { text: "it", italic: true }, { text: " b" }],
      },
    ]);
  });

  it("parses italic from underscores", () => {
    expect(parseMarkdown("a _it_ b")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "a " }, { text: "it", italic: true }, { text: " b" }],
      },
    ]);
  });

  it("parses nested bold + italic", () => {
    expect(parseMarkdown("***both***")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "both", bold: true, italic: true }],
      },
    ]);
  });

  it("parses strikethrough", () => {
    expect(parseMarkdown("a ~~gone~~ b")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "a " }, { text: "gone", strike: true }, { text: " b" }],
      },
    ]);
  });

  it("parses inline code", () => {
    expect(parseMarkdown("call `foo()` now")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "call " }, { text: "foo()", code: true }, { text: " now" }],
      },
    ]);
  });

  it("parses a link with URL", () => {
    expect(parseMarkdown("see [docs](https://example.com)")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "see " }, { text: "docs", link: "https://example.com" }],
      },
    ]);
  });

  it("parses an autolink as a link with text == href", () => {
    expect(parseMarkdown("at <https://example.com>")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "at " }, { text: "https://example.com", link: "https://example.com" }],
      },
    ]);
  });

  it("treats backslash escapes as literal text", () => {
    expect(parseMarkdown("\\*not bold\\*")).toEqual([
      { kind: "paragraph", runs: [{ text: "*not bold*" }] },
    ]);
  });

  it("converts a soft line break to a single space", () => {
    expect(parseMarkdown("line one\nline two")).toEqual([
      { kind: "paragraph", runs: [{ text: "line one line two" }] },
    ]);
  });

  it("converts a hard line break (two trailing spaces) to U+000B", () => {
    expect(parseMarkdown("line one  \nline two")).toEqual([
      { kind: "paragraph", runs: [{ text: "line one\vline two" }] },
    ]);
  });

  it("applies marks inside headings", () => {
    expect(parseMarkdown("# **bold** heading")).toEqual([
      {
        kind: "heading",
        level: 1,
        runs: [{ text: "bold", bold: true }, { text: " heading" }],
      },
    ]);
  });

  it("applies strikethrough inside a heading", () => {
    expect(parseMarkdown("# ~~struck~~ heading")).toEqual([
      {
        kind: "heading",
        level: 1,
        runs: [{ text: "struck", strike: true }, { text: " heading" }],
      },
    ]);
  });

  it("combines strikethrough and link in a single run", () => {
    // link_open sets s.link; s_open sets s.strikeDepth — both flags land on
    // the same run because flattenInline reads the full InlineState at push time.
    expect(parseMarkdown("[~~gone~~](https://x)")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "gone", strike: true, link: "https://x" }],
      },
    ]);
  });

  it("merges adjacent runs with identical formatting", () => {
    // A softbreak inside a bold span produces three adjacent bold runs
    // (text, space, text) — they should collapse into one.
    expect(parseMarkdown("**one\ntwo**")).toEqual([
      { kind: "paragraph", runs: [{ text: "one two", bold: true }] },
    ]);
  });

  it("composes bold and inline code in a single run", () => {
    expect(parseMarkdown("**bold `code` bold**")).toEqual([
      {
        kind: "paragraph",
        runs: [
          { text: "bold ", bold: true },
          { text: "code", bold: true, code: true },
          { text: " bold", bold: true },
        ],
      },
    ]);
  });

  it("composes a link and bold into a single bold-link run", () => {
    expect(parseMarkdown("[**bold link**](https://x)")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "bold link", bold: true, link: "https://x" }],
      },
    ]);
  });

  it("composes a link and italic into a single italic-link run", () => {
    expect(parseMarkdown("[*italic link*](https://x)")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "italic link", italic: true, link: "https://x" }],
      },
    ]);
  });

  it("composes a link with bold and italic into a single run carrying all three marks", () => {
    expect(parseMarkdown("[***bold italic link***](https://x)")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "bold italic link", bold: true, italic: true, link: "https://x" }],
      },
    ]);
  });

  it("composes italic and inline code in a single run", () => {
    expect(parseMarkdown("*italic `code` italic*")).toEqual([
      {
        kind: "paragraph",
        runs: [
          { text: "italic ", italic: true },
          { text: "code", italic: true, code: true },
          { text: " italic", italic: true },
        ],
      },
    ]);
  });
});

describe("parseMarkdown — lists", () => {
  it("parses a flat bullet list", () => {
    expect(parseMarkdown("- a\n- b\n- c")).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "a" }],
      },
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "b" }],
      },
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "c" }],
      },
    ]);
  });

  it("parses a flat ordered list", () => {
    expect(parseMarkdown("1. one\n2. two")).toEqual([
      {
        kind: "listItem",
        ordered: true,
        depth: 0,
        listId: 1,
        runs: [{ text: "one" }],
      },
      {
        kind: "listItem",
        ordered: true,
        depth: 0,
        listId: 1,
        runs: [{ text: "two" }],
      },
    ]);
  });

  it("assigns a fresh listId to each separate top-level list", () => {
    // Blank line between the lists splits them into two scopes; the
    // second list should get a new listId so the applier knows to start
    // a new Word.List instead of continuing numbering.
    expect(parseMarkdown("- a\n\n1. b")).toEqual([
      { kind: "listItem", ordered: false, depth: 0, listId: 1, runs: [{ text: "a" }] },
      { kind: "listItem", ordered: true, depth: 0, listId: 2, runs: [{ text: "b" }] },
    ]);
  });

  it("tracks depth and ordered-ness on nested lists", () => {
    const out = parseMarkdown("- a\n  1. nested\n  2. also\n- b");
    expect(out).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "a" }],
      },
      {
        kind: "listItem",
        ordered: true,
        depth: 1,
        listId: 1,
        runs: [{ text: "nested" }],
      },
      {
        kind: "listItem",
        ordered: true,
        depth: 1,
        listId: 1,
        runs: [{ text: "also" }],
      },
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "b" }],
      },
    ]);
  });

  it("keeps a nested list within the same listId as its parent", () => {
    expect(parseMarkdown("- top\n  - nested")).toEqual([
      { kind: "listItem", ordered: false, depth: 0, listId: 1, runs: [{ text: "top" }] },
      { kind: "listItem", ordered: false, depth: 1, listId: 1, runs: [{ text: "nested" }] },
    ]);
  });

  it("converts a hard line break inside a list item to U+000B", () => {
    // Two trailing spaces before a continued list line produce a hardbreak
    // token that flattenInline converts to "\v", same as in a plain paragraph.
    expect(parseMarkdown("- line one  \n  line two")).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "line one\vline two" }],
      },
    ]);
  });

  it("applies inline marks inside list items", () => {
    expect(parseMarkdown("- a **bold** item")).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "a " }, { text: "bold", bold: true }, { text: " item" }],
      },
    ]);
  });

  it("applies a link inside a list item", () => {
    expect(parseMarkdown("- see [docs](https://example.com) here")).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "see " }, { text: "docs", link: "https://example.com" }, { text: " here" }],
      },
    ]);
  });

  it("applies inline code inside a list item", () => {
    expect(parseMarkdown("- call `foo()` now")).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "call " }, { text: "foo()", code: true }, { text: " now" }],
      },
    ]);
  });

  it("interleaves headings and lists without dropping order", () => {
    // Also verifies: heading level, list orderedness, fresh listId per
    // top-level list scope, and text content — not just block kinds.
    expect(parseMarkdown("# H\n\n- a\n\np\n\n1. one")).toEqual([
      { kind: "heading", level: 1, runs: [{ text: "H" }] },
      { kind: "listItem", ordered: false, depth: 0, listId: 1, runs: [{ text: "a" }] },
      { kind: "paragraph", runs: [{ text: "p" }] },
      { kind: "listItem", ordered: true, depth: 0, listId: 2, runs: [{ text: "one" }] },
    ]);
  });

  it("applies strikethrough inside a list item", () => {
    expect(parseMarkdown("- ~~struck~~ item")).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "struck", strike: true }, { text: " item" }],
      },
    ]);
  });

  it("drops a list item whose inline content resolves to no runs and has no task marker", () => {
    // [](url) inside a list item: consumeTaskPrefix returns undefined (first
    // child is link_open, not text), flattenInline returns [] — both arms of
    // the `runs.length === 0 && checked === undefined` guard are true, so the
    // item is silently skipped.
    expect(parseMarkdown("- [](url)")).toEqual([]);
  });

  it("emits each paragraph in a loose list item as a separate listItem block", () => {
    // A loose list (blank lines between items) wraps item text in
    // paragraph_open tokens. A continuation paragraph within the same
    // list_item_open…close is therefore emitted as a second listItem at
    // the same depth and listId, because the parser treats every
    // paragraph_open inside a list as one list-item block.
    expect(parseMarkdown("- item one\n\n  continued\n\n- item two")).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "item one" }],
      },
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "continued" }],
      },
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "item two" }],
      },
    ]);
  });
});

describe("parseMarkdown — thematic breaks", () => {
  it.each([
    ["dashes", "---"],
    ["asterisks", "***"],
    ["underscores", "___"],
  ])("parses %s as a thematic break", (_, source) => {
    expect(parseMarkdown(source)).toEqual([{ kind: "thematicBreak" }]);
  });

  it("interleaves a thematic break between paragraphs", () => {
    expect(parseMarkdown("before\n\n---\n\nafter")).toEqual([
      { kind: "paragraph", runs: [{ text: "before" }] },
      { kind: "thematicBreak" },
      { kind: "paragraph", runs: [{ text: "after" }] },
    ]);
  });

  it("drops a thematic break inside a blockquote", () => {
    // Same lossy stance as code blocks inside blockquotes: rare in
    // practice, and a quoted HR has no clean Word representation.
    expect(parseMarkdown("> a\n>\n> ---\n>\n> b")).toEqual([
      { kind: "paragraph", runs: [{ text: "a" }], quoteDepth: 1 },
      { kind: "paragraph", runs: [{ text: "b" }], quoteDepth: 1 },
    ]);
  });
});

describe("parseMarkdown — code blocks", () => {
  it("parses a fenced code block", () => {
    expect(parseMarkdown("```\nhello\nworld\n```")).toEqual([
      { kind: "codeBlock", content: "hello\nworld\n" },
    ]);
  });

  it("captures the language tag from a fence", () => {
    expect(parseMarkdown("```javascript\nfoo();\n```")).toEqual([
      { kind: "codeBlock", content: "foo();\n", language: "javascript" },
    ]);
  });

  it("trims whitespace around the language tag", () => {
    expect(parseMarkdown("```  ts \nx\n```")).toEqual([
      { kind: "codeBlock", content: "x\n", language: "ts" },
    ]);
  });

  it("parses an indented (4-space) code block", () => {
    expect(parseMarkdown("    hello\n    world")).toEqual([
      { kind: "codeBlock", content: "hello\nworld\n" },
    ]);
  });

  it("does not parse markdown syntax inside a code block", () => {
    expect(parseMarkdown("```\n**bold** `code`\n```")).toEqual([
      { kind: "codeBlock", content: "**bold** `code`\n" },
    ]);
  });

  it("preserves leading whitespace and blank lines inside the block", () => {
    expect(parseMarkdown("```\n  indented\n\n  more\n```")).toEqual([
      { kind: "codeBlock", content: "  indented\n\n  more\n" },
    ]);
  });

  it("treats an empty fence as an empty code block (no language)", () => {
    expect(parseMarkdown("```\n```")).toEqual([{ kind: "codeBlock", content: "" }]);
  });

  it("interleaves a code block between paragraphs", () => {
    expect(parseMarkdown("before\n\n```\ncode\n```\n\nafter")).toEqual([
      { kind: "paragraph", runs: [{ text: "before" }] },
      { kind: "codeBlock", content: "code\n" },
      { kind: "paragraph", runs: [{ text: "after" }] },
    ]);
  });

  it("emits a fenced code block inside a list item as a standalone codeBlock", () => {
    // Unlike tables (which are gated on listStack.length === 0), the fence
    // handler has no list guard — code fences inside list items fall through
    // and are emitted as top-level codeBlock blocks after their list item.
    expect(parseMarkdown("- item\n\n  ```\n  code\n  ```")).toEqual([
      { kind: "listItem", ordered: false, depth: 0, listId: 1, runs: [{ text: "item" }] },
      { kind: "codeBlock", content: "code\n" },
    ]);
  });
});

describe("parseMarkdown — task lists", () => {
  it("parses an unchecked task as checked: false", () => {
    expect(parseMarkdown("- [ ] todo")).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "todo" }],
        checked: false,
      },
    ]);
  });

  it.each([
    ["lowercase", "- [x] done"],
    ["uppercase", "- [X] done"],
  ])("parses a checked task (%s x) as checked: true", (_, source) => {
    expect(parseMarkdown(source)).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "done" }],
        checked: true,
      },
    ]);
  });

  it("preserves inline marks inside a task item", () => {
    expect(parseMarkdown("- [ ] a **bold** task")).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "a " }, { text: "bold", bold: true }, { text: " task" }],
        checked: false,
      },
    ]);
  });

  it("mixes task and non-task items in one list", () => {
    expect(parseMarkdown("- [ ] todo\n- plain\n- [x] done")).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "todo" }],
        checked: false,
      },
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "plain" }],
      },
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "done" }],
        checked: true,
      },
    ]);
  });

  it("supports task markers in ordered lists", () => {
    expect(parseMarkdown("1. [ ] step one\n2. [x] step two")).toEqual([
      {
        kind: "listItem",
        ordered: true,
        depth: 0,
        listId: 1,
        runs: [{ text: "step one" }],
        checked: false,
      },
      {
        kind: "listItem",
        ordered: true,
        depth: 0,
        listId: 1,
        runs: [{ text: "step two" }],
        checked: true,
      },
    ]);
  });

  it("emits an empty task item when the marker is the only content", () => {
    expect(parseMarkdown("- [ ]")).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [],
        checked: false,
      },
    ]);
  });

  it("treats a bracketed string that isn't a real task marker as literal text", () => {
    // [todo] is not a valid task marker — the inner character must be
    // exactly one of " ", "x", "X".
    expect(parseMarkdown("- [todo] item")).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "[todo] item" }],
      },
    ]);
  });

  it("emits a plain list item (no checked field) when the item starts with inline markup", () => {
    // consumeTaskPrefix reads children[0]; when the first child is
    // strong_open (not a text node), it returns undefined and the item
    // is emitted without a checked property — same as a non-task item.
    expect(parseMarkdown("- **bold item**")).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "bold item", bold: true }],
      },
    ]);
  });

  it("same first-child guard applies to ordered list items starting with markup", () => {
    expect(parseMarkdown("1. *italic step*")).toEqual([
      {
        kind: "listItem",
        ordered: true,
        depth: 0,
        listId: 1,
        runs: [{ text: "italic step", italic: true }],
      },
    ]);
  });

  it("correctly strips a task marker with no space between ] and the item text", () => {
    // The regex uses `] ?` (optional trailing space) so `[x]done` is valid.
    expect(parseMarkdown("- [x]done")).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ text: "done" }],
        checked: true,
      },
    ]);
  });
});

describe("parseMarkdown — bare-URL autolinks", () => {
  it("wraps a bare URL in a link run", () => {
    expect(parseMarkdown("https://example.com")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "https://example.com", link: "https://example.com" }],
      },
    ]);
  });

  it("autolinks a URL mid-paragraph and preserves surrounding text", () => {
    expect(parseMarkdown("see https://example.com today")).toEqual([
      {
        kind: "paragraph",
        runs: [
          { text: "see " },
          { text: "https://example.com", link: "https://example.com" },
          { text: " today" },
        ],
      },
    ]);
  });
});

describe("parseMarkdown — blockquotes", () => {
  it("emits a single-line blockquote as a quoted paragraph", () => {
    expect(parseMarkdown("> quoted line")).toEqual([
      { kind: "paragraph", runs: [{ text: "quoted line" }], quoteDepth: 1 },
    ]);
  });

  it("collapses a multi-line blockquote into one paragraph with a soft break", () => {
    // A bare newline inside a blockquote is a softbreak, not a paragraph
    // break — so `> a\n> b` is one quoted paragraph with " " between.
    expect(parseMarkdown("> a\n> b")).toEqual([
      { kind: "paragraph", runs: [{ text: "a b" }], quoteDepth: 1 },
    ]);
  });

  it("emits one paragraph per blank-line-separated chunk inside the blockquote", () => {
    expect(parseMarkdown("> a\n>\n> b")).toEqual([
      { kind: "paragraph", runs: [{ text: "a" }], quoteDepth: 1 },
      { kind: "paragraph", runs: [{ text: "b" }], quoteDepth: 1 },
    ]);
  });

  it("preserves surrounding blocks in order", () => {
    expect(parseMarkdown("before\n\n> quoted\n\nafter")).toEqual([
      { kind: "paragraph", runs: [{ text: "before" }] },
      { kind: "paragraph", runs: [{ text: "quoted" }], quoteDepth: 1 },
      { kind: "paragraph", runs: [{ text: "after" }] },
    ]);
  });

  it("tracks nested quote depth", () => {
    expect(parseMarkdown("> outer\n>\n> > inner")).toEqual([
      { kind: "paragraph", runs: [{ text: "outer" }], quoteDepth: 1 },
      { kind: "paragraph", runs: [{ text: "inner" }], quoteDepth: 2 },
    ]);
  });

  it("downgrades a heading inside a blockquote to a quoted paragraph (lossy)", () => {
    expect(parseMarkdown("> # not really a heading")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "not really a heading" }],
        quoteDepth: 1,
      },
    ]);
  });

  it("preserves inline marks inside a blockquote", () => {
    expect(parseMarkdown("> a **bold** quote")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "a " }, { text: "bold", bold: true }, { text: " quote" }],
        quoteDepth: 1,
      },
    ]);
  });

  it("preserves a link inside a blockquote", () => {
    expect(parseMarkdown("> see [docs](https://example.com) here")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "see " }, { text: "docs", link: "https://example.com" }, { text: " here" }],
        quoteDepth: 1,
      },
    ]);
  });

  it("converts a hard line break inside a blockquote to U+000B", () => {
    // Two trailing spaces before a continued blockquote line produce a hardbreak
    // token that flattenInline converts to "\v", same as in a plain paragraph.
    expect(parseMarkdown("> line one  \n> line two")).toEqual([
      { kind: "paragraph", runs: [{ text: "line one\vline two" }], quoteDepth: 1 },
    ]);
  });

  it("drops a code block inside a blockquote (lossy)", () => {
    // fence tokens fall through the blockquote branch that only handles
    // paragraph_open and heading_open; surrounding quoted paragraphs survive.
    expect(parseMarkdown("> before\n>\n> ```\n> code\n> ```\n>\n> after")).toEqual([
      { kind: "paragraph", runs: [{ text: "before" }], quoteDepth: 1 },
      { kind: "paragraph", runs: [{ text: "after" }], quoteDepth: 1 },
    ]);
  });

  it("flattens list items inside a blockquote to quoted paragraphs (lossy)", () => {
    expect(parseMarkdown("> - one\n> - two")).toEqual([
      { kind: "paragraph", runs: [{ text: "one" }], quoteDepth: 1 },
      { kind: "paragraph", runs: [{ text: "two" }], quoteDepth: 1 },
    ]);
  });

  it("renders an image inside a blockquote as a quoted paragraph", () => {
    expect(parseMarkdown("> ![logo](https://x/a.png)")).toEqual([
      { kind: "paragraph", runs: [{ src: "https://x/a.png", alt: "logo" }], quoteDepth: 1 },
    ]);
  });

  it("drops a blockquote paragraph whose inline content resolves to no runs", () => {
    // A link with no visible text ([](url)) produces link_open/link_close
    // tokens but no text child, so flattenInline returns []. The guard
    // silently skips emitting an empty paragraph block.
    expect(parseMarkdown("> [](url)")).toEqual([]);
  });

  it("preserves strikethrough inside a blockquote", () => {
    expect(parseMarkdown("> ~~struck~~")).toEqual([
      { kind: "paragraph", runs: [{ text: "struck", strike: true }], quoteDepth: 1 },
    ]);
  });
});

describe("parseMarkdown — tables", () => {
  it("parses a simple table with default left alignment", () => {
    const src = "| H1 | H2 |\n| --- | --- |\n| a | b |";
    expect(parseMarkdown(src)).toEqual([
      {
        kind: "table",
        header: [[{ text: "H1" }], [{ text: "H2" }]],
        rows: [[[{ text: "a" }], [{ text: "b" }]]],
        alignments: ["left", "left"],
      },
    ]);
  });

  it("captures per-column alignment from the divider row", () => {
    const src = "| L | C | R |\n| :--- | :---: | ---: |\n| a | b | c |";
    expect(parseMarkdown(src)).toEqual([
      {
        kind: "table",
        header: [[{ text: "L" }], [{ text: "C" }], [{ text: "R" }]],
        rows: [[[{ text: "a" }], [{ text: "b" }], [{ text: "c" }]]],
        alignments: ["left", "center", "right"],
      },
    ]);
  });

  it("preserves inline marks inside cells", () => {
    const src = "| Plain | Styled |\n| --- | --- |\n| a | **b** `c` |";
    expect(parseMarkdown(src)).toEqual([
      {
        kind: "table",
        header: [[{ text: "Plain" }], [{ text: "Styled" }]],
        rows: [
          [[{ text: "a" }], [{ text: "b", bold: true }, { text: " " }, { text: "c", code: true }]],
        ],
        alignments: ["left", "left"],
      },
    ]);
  });

  it("preserves inline marks in header cells", () => {
    // Header cells go through the same flattenInline path as body cells;
    // this guards against a regression where header inline content is dropped.
    const src = "| [Link](https://x) | **Bold** |\n| --- | --- |\n| a | b |";
    expect(parseMarkdown(src)).toEqual([
      {
        kind: "table",
        header: [[{ text: "Link", link: "https://x" }], [{ text: "Bold", bold: true }]],
        rows: [[[{ text: "a" }], [{ text: "b" }]]],
        alignments: ["left", "left"],
      },
    ]);
  });

  it("handles multiple body rows", () => {
    const src = "| H |\n| --- |\n| a |\n| b |\n| c |";
    expect(parseMarkdown(src)).toEqual([
      {
        kind: "table",
        header: [[{ text: "H" }]],
        rows: [[[{ text: "a" }]], [[{ text: "b" }]], [[{ text: "c" }]]],
        alignments: ["left"],
      },
    ]);
  });

  it("interleaves a table with surrounding blocks", () => {
    // Full check: surrounding paragraph texts and table structure, not just kinds.
    expect(parseMarkdown("before\n\n| H |\n| --- |\n| x |\n\nafter")).toEqual([
      { kind: "paragraph", runs: [{ text: "before" }] },
      {
        kind: "table",
        header: [[{ text: "H" }]],
        rows: [[[{ text: "x" }]]],
        alignments: ["left"],
      },
      { kind: "paragraph", runs: [{ text: "after" }] },
    ]);
  });

  it("drops a table nested inside a blockquote (lossy)", () => {
    // Same lossy stance as code blocks / hr inside blockquotes —
    // table tokens fall through the inside-blockquote branch.
    const src = "> | H |\n> | --- |\n> | x |";
    expect(parseMarkdown(src)).toEqual([]);
  });

  it("drops a table nested inside a list (lossy)", () => {
    // The list_open guard keeps the table branch inert; the table
    // tokens drift through and produce nothing. The list item before
    // the table must still be emitted — this test verifies both sides.
    const src = "- item\n  | H |\n  | --- |\n  | x |";
    expect(parseMarkdown(src)).toEqual([
      { kind: "listItem", ordered: false, depth: 0, listId: 1, runs: [{ text: "item" }] },
    ]);
  });

  it("parses two consecutive tables as separate blocks", () => {
    // Verifies the tableCtx is reset to null after table_close so the
    // second table_open starts a fresh context rather than corrupting it.
    const src = "| A |\n| --- |\n| r1 |\n\n| B |\n| --- |\n| r2 |";
    expect(parseMarkdown(src)).toEqual([
      {
        kind: "table",
        header: [[{ text: "A" }]],
        rows: [[[{ text: "r1" }]]],
        alignments: ["left"],
      },
      {
        kind: "table",
        header: [[{ text: "B" }]],
        rows: [[[{ text: "r2" }]]],
        alignments: ["left"],
      },
    ]);
  });

  it("places an image inside a header cell", () => {
    // flattenInline is called on every inline token in a table, including
    // header cells, so images are handled identically to body-cell text.
    const src = "| ![logo](https://x/a.png) |\n| --- |\n| text |";
    expect(parseMarkdown(src)).toEqual([
      {
        kind: "table",
        header: [[{ src: "https://x/a.png", alt: "logo" }]],
        rows: [[[{ text: "text" }]]],
        alignments: ["left"],
      },
    ]);
  });

  it("places an image inside a table body cell", () => {
    // flattenInline is called for every inline token including body-cell
    // inline tokens, so images in body cells are handled the same way.
    const src = "| header |\n| --- |\n| ![logo](https://x/a.png) |";
    expect(parseMarkdown(src)).toEqual([
      {
        kind: "table",
        header: [[{ text: "header" }]],
        rows: [[[{ src: "https://x/a.png", alt: "logo" }]]],
        alignments: ["left"],
      },
    ]);
  });
});

describe("parseMarkdown — images", () => {
  it("parses a standalone image", () => {
    expect(parseMarkdown("![logo](https://x.com/a.png)")).toEqual([
      {
        kind: "paragraph",
        runs: [{ src: "https://x.com/a.png", alt: "logo" }],
      },
    ]);
  });

  it("captures the title attribute when present", () => {
    expect(parseMarkdown('![logo](https://x/a.png "tooltip")')).toEqual([
      {
        kind: "paragraph",
        runs: [
          {
            src: "https://x/a.png",
            alt: "logo",
            title: "tooltip",
          },
        ],
      },
    ]);
  });

  it("preserves an empty alt for purely decorative images", () => {
    expect(parseMarkdown("![](https://x/a.png)")).toEqual([
      {
        kind: "paragraph",
        runs: [{ src: "https://x/a.png", alt: "" }],
      },
    ]);
  });

  it("places images inline with surrounding text", () => {
    expect(parseMarkdown("see ![logo](https://x/a.png) here")).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "see " }, { src: "https://x/a.png", alt: "logo" }, { text: " here" }],
      },
    ]);
  });

  it("does not merge text runs across an image boundary", () => {
    // The image breaks the adjacency, so the trailing text stays its own
    // run even though it has the same formatting as the leading text.
    const out = parseMarkdown("a ![](u) b");
    expect(out).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "a " }, { src: "u", alt: "" }, { text: " b" }],
      },
    ]);
  });

  it("works inside list items", () => {
    expect(parseMarkdown("- ![logo](https://x/a.png)")).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ src: "https://x/a.png", alt: "logo" }],
      },
    ]);
  });

  it("captures the title attribute when the image is in a list item", () => {
    // Exercises the if (title) branch of flattenInline's image case inside
    // the list-item code path — previously only covered in plain paragraphs.
    expect(parseMarkdown('- ![logo](https://x/a.png "tooltip")')).toEqual([
      {
        kind: "listItem",
        ordered: false,
        depth: 0,
        listId: 1,
        runs: [{ src: "https://x/a.png", alt: "logo", title: "tooltip" }],
      },
    ]);
  });

  it("resolves a reference-style image like its inline counterpart", () => {
    const src = "![logo][ref]\n\n[ref]: https://x/a.png";
    expect(parseMarkdown(src)).toEqual([
      {
        kind: "paragraph",
        runs: [{ src: "https://x/a.png", alt: "logo" }],
      },
    ]);
  });

  it("drops the outer link when an image is wrapped in a link", () => {
    // [![alt](src)](href) — link_open sets s.link but Image has no link
    // field, so the href is silently discarded; only the image survives.
    expect(parseMarkdown("[![logo](https://x/img.png)](https://x)")).toEqual([
      {
        kind: "paragraph",
        runs: [{ src: "https://x/img.png", alt: "logo" }],
      },
    ]);
  });

  it("places an inline image inside a heading alongside text runs", () => {
    expect(parseMarkdown("# lead ![logo](https://x/icon.png) text")).toEqual([
      {
        kind: "heading",
        level: 1,
        runs: [{ text: "lead " }, { src: "https://x/icon.png", alt: "logo" }, { text: " text" }],
      },
    ]);
  });
});

describe("parseMarkdown — flavors of basic syntax", () => {
  it("handles Setext H1 (underlined with =) like ATX #", () => {
    expect(parseMarkdown("Hello\n=====")).toEqual([
      { kind: "heading", level: 1, runs: [{ text: "Hello" }] },
    ]);
  });

  it("handles Setext H2 (underlined with -) like ATX ##", () => {
    expect(parseMarkdown("Hello\n-----")).toEqual([
      { kind: "heading", level: 2, runs: [{ text: "Hello" }] },
    ]);
  });

  it("resolves reference-style links", () => {
    const src = "see [docs][1]\n\n[1]: https://example.com";
    expect(parseMarkdown(src)).toEqual([
      {
        kind: "paragraph",
        runs: [{ text: "see " }, { text: "docs", link: "https://example.com" }],
      },
    ]);
  });

  it("decodes named HTML entities", () => {
    expect(parseMarkdown("rights &copy; 2026")).toEqual([
      { kind: "paragraph", runs: [{ text: "rights © 2026" }] },
    ]);
  });

  it("decodes numeric HTML character references", () => {
    expect(parseMarkdown("snowman &#9731;")).toEqual([
      { kind: "paragraph", runs: [{ text: "snowman ☃" }] },
    ]);
  });
});
