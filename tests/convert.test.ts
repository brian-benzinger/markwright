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
    const out = parseMarkdown("- a\n\n1. b");
    expect(out.map((b) => "listId" in b && b.listId)).toEqual([1, 2]);
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
    const out = parseMarkdown("- top\n  - nested");
    const ids = out.map((b) => "listId" in b && b.listId);
    expect(ids).toEqual([1, 1]);
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

  it("interleaves headings and lists without dropping order", () => {
    const out = parseMarkdown("# H\n\n- a\n\np\n\n1. one");
    expect(out.map((b) => b.kind)).toEqual(["heading", "listItem", "paragraph", "listItem"]);
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

  it("flattens list items inside a blockquote to quoted paragraphs (lossy)", () => {
    expect(parseMarkdown("> - one\n> - two")).toEqual([
      { kind: "paragraph", runs: [{ text: "one" }], quoteDepth: 1 },
      { kind: "paragraph", runs: [{ text: "two" }], quoteDepth: 1 },
    ]);
  });
});
