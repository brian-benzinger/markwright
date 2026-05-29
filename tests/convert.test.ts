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
        runs: [
          { text: "call " },
          { text: "foo()", code: true },
          { text: " now" },
        ],
      },
    ]);
  });

  it("parses a link with URL", () => {
    expect(parseMarkdown("see [docs](https://example.com)")).toEqual([
      {
        kind: "paragraph",
        runs: [
          { text: "see " },
          { text: "docs", link: "https://example.com" },
        ],
      },
    ]);
  });

  it("parses an autolink as a link with text == href", () => {
    expect(parseMarkdown("at <https://example.com>")).toEqual([
      {
        kind: "paragraph",
        runs: [
          { text: "at " },
          { text: "https://example.com", link: "https://example.com" },
        ],
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
        runs: [
          { text: "a " },
          { text: "bold", bold: true },
          { text: " item" },
        ],
      },
    ]);
  });

  it("interleaves headings and lists without dropping order", () => {
    const out = parseMarkdown("# H\n\n- a\n\np\n\n1. one");
    expect(out.map((b) => b.kind)).toEqual([
      "heading",
      "listItem",
      "paragraph",
      "listItem",
    ]);
  });
});

describe("parseMarkdown — blockquotes (M3 placeholder)", () => {
  // Blockquotes aren't implemented yet; until they are, their contents
  // are dropped rather than emitted as bare paragraphs that lose the
  // quote semantics. These tests pin that behavior so the gap is visible.
  it("suppresses contents nested inside a blockquote", () => {
    expect(parseMarkdown("> quoted line")).toEqual([]);
  });

  it("suppresses blockquote contents but keeps surrounding blocks in order", () => {
    expect(parseMarkdown("before\n\n> quoted\n\nafter")).toEqual([
      { kind: "paragraph", runs: [{ text: "before" }] },
      { kind: "paragraph", runs: [{ text: "after" }] },
    ]);
  });

  it("suppresses contents of nested blockquotes", () => {
    expect(parseMarkdown("> outer\n>\n> > inner\n\nresumed")).toEqual([
      { kind: "paragraph", runs: [{ text: "resumed" }] },
    ]);
  });
});
