import { describe, it, expect } from "vitest";
import { parseMarkdown } from "../src/convert";

describe("parseMarkdown", () => {
  it("returns an empty array for empty input", () => {
    expect(parseMarkdown("")).toEqual([]);
  });

  it("returns an empty array for whitespace-only input", () => {
    expect(parseMarkdown("   \n  \n")).toEqual([]);
  });

  it("parses a single paragraph", () => {
    expect(parseMarkdown("hello world")).toEqual([
      { kind: "paragraph", text: "hello world" },
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
      { kind: "heading", level, text: "heading text" },
    ]);
  });

  it("preserves block order across mixed input", () => {
    const out = parseMarkdown("# H\n\np1\n\n## S\n\np2");
    expect(out).toEqual([
      { kind: "heading", level: 1, text: "H" },
      { kind: "paragraph", text: "p1" },
      { kind: "heading", level: 2, text: "S" },
      { kind: "paragraph", text: "p2" },
    ]);
  });

  it("preserves XML special characters verbatim (Office.js handles escaping)", () => {
    expect(parseMarkdown("a & b < c > d")).toEqual([
      { kind: "paragraph", text: "a & b < c > d" },
    ]);
  });

  it("ignores block types not yet supported in M3 (lists, code, blockquote)", () => {
    // Strict pre-M3 scope: lists/code/blockquote are silently skipped for now.
    // Adding coverage for these is the next milestone and will update this test.
    const out = parseMarkdown("- one\n- two\n\n> quote\n\n```\ncode\n```");
    expect(out).toEqual([]);
  });
});
