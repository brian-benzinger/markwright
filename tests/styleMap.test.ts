import { describe, expect, it } from "vitest";
import {
  defaultStyleMap,
  mergeStyleMap,
  STYLE_TARGETS,
  type StyleMap,
} from "../src/convert/styleMap";

describe("STYLE_TARGETS", () => {
  it("contains exactly the 9 mappable block constructs in display order", () => {
    // The UI iterates STYLE_TARGETS to render the mapping rows, so order
    // matters: headings first (1..6), then paragraph, blockquote, codeBlock.
    expect(STYLE_TARGETS).toEqual([
      "heading1",
      "heading2",
      "heading3",
      "heading4",
      "heading5",
      "heading6",
      "paragraph",
      "blockquote",
      "codeBlock",
    ]);
  });
});

describe("defaultStyleMap", () => {
  it("binds every target, headings to Heading 1..6 and the rest to Normal/Quote", () => {
    const map = defaultStyleMap();
    for (const target of STYLE_TARGETS) {
      expect(map[target]).toBeDefined();
    }
    expect(map.heading1).toEqual({ builtIn: "heading1" });
    expect(map.heading6).toEqual({ builtIn: "heading6" });
    expect(map.paragraph).toEqual({ builtIn: "normal" });
    expect(map.codeBlock).toEqual({ builtIn: "normal" });
    expect(map.blockquote).toEqual({ builtIn: "quote" });
  });

  it("returns a fresh object each call (no shared mutable state)", () => {
    const a = defaultStyleMap();
    const b = defaultStyleMap();
    expect(a).not.toBe(b);
    a.paragraph = { custom: "Body Text" };
    expect(b.paragraph).toEqual({ builtIn: "normal" });
  });
});

describe("mergeStyleMap", () => {
  it("returns defaults for non-object input", () => {
    const def = defaultStyleMap();
    expect(mergeStyleMap(null)).toEqual(def);
    expect(mergeStyleMap(undefined)).toEqual(def);
    expect(mergeStyleMap("garbage")).toEqual(def);
    expect(mergeStyleMap(42)).toEqual(def);
  });

  it("overlays a valid custom choice and keeps defaults for the rest", () => {
    const map = mergeStyleMap({ heading1: { custom: "Title" } });
    expect(map.heading1).toEqual({ custom: "Title" });
    expect(map.paragraph).toEqual({ builtIn: "normal" });
  });

  it("accepts a valid built-in token override", () => {
    const map = mergeStyleMap({ codeBlock: { builtIn: "quote" } });
    expect(map.codeBlock).toEqual({ builtIn: "quote" });
  });

  it("falls back to default for unknown built-in tokens", () => {
    const map = mergeStyleMap({ heading2: { builtIn: "bogus" } });
    expect(map.heading2).toEqual({ builtIn: "heading2" });
  });

  it("falls back to default for empty or non-string custom names", () => {
    expect(mergeStyleMap({ paragraph: { custom: "" } }).paragraph).toEqual({ builtIn: "normal" });
    expect(mergeStyleMap({ paragraph: { custom: 5 } }).paragraph).toEqual({ builtIn: "normal" });
  });

  it("falls back to default for malformed entries (no builtIn/custom key, null, non-object)", () => {
    const def = defaultStyleMap();
    expect(mergeStyleMap({ blockquote: {} }).blockquote).toEqual(def.blockquote);
    expect(mergeStyleMap({ blockquote: null }).blockquote).toEqual(def.blockquote);
    expect(mergeStyleMap({ blockquote: "Quote" }).blockquote).toEqual(def.blockquote);
    expect(mergeStyleMap({ blockquote: { other: 1 } }).blockquote).toEqual(def.blockquote);
  });

  it("ignores keys that are not style targets", () => {
    const map: StyleMap = mergeStyleMap({ notATarget: { custom: "X" } });
    expect(map).toEqual(defaultStyleMap());
  });

  it("prefers builtIn over custom when both keys are present in the same entry", () => {
    // asStyleChoice checks "builtIn" in value first, so an entry carrying both
    // keys resolves as builtIn rather than custom. This matches Settings-store
    // objects that may accumulate stale keys across schema changes.
    const map = mergeStyleMap({ paragraph: { builtIn: "quote", custom: "My Style" } });
    expect(map.paragraph).toEqual({ builtIn: "quote" });
  });

  it("does not mutate the persisted argument", () => {
    const persisted = { paragraph: { builtIn: "quote" } };
    const snapshot = JSON.parse(JSON.stringify(persisted)) as unknown;
    mergeStyleMap(persisted);
    expect(persisted).toEqual(snapshot);
  });

  it("round-trips a fully customised map through JSON", () => {
    const original: StyleMap = {
      heading1: { custom: "Title" },
      heading2: { custom: "Heading A" },
      heading3: { builtIn: "heading3" },
      heading4: { builtIn: "heading4" },
      heading5: { builtIn: "heading5" },
      heading6: { builtIn: "heading6" },
      paragraph: { custom: "Body Text" },
      blockquote: { custom: "Intense Quote" },
      codeBlock: { custom: "HTML Preformatted" },
    };
    const restored = mergeStyleMap(JSON.parse(JSON.stringify(original)));
    expect(restored).toEqual(original);
  });
});
