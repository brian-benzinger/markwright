import { describe, it, expect } from "vitest";
import { markdownToOoxml } from "../src/convert";

describe("markdownToOoxml", () => {
  describe("Flat OPC envelope", () => {
    it("emits the package wrapper and mso processing instruction", () => {
      const out = markdownToOoxml("hello");
      expect(out.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')).toBe(true);
      expect(out).toContain('<?mso-application progid="Word.Document"?>');
      expect(out).toContain("<pkg:package");
      expect(out.endsWith("</pkg:package>")).toBe(true);
    });

    it("declares the four required parts", () => {
      const out = markdownToOoxml("hello");
      const parts = [
        "/_rels/.rels",
        "/word/_rels/document.xml.rels",
        "/word/document.xml",
        "/word/styles.xml",
      ];
      for (const p of parts) {
        expect(out).toContain(`pkg:name="${p}"`);
      }
    });

    it("links document.xml to styles.xml via a relationship", () => {
      const out = markdownToOoxml("hello");
      expect(out).toMatch(
        /<Relationship[^>]+Type="[^"]*\/relationships\/styles"[^>]+Target="styles\.xml"/,
      );
    });

    it("declares Normal as default and Heading1-6 in the styles part", () => {
      const out = markdownToOoxml("");
      expect(out).toMatch(/w:default="1" w:styleId="Normal"/);
      for (const n of [1, 2, 3, 4, 5, 6]) {
        expect(out).toContain(`w:styleId="Heading${n}"`);
      }
    });
  });

  describe("paragraphs", () => {
    it("renders a plain paragraph without a style reference", () => {
      const out = markdownToOoxml("hello world");
      expect(out).toContain(
        '<w:p><w:r><w:t xml:space="preserve">hello world</w:t></w:r></w:p>',
      );
    });

    it("emits an empty body for empty input", () => {
      const out = markdownToOoxml("");
      expect(out).toContain("<w:body></w:body>");
    });

    it("preserves multiple block-level elements in document order", () => {
      const out = markdownToOoxml("# H\n\np1\n\n## S\n\np2");
      const body = bodyOf(out);
      const order = [
        body.indexOf('w:val="Heading1"'),
        body.indexOf(">H<"),
        body.indexOf(">p1<"),
        body.indexOf('w:val="Heading2"'),
        body.indexOf(">S<"),
        body.indexOf(">p2<"),
      ];
      expect(order.every((i) => i >= 0)).toBe(true);
      for (let i = 1; i < order.length; i++) {
        expect(order[i]).toBeGreaterThan(order[i - 1]);
      }
    });
  });

  describe("headings", () => {
    it.each([
      ["#", "Heading1"],
      ["##", "Heading2"],
      ["###", "Heading3"],
      ["####", "Heading4"],
      ["#####", "Heading5"],
      ["######", "Heading6"],
    ])("maps %s to %s", (hashes, styleId) => {
      const out = markdownToOoxml(`${hashes} title`);
      expect(bodyOf(out)).toContain(
        `<w:p><w:pPr><w:pStyle w:val="${styleId}"/></w:pPr><w:r><w:t xml:space="preserve">title</w:t></w:r></w:p>`,
      );
    });
  });

  describe("escaping", () => {
    it("escapes & < > in text content", () => {
      const out = markdownToOoxml("a & b < c > d");
      expect(bodyOf(out)).toContain("a &amp; b &lt; c &gt; d");
    });

    it("does not leave a naked ampersand in the body", () => {
      const body = bodyOf(markdownToOoxml("Q&A and 1 < 2"));
      // Every & in body text must be followed by an entity name.
      const naked = body.match(/&(?!amp;|lt;|gt;|quot;|apos;|#\d+;)/g);
      expect(naked).toBeNull();
    });
  });
});

function bodyOf(ooxml: string): string {
  const m = ooxml.match(/<w:body>([\s\S]*)<\/w:body>/);
  if (!m) throw new Error("no body found in OOXML");
  return m[1];
}
