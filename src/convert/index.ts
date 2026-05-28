import MarkdownIt from "markdown-it";
import type Token from "markdown-it/lib/token.mjs";

const md = new MarkdownIt({ html: false, linkify: false, typographer: false });

const HEADING_STYLE: Record<string, string> = {
  h1: "Heading1",
  h2: "Heading2",
  h3: "Heading3",
  h4: "Heading4",
  h5: "Heading5",
  h6: "Heading6",
};

export function markdownToOoxml(source: string): string {
  const tokens = md.parse(source, {});
  const body = renderBlocks(tokens);
  return wrapFlatOpc(body);
}

function renderBlocks(tokens: Token[]): string {
  const out: string[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.type === "heading_open") {
      const style = HEADING_STYLE[t.tag] ?? "Heading1";
      const inline = tokens[i + 1];
      const text = inline?.type === "inline" ? renderInline(inline) : "";
      out.push(paragraph(text, style));
      // Skip the inline and the matching heading_close.
      i += 2;
    } else if (t.type === "paragraph_open") {
      const inline = tokens[i + 1];
      const text = inline?.type === "inline" ? renderInline(inline) : "";
      out.push(paragraph(text));
      i += 2;
    }
    // Strict M2: everything else is ignored. Lists, code, emphasis,
    // tables, etc. land in Milestone 3.
  }
  return out.join("");
}

function renderInline(token: Token): string {
  // Strict M2: emit only the text content. Marks (em, strong, code, link)
  // are added in Milestone 3.
  const children = token.children ?? [];
  return children
    .filter((c) => c.type === "text")
    .map((c) => escapeXml(c.content))
    .join("");
}

function paragraph(textXml: string, styleId?: string): string {
  const pPr = styleId ? `<w:pPr><w:pStyle w:val="${escapeAttr(styleId)}"/></w:pPr>` : "";
  const run = textXml ? `<w:r><w:t xml:space="preserve">${textXml}</w:t></w:r>` : "";
  return `<w:p>${pPr}${run}</w:p>`;
}

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeAttr(s: string): string {
  return escapeXml(s).replace(/"/g, "&quot;");
}

function wrapFlatOpc(bodyXml: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<?mso-application progid="Word.Document"?>` +
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    rootRelsPart() +
    documentRelsPart() +
    documentPart(bodyXml) +
    stylesPart() +
    `</pkg:package>`
  );
}

function rootRelsPart(): string {
  return (
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">` +
    `<pkg:xmlData>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships>` +
    `</pkg:xmlData>` +
    `</pkg:part>`
  );
}

function documentRelsPart(): string {
  return (
    `<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="256">` +
    `<pkg:xmlData>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    `</Relationships>` +
    `</pkg:xmlData>` +
    `</pkg:part>`
  );
}

function documentPart(bodyXml: string): string {
  return (
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">` +
    `<pkg:xmlData>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${bodyXml}</w:body>` +
    `</w:document>` +
    `</pkg:xmlData>` +
    `</pkg:part>`
  );
}

function stylesPart(): string {
  // Declares the styleIds we reference in document.xml. Word merges these
  // with the host document's styles by ID — for built-in IDs like
  // "Heading1", the host's appearance wins, so we get the document's own
  // heading styling rather than a Markwright default.
  const headings = [1, 2, 3, 4, 5, 6]
    .map(
      (n) =>
        `<w:style w:type="paragraph" w:styleId="Heading${n}">` +
        `<w:name w:val="heading ${n}"/>` +
        `<w:basedOn w:val="Normal"/>` +
        `<w:next w:val="Normal"/>` +
        `<w:uiPriority w:val="9"/>` +
        `<w:qFormat/>` +
        `</w:style>`
    )
    .join("");
  const stylesXml =
    `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">` +
    `<w:name w:val="Normal"/>` +
    `<w:qFormat/>` +
    `</w:style>` +
    headings +
    `</w:styles>`;
  return (
    `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml">` +
    `<pkg:xmlData>` +
    stylesXml +
    `</pkg:xmlData>` +
    `</pkg:part>`
  );
}
