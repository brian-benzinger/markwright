import { type Alignment, type Block, type Image, type Inline, type Run } from "../convert";

/**
 * Walks `Block[]` and writes the content into the active Word document
 * via the Office.js object model. Opens its own `Word.run` so callers
 * just `await applyBlocks(blocks)`.
 *
 * The first block lands in a fresh empty paragraph inserted before the
 * cursor; each subsequent block chains through `nextParagraph`, which
 * decides whether to reuse the previous paragraph (right after a table
 * or other anchor-reusing block) or insert a new one.
 */
export async function applyBlocks(blocks: Block[]): Promise<void> {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText("", Word.InsertLocation.replace);
    const state: RenderState = {
      currentList: null,
      currentListId: 0,
      configuredLevels: new Map(),
      paragraphIsEmpty: false,
    };
    let para = selection.insertParagraph("", Word.InsertLocation.before);
    applyBlock(para, blocks[0], state);
    for (let i = 1; i < blocks.length; i++) {
      para = nextParagraph(para, blocks[i], state);
      applyBlock(para, blocks[i], state);
    }
    await context.sync();
  });
}

type RenderState = {
  currentList: Word.List | null;
  currentListId: number;
  configuredLevels: Map<number, Set<number>>;
  // True when the active paragraph reference is empty and should be
  // reused for the next block. Set by table application — Word's
  // paragraph.insertTable can only insert Before/After, so we anchor
  // the table to an empty paragraph and reuse it for whatever follows.
  paragraphIsEmpty: boolean;
};

// Returns the paragraph the next block should land in, and updates the
// list state when we cross a list boundary.
function nextParagraph(prev: Word.Paragraph, block: Block, state: RenderState): Word.Paragraph {
  if (state.paragraphIsEmpty) {
    state.paragraphIsEmpty = false;
    return prev;
  }
  if (
    block.kind === "listItem" &&
    state.currentList !== null &&
    block.listId === state.currentListId
  ) {
    return state.currentList.insertParagraph("", Word.InsertLocation.end);
  }
  const next = prev.insertParagraph("", Word.InsertLocation.after);
  if (block.kind !== "listItem") {
    state.currentList = null;
    state.currentListId = 0;
  }
  return next;
}

function applyBlock(paragraph: Word.Paragraph, block: Block, state: RenderState): void {
  if (block.kind === "thematicBreak") {
    // Word interprets <hr/> as a bottom-bordered paragraph — the
    // conventional thematic-break representation. insertHtml on the
    // paragraph itself (Replace) avoids any styles-clobbering surprise
    // we hit with insertOoxml earlier.
    paragraph.insertHtml("<hr/>", Word.InsertLocation.replace);
    return;
  }
  if (block.kind === "codeBlock") {
    paragraph.styleBuiltIn = Word.BuiltInStyleName.normal;
    // markdown-it always emits a trailing newline; drop it. Remaining
    // newlines become in-paragraph line breaks so the whole snippet
    // lives in one Word paragraph and a stray Enter doesn't split it.
    const text = block.content.replace(/\n$/, "").replace(/\n/g, "\v");
    if (text === "") return;
    const range = paragraph.insertText(text, Word.InsertLocation.end);
    range.font.name = "Consolas";
    return;
  }
  if (block.kind === "table") {
    applyTable(paragraph, block, state);
    return;
  }
  if (block.kind === "listItem") {
    if (state.currentList === null || state.currentListId !== block.listId) {
      state.currentList = paragraph.startNewList();
      state.currentListId = block.listId;
    }
    configureListLevel(state, block.depth, block.ordered);
    paragraph.styleBuiltIn = Word.BuiltInStyleName.listParagraph;
    paragraph.listItem.level = block.depth;
    if (block.checked !== undefined) {
      // The list bullet still renders alongside the checkbox. Swapping
      // to a custom level bullet via setLevelBullet(custom, ...) would
      // be per-level and would break mixed task / non-task items in
      // the same list scope.
      paragraph.insertText(block.checked ? "☑ " : "☐ ", Word.InsertLocation.end);
    }
  } else if (block.kind === "paragraph" && block.quoteDepth) {
    paragraph.styleBuiltIn = Word.BuiltInStyleName.quote;
    // TODO: visually scale indent for quoteDepth > 1. Word's Quote style
    // sets its own left indent; layering an additive override needs a
    // load+sync of the style's defaults first, which we'd rather not pay
    // for the common single-depth case.
  } else {
    paragraph.styleBuiltIn =
      block.kind === "heading" ? headingStyle(block.level) : Word.BuiltInStyleName.normal;
  }
  for (const inline of block.runs) {
    applyInline(paragraph, inline);
  }
}

function applyTable(
  anchor: Word.Paragraph,
  block: Block & { kind: "table" },
  state: RenderState
): void {
  const colCount = block.alignments.length;
  const rowCount = 1 + block.rows.length;
  // Insert the table BEFORE our empty anchor paragraph so the anchor
  // survives below the table. nextParagraph reuses the anchor on the
  // following block (paragraphIsEmpty flag).
  const table = anchor.insertTable(rowCount, colCount, Word.InsertLocation.before);
  for (let c = 0; c < colCount; c++) {
    applyCell(table.getCell(0, c), block.header[c] ?? [], block.alignments[c], true);
  }
  for (let r = 0; r < block.rows.length; r++) {
    for (let c = 0; c < colCount; c++) {
      applyCell(table.getCell(r + 1, c), block.rows[r][c] ?? [], block.alignments[c], false);
    }
  }
  // A table breaks any active list scope; the anchor remains empty for
  // the next block to fill in.
  state.currentList = null;
  state.currentListId = 0;
  state.paragraphIsEmpty = true;
}

function applyCell(
  cell: Word.TableCell,
  inlines: Inline[],
  alignment: Alignment,
  isHeader: boolean
): void {
  cell.horizontalAlignment = alignToWord(alignment);
  const cellPara = cell.body.paragraphs.getFirst();
  for (const inline of inlines) {
    if ("src" in inline) {
      applyImage(cellPara, inline);
    } else {
      const range = cellPara.insertText(inline.text, Word.InsertLocation.end);
      formatRange(range, inline, isHeader);
    }
  }
}

function alignToWord(a: Alignment): Word.Alignment {
  switch (a) {
    case "left":
      return Word.Alignment.left;
    case "center":
      return Word.Alignment.centered;
    case "right":
      return Word.Alignment.right;
  }
}

function configureListLevel(state: RenderState, depth: number, ordered: boolean): void {
  if (state.currentList === null) return;
  let levels = state.configuredLevels.get(state.currentListId);
  if (!levels) {
    levels = new Set();
    state.configuredLevels.set(state.currentListId, levels);
  }
  if (levels.has(depth)) return;
  levels.add(depth);
  if (ordered) {
    state.currentList.setLevelNumbering(depth, Word.ListNumbering.arabic);
  } else {
    state.currentList.setLevelBullet(depth, Word.ListBullet.solid);
  }
}

function applyInline(paragraph: Word.Paragraph, inline: Inline): void {
  if ("src" in inline) {
    applyImage(paragraph, inline);
  } else {
    applyRun(paragraph, inline);
  }
}

function applyRun(paragraph: Word.Paragraph, run: Run): void {
  const range = paragraph.insertText(run.text, Word.InsertLocation.end);
  formatRange(range, run, false);
}

function applyImage(paragraph: Word.Paragraph, image: Image): void {
  // Word's HTML paste pipeline fetches the URL and embeds the image.
  // If the fetch fails (offline / CORS / 404) Word falls back to the
  // alt text, which matches a browser's <img> behavior.
  const html =
    `<img src="${escapeHtml(image.src)}" alt="${escapeHtml(image.alt)}"` +
    (image.title ? ` title="${escapeHtml(image.title)}"` : "") +
    " />";
  paragraph.insertHtml(html, Word.InsertLocation.end);
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function formatRange(range: Word.Range, run: Run, forceBold: boolean): void {
  if (run.bold || forceBold) range.font.bold = true;
  if (run.italic) range.font.italic = true;
  if (run.strike) range.font.strikeThrough = true;
  // Word has no built-in "code" character style; fall back to a monospace font.
  if (run.code) range.font.name = "Consolas";
  if (run.link) range.hyperlink = run.link;
}

function headingStyle(level: 1 | 2 | 3 | 4 | 5 | 6): Word.BuiltInStyleName {
  switch (level) {
    case 1:
      return Word.BuiltInStyleName.heading1;
    case 2:
      return Word.BuiltInStyleName.heading2;
    case 3:
      return Word.BuiltInStyleName.heading3;
    case 4:
      return Word.BuiltInStyleName.heading4;
    case 5:
      return Word.BuiltInStyleName.heading5;
    case 6:
      return Word.BuiltInStyleName.heading6;
  }
}
