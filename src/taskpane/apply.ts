import { type Alignment, type Block, type Image, type Inline, type Run } from "../convert";
import {
  defaultStyleMap,
  type StyleChoice,
  type StyleMap,
  type StyleToken,
} from "../convert/styleMap";

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
export async function applyBlocks(
  blocks: Block[],
  styleMap: StyleMap = defaultStyleMap()
): Promise<void> {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    // Capture the insertion point's default font name with a single sync
    // up front. We need it to reset the font after an inline-code run:
    // insertText("…", End) inherits the previous character's font, so
    // `code` text` would leave `text` in Consolas for the rest of the
    // paragraph. This is the one round-trip the formatRange comment was
    // unwilling to pay PER RUN — paid once here it's cheap.
    selection.load("font/name");
    await context.sync();
    const defaultFontName = selection.font.name;
    selection.insertText("", Word.InsertLocation.replace);
    const state: RenderState = {
      currentList: null,
      currentListId: 0,
      configuredLevels: new Map(),
      paragraphIsEmpty: false,
      defaultFontName,
      styleMap,
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
  // The insertion point's default font name, captured once up front.
  // Used to reset a run's font after an inline-code run so Consolas
  // doesn't bleed into the rest of the paragraph. May be empty if Word
  // reported no single font (e.g. a mixed-font selection); runFontName
  // then skips the reset rather than clearing the font.
  defaultFontName: string;
  // The user-configured (or default) mapping from Markdown constructs to
  // Word paragraph styles. setParagraphStyle reads it per block.
  styleMap: StyleMap;
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
    setParagraphStyle(paragraph, state.styleMap.codeBlock);
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
    // Don't set styleBuiltIn = listParagraph here. The list-level config
    // (setLevelBullet / setLevelNumbering) plus Word's default list
    // indent already produce a bulleted/numbered item; setting the
    // List Paragraph style on TOP of that double-indents the first item
    // and — worse — detaches subsequent items from the list (Office.js
    // styleBuiltIn assignment can clear list membership).
    if (state.currentList === null || state.currentListId !== block.listId) {
      // The first item was created by prev.insertParagraph(After), so it
      // inherited the previous block's style — Heading 3 when the list
      // follows a heading. Reset to Normal BEFORE startNewList so the
      // item (and the rest of the list/page) isn't heading-styled. Doing
      // it before list attachment sidesteps the styleBuiltIn-clears-list-
      // membership problem noted above; continuation items come from the
      // list itself and are already Normal, so they're left untouched.
      paragraph.styleBuiltIn = Word.BuiltInStyleName.normal;
      state.currentList = paragraph.startNewList();
      state.currentListId = block.listId;
    }
    configureListLevel(state, block.depth, block.ordered);
    if (block.depth > 0) {
      // depth 0 is the default after startNewList; setting it can race
      // with the queued attachment and throw ItemNotFound.
      paragraph.listItem.level = block.depth;
    }
    if (block.checked !== undefined) {
      // The list bullet still renders alongside the checkbox. Swapping
      // to a custom level bullet via setLevelBullet(custom, ...) would
      // be per-level and would break mixed task / non-task items in
      // the same list scope.
      paragraph.insertText(block.checked ? "☑ " : "☐ ", Word.InsertLocation.end);
    }
  } else if (block.kind === "paragraph" && block.quoteDepth) {
    setParagraphStyle(paragraph, state.styleMap.blockquote);
    // Word's default Quote style is centered in most themes, which
    // fights the Markdown convention of left-aligned indented quote
    // text. Force left alignment only while the blockquote mapping is
    // left at its default (built-in Quote); once the user remaps it to a
    // style of their own, respect that style's alignment.
    if ("builtIn" in state.styleMap.blockquote) paragraph.alignment = Word.Alignment.left;
    // TODO: visually scale indent for quoteDepth > 1. Word's Quote style
    // sets its own left indent; layering an additive override needs a
    // load+sync of the style's defaults first, which we'd rather not pay
    // for the common single-depth case.
  } else {
    const choice =
      block.kind === "heading"
        ? state.styleMap[headingTarget(block.level)]
        : state.styleMap.paragraph;
    setParagraphStyle(paragraph, choice);
  }
  // Track whether the previous run was inline code so the next run can
  // reset its font (Consolas would otherwise bleed forward — see
  // runFontName / formatRange).
  let prevWasCode = false;
  for (const inline of block.runs) {
    prevWasCode = applyInline(paragraph, inline, prevWasCode, state.defaultFontName);
  }
}

function applyTable(
  anchor: Word.Paragraph,
  block: Block & { kind: "table" },
  state: RenderState
): void {
  const colCount = block.alignments.length;
  const rowCount = 1 + block.rows.length;
  // The anchor was created by `prev.insertParagraph(After)`, so it
  // inherited the previous block's style. When a table follows a
  // heading, that means the anchor is Heading 3 (etc.). Reset it to
  // Normal: if the table is the last block, this trailing empty
  // paragraph — and anything the user types after it — must not stay
  // stuck in the heading style. If another block reuses the anchor
  // (paragraphIsEmpty), applyBlock restyles it anyway, so this is safe.
  anchor.styleBuiltIn = Word.BuiltInStyleName.normal;
  // Insert the table BEFORE our empty anchor paragraph so the anchor
  // survives below the table. nextParagraph reuses the anchor on the
  // following block (paragraphIsEmpty flag).
  const table = anchor.insertTable(rowCount, colCount, Word.InsertLocation.before);
  for (let c = 0; c < colCount; c++) {
    applyCell(
      table.getCell(0, c),
      block.header[c] ?? [],
      block.alignments[c],
      true,
      state.defaultFontName
    );
  }
  for (let r = 0; r < block.rows.length; r++) {
    for (let c = 0; c < colCount; c++) {
      applyCell(
        table.getCell(r + 1, c),
        block.rows[r][c] ?? [],
        block.alignments[c],
        false,
        state.defaultFontName
      );
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
  isHeader: boolean,
  defaultFontName: string
): void {
  cell.horizontalAlignment = alignToWord(alignment);
  const cellPara = cell.body.paragraphs.getFirst();
  let prevWasCode = false;
  for (const inline of inlines) {
    if ("src" in inline) {
      applyImage(cellPara, inline);
      prevWasCode = false;
    } else {
      const range = cellPara.insertText(inline.text, Word.InsertLocation.end);
      formatRange(range, inline, isHeader, runFontName(inline, prevWasCode, defaultFontName));
      prevWasCode = !!inline.code;
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
    // Without a format string, Word renders a bare "1" with no trailing
    // separator. The format array places the level's number (the integer
    // placeholder for this level) followed by a literal ".", yielding
    // "1." / "2." as Markdown ordered lists expect.
    state.currentList.setLevelNumbering(depth, Word.ListNumbering.arabic, [depth, "."]);
  } else {
    state.currentList.setLevelBullet(depth, Word.ListBullet.solid);
  }
}

// Renders one inline and returns whether it was an inline-code run, so
// the caller can tell the next run to reset its font.
function applyInline(
  paragraph: Word.Paragraph,
  inline: Inline,
  prevWasCode: boolean,
  defaultFontName: string
): boolean {
  if ("src" in inline) {
    applyImage(paragraph, inline);
    return false;
  }
  applyRun(paragraph, inline, prevWasCode, defaultFontName);
  return !!inline.code;
}

function applyRun(
  paragraph: Word.Paragraph,
  run: Run,
  prevWasCode: boolean,
  defaultFontName: string
): void {
  const range = paragraph.insertText(run.text, Word.InsertLocation.end);
  formatRange(range, run, false, runFontName(run, prevWasCode, defaultFontName));
}

// Decides the font name to force on a run, or undefined to leave it
// alone. Code runs get Consolas. A non-code run that FOLLOWS a code run
// is reset to the document default so Consolas doesn't bleed forward;
// every other non-code run is left untouched so heading/quote fonts
// survive. The reset is skipped when no default font name is known.
function runFontName(run: Run, prevWasCode: boolean, defaultFontName: string): string | undefined {
  if (run.code) return "Consolas";
  if (prevWasCode && defaultFontName) return defaultFontName;
  return undefined;
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

function formatRange(
  range: Word.Range,
  run: Run,
  forceBold: boolean,
  fontName: string | undefined
): void {
  // Always assign the toggleable marks explicitly. insertText("…", End)
  // makes the new range inherit the previous character's formatting in
  // Word, so a `**bold** plain` paragraph would leave `plain` bold
  // unless we set bold = false explicitly here.
  range.font.bold = !!run.bold || forceBold;
  range.font.italic = !!run.italic;
  range.font.strikeThrough = !!run.strike;
  // Word has no built-in "code" character style; code runs fall back to a
  // monospace font (fontName === "Consolas"). The same inheritance that
  // bleeds bold also bleeds that font into the next run, so runFontName
  // hands us the document default to reset it; undefined means leave the
  // font alone (preserving the paragraph style's own font).
  if (fontName !== undefined) range.font.name = fontName;
  if (run.link) range.hyperlink = run.link;
}

// Applies a configured StyleChoice to a paragraph. Built-in tokens go
// through styleBuiltIn (locale-invariant); custom choices carry a host
// style's localised name read live from getStyles(), applied via
// paragraph.style — valid on the install it was chosen on.
function setParagraphStyle(paragraph: Word.Paragraph, choice: StyleChoice): void {
  if ("builtIn" in choice) {
    paragraph.styleBuiltIn = builtInStyle(choice.builtIn);
  } else {
    paragraph.style = choice.custom;
  }
}

function builtInStyle(token: StyleToken): Word.BuiltInStyleName {
  switch (token) {
    case "heading1":
      return Word.BuiltInStyleName.heading1;
    case "heading2":
      return Word.BuiltInStyleName.heading2;
    case "heading3":
      return Word.BuiltInStyleName.heading3;
    case "heading4":
      return Word.BuiltInStyleName.heading4;
    case "heading5":
      return Word.BuiltInStyleName.heading5;
    case "heading6":
      return Word.BuiltInStyleName.heading6;
    case "quote":
      return Word.BuiltInStyleName.quote;
    case "normal":
      return Word.BuiltInStyleName.normal;
  }
}

function headingTarget(
  level: 1 | 2 | 3 | 4 | 5 | 6
): "heading1" | "heading2" | "heading3" | "heading4" | "heading5" | "heading6" {
  return `heading${level}` as const;
}
