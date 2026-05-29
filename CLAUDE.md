# CLAUDE.md — Markwright

Notes for future Claude sessions. Things that aren't obvious from the
code or that bit me in earlier sessions.

## Architecture seam

The load-bearing abstraction is `Block[]` in `src/convert/index.ts`.
Parsing markdown is one concern; applying it to Word is another.
Adding a new block type means: extend `Block` (and `Run` if it's
inline-level), teach `parseMarkdown` to emit it, teach `applyBlock`
in `src/taskpane/apply.ts` to render it. The task pane UI in
`src/taskpane/taskpane.ts` is just DOM glue that calls
`applyBlocks(blocks)`. Tests live in `tests/convert.test.ts` and
stay pure — they never touch Office.js.

A few patterns that recur in the applier:

- `nextParagraph(prev, block, state)` decides where the next paragraph
  lands and updates list state on boundary crossings. The main loop
  stays a two-liner: `para = nextParagraph(...); applyBlock(...)`.
- Inline state inside `flattenInline` is a single `InlineState` struct
  (bold/italic/strike depths + current link). `makeRun(text, state, code?)`
  is the only constructor — if you find yourself passing 4+ positional
  flags, rebuild it on top of the struct.
- Inline content is `Inline[] = (Run | Image)[]`. The field name is
  `runs` for backward compatibility but the type widened when images
  shipped. The applier discriminates structurally with `"src" in
  inline`; `pushRun` only merges adjacent Runs, never coalesces across
  an image. New inline kinds should follow the same `"discriminator"
  in inline` pattern rather than adding a `kind` tag.

## Office.js gotchas

- `Range.insertParagraph` accepts only `Before`/`After`/`Start`/`End`,
  never `Replace`. To replace a selection, call
  `range.insertText("", Replace)` first, then
  `selection.insertParagraph("", Before)`.
- Use `paragraph.styleBuiltIn`, not `paragraph.style`. The `style`
  string is the localised name and fails in non-English Word installs.
  `Word.BuiltInStyleName.heading1` etc. are locale-invariant.
- `BuiltInStyleName` has no code character style. Inline code falls
  back to `range.font.name = "Consolas"`.
- Hard line breaks inside a paragraph are `"\v"` (U+000B) in the
  inserted text — Word renders that as a soft return.
- Lists need `paragraph.startNewList()` + `list.insertParagraph("", End)`
  for continuation. Chaining via `paragraph.insertParagraph(After)`
  makes separate paragraphs, not list continuation. Memoise
  `setLevelBullet`/`setLevelNumbering` per `(list, level)` — calling
  them repeatedly resets Arabic numbering.
- `paragraph.insertTable(rows, cols, location)` accepts only
  Before/After (no Replace, like other paragraph inserts). We anchor
  to an empty paragraph and reuse it for the next block via
  `state.paragraphIsEmpty`. Cell contents live in `cell.body` (a full
  `Word.Body`); reach the writable paragraph via
  `cell.body.paragraphs.getFirst()` then `insertText` into it.
- `Word.Alignment.centered` (not `center`). The setter on
  `cell.horizontalAlignment` accepts the enum or the matching string
  ("Left"/"Centered"/"Right"/"Justified").

## Why the Block AST exists (don't redo the OOXML pivot)

PR #5 abandoned `Range.insertOoxml` with a Flat OPC fragment. Word
treats the package's minimal style declarations as the authoritative
definitions of Heading1–6, flattening the output silently — the
paragraphs *are* "Heading1", just stripped of their formatting. The
object-model path binds to the host doc's real styles. Don't reach
for `insertOoxml` again unless you're shipping tables/images/footnotes
(where it remains the right tool — a stretch milestone).

## markdown-it specifics

- Default config: `MarkdownIt({ html: false, linkify: false, typographer: false })`.
- `softbreak` → single space. `hardbreak` → `"\v"`.
- `~~strike~~` parses as `s_open`/`s_close` by default, no plugin.
- `code_inline` is atomic (no open/close); content lives on the token.
- Autolinks `<https://x>` arrive as ordinary `link_open` with
  `markup: "autolink"`. Bare-URL autolinks (`linkify: true`) produce
  the same shape — flattenInline handles them without extra code.
- GFM task markers are not a markdown-it feature in core. We detect
  `[ ] ` / `[x] ` / `[X] ` on the FIRST text child of a list-item
  paragraph and strip it before flattening; no plugin required. If you
  add `markdown-it-task-lists` later the prefix would already be gone
  before it ran, so don't.
- Container tokens (`bullet_list_open`, `ordered_list_open`,
  `blockquote_open`, `list_item_open`) wrap inner `paragraph_open`
  tokens. The parser tracks depth and `listStack` to decide whether
  to emit a `paragraph`, `listItem`, or skip entirely.

## Workflow

- One feature per branch off `main`. Branch naming:
  `claude/markwright-<slug>`.
- Run all four gates locally before push:
  `npm run typecheck && npm run lint && npm test && npm run build`.
  CI runs the same set.
- `npm run lint:fix` cleans up Prettier nits and auto-fixable lint.
- Lint pipeline is direct: `eslint` (with `eslint-plugin-office-addins`
  for Office.js-specific rules + `@typescript-eslint` recommended +
  our strict additions: `no-explicit-any`, `consistent-type-imports`
  inline-style, `no-non-null-assertion`, `no-inferrable-types`) and
  Prettier with `.prettierrc.json`. The `office-addin-lint` wrapper
  was removed for explicit version control. Don't add format rules to
  ESLint — Prettier owns formatting.
- Sourcemaps are dev-only (`devtool: dev ? "source-map" : false` in
  `webpack.config.js`). Production builds ship code only.
- No core-js polyfill. Word desktop (WebView2) and Word web both
  speak ES2020+ which matches our tsconfig target.
- No `npm start` auto-sideload — sideload `dist/manifest.xml` once via
  Word's Insert > My Add-ins > Upload. Dropping `office-addin-debugging`
  saved ~290 MB of Azure SDK transitives.
- Coverage gate runs in `npm test` via `vitest run --coverage`. Only
  `src/convert/` is in scope — the Office.js applier can't execute
  under vitest because the `Word` global only exists in the host.
  Thresholds in `vitest.config.ts`: lines 90, statements 90,
  functions 100, branches 90.

## Outstanding work in M3 / M4

- Visual indent scaling for nested blockquotes (`quoteDepth > 1`). The
  Quote style sets its own left indent; layering an additive override
  for deeper levels needs a `load()` + `sync()` of the style's defaults
  before each Word.run iteration, which we didn't want to pay for the
  common single-depth case. Revisit if real docs need it.
- Code blocks, thematic breaks, and tables inside blockquotes:
  currently dropped. The inside-blockquote branch only handles
  `paragraph_open` and `heading_open`; everything else falls through
  to `continue`. Rare in practice — add when needed.
- Tables inside lists: dropped (the `table_open` branch is gated on
  `listStack.length === 0`). GFM doesn't formally allow this anyway.

Blockquote semantics are deliberately lossy: a heading or list item
inside a blockquote becomes a flat quote-styled paragraph. This keeps
the AST simple and matches conventional Markdown rendering. If a real
use case turns up that needs styled headings inside quotes, the AST
seam can grow a `quoteDepth?` field on `heading` and `listItem` too —
but don't preempt that.

Code-block rendering choices:
- markdown-it always ends `fence`/`code_block` `content` with `\n`. The
  applier strips that trailing newline before insertion.
- Internal `\n`s are converted to `\v` so the snippet stays in one
  Word paragraph (one block visually, no stray Enter splits it).
- Word has no built-in "Code" paragraph style. `styleBuiltIn = normal`
  + `range.font.name = "Consolas"` is the simplest path that binds to
  a font the host actually ships.

Thematic-break rendering:
- `paragraph.insertHtml("<hr/>", Replace)` is the simplest reliable
  path. Word interprets `<hr/>` as a bottom-bordered paragraph. The
  borders API (`paragraph.borders.items`) would require iterating to
  find Bottom and setting type/color/width by hand — more API surface
  for no visual win.
- `insertHtml` doesn't trip the OOXML-styles-clobbering problem the
  way `insertOoxml` does; it's character/range-level, not package-level.

Table rendering:
- `applyTable` inserts the Word table BEFORE the empty anchor
  paragraph that the loop already created, then sets
  `state.paragraphIsEmpty = true` so `nextParagraph` reuses the anchor
  for whatever follows. No extra blank lines around tables.
- Per-cell formatting: alignment is set once at the cell level
  (`cell.horizontalAlignment`); inline marks reuse `formatRange` which
  is shared with the non-table run path.
- Header bolding is forced via `formatRange(range, run, isHeader)` —
  one extra bool, no extra style indirection. If we ever want a
  proper "Header Row" style binding instead, that's a per-cell
  styleBuiltIn change in `applyCell`.

Image rendering:
- We pass `<img src="…" alt="…" title="…"/>` to
  `paragraph.insertHtml(html, End)` and let Word's HTML paste pipeline
  fetch the URL. Data URIs decode inline; if the fetch fails (offline
  / CORS / 404) Word shows the alt text — matches `<img>` behavior in
  a browser.
- `applyImage` escapes the src/alt/title via the local `escapeHtml`
  helper before splicing them into the tag. Don't skip this — the
  src/alt strings are untrusted user input that lands inside an HTML
  attribute.
- The cell path in `applyCell` calls `applyImage` directly on the
  cell's first paragraph (no separate run/range dance — images are
  block-shaped inside their inline slot).

After basic Markdown: style-mapping UI (M5), footnotes/math (M6),
polish + distribution (M7).
