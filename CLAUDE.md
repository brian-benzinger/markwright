# CLAUDE.md — Markwright

Notes for future Claude sessions. Things that aren't obvious from the
code or that bit me in earlier sessions.

## Architecture seam

The load-bearing abstraction is `Block[]` in `src/convert/index.ts`.
Parsing markdown is one concern; applying it to Word is another.
Adding a new block type means: extend `Block` (and `Run` if it's
inline-level), teach `parseMarkdown` to emit it, teach `applyBlock`
in `src/taskpane/taskpane.ts` to render it. Tests live in
`tests/convert.test.ts` and stay pure — they never touch Office.js.

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
  `markup: "autolink"`.
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
- `npm run lint:fix` cleans up Prettier nits.
- Coverage gate runs in `npm test` via `vitest run --coverage`. Only
  `src/convert/` is in scope — the Office.js applier can't execute
  under vitest because the `Word` global only exists in the host.
  Thresholds in `vitest.config.ts`: lines 90, statements 90,
  functions 100, branches 90.

## Outstanding M3 work

- Blockquotes (parser currently suppresses contents)
- Fenced and indented code blocks
- Thematic breaks (`---`)

After M3: tables (M4), style-mapping UI (M5),
images/footnotes/math (M6), polish + distribution (M7).
