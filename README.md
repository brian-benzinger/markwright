# Markwright

> Markwright crafts polished Word documents from raw markdown.

A cross-platform Word add-in (Windows, Mac, web) that pastes markdown into the
active document and maps it onto that document's own styles. Built on
Office.js + TypeScript.

## Status

Mid-**Milestone 3**. The task pane parses markdown into a small Block AST and
writes it into the document via the Word object model (`paragraph.styleBuiltIn`,
`range.font.*`, `paragraph.startNewList()`), so headings, paragraphs, inline
marks and lists bind to the host document's own styles.

**Supported today**

- Paragraphs
- Headings `#` through `######` (bind to Word's built-in Heading 1–6 styles)
- Inline: `**bold**`, `*italic*` / `_italic_`, `~~strike~~`, `` `inline code` ``
  (monospace font), `[links](url)`, autolinks `<https://…>`, hard line breaks
  (two trailing spaces), backslash escapes
- Bullet lists, ordered lists, mixed nested lists (continuous numbering per
  Markdown list scope)
- Blockquotes (bind to Word's "Quote" style; nesting tracked in the AST as
  `quoteDepth` though depth-2+ doesn't visually scale indent yet)
- Fenced code blocks (` ``` `, with language tag captured) and 4-space
  indented code blocks — rendered as one monospaced (Consolas) Word
  paragraph with `\v` line breaks so the snippet stays one block
- Thematic breaks (`---`, `***`, `___`) — rendered via
  `paragraph.insertHtml("<hr/>", Replace)`, which Word interprets as a
  bottom-bordered paragraph
- GFM **task lists** (`- [ ]`, `- [x]`, `- [X]`, including in ordered
  lists) — rendered with `☐` / `☑` prepended to the run text
- GFM **bare-URL autolinks** (`https://x` in flowing text becomes a
  hyperlink)
- GFM **tables** with per-column alignment (`:---`, `:---:`, `---:`),
  inline marks inside cells, and bold header rows — rendered via
  `paragraph.insertTable` against an empty anchor that we reuse for
  the next block

**Not yet** — images, footnotes, math, the style-mapping UI, and the
OOXML fast path for bulk insertion.

## Prerequisites

- Node.js 18+ and npm
- Microsoft Word (desktop or web). For desktop sideloading, see Microsoft's
  [add-in sideloading docs](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins).

## Develop

```bash
npm install
npx office-addin-dev-certs install   # first time per machine — installs the local CA
npm run dev-server                   # serves the bundle on https://localhost:3000
```

Word refuses to load add-ins over plain HTTP, hence the cert install. Then
sideload the manifest **once** per host:

- **Word desktop:** *Insert → My Add-ins → Manage My Add-ins → Upload My Add-in*
  → pick `dist/manifest.xml`.
- **Word web:** *Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in*
  → pick `dist/manifest.xml`.

After the one-time sideload, edits hot-reload through the dev server — just
refresh the task pane.

> Earlier versions of this repo used `office-addin-debugging` for one-command
> sideloading; we dropped it because its `@microsoft/m365agentstoolkit-cli`
> peer dragged in the Azure ARM SDK (~290 MB of `node_modules`) for a
> first-time-only convenience.

## Check before pushing

```bash
npm run typecheck   # tsc --noEmit
npm run lint        # eslint . && prettier --check ...
npm run lint:fix    # auto-fix formatting and trivially-fixable lint
npm test            # vitest run --coverage
npm run build       # production webpack build (no sourcemaps)
```

All four (typecheck, lint, test, build) also run in CI on every PR via
`.github/workflows/ci.yml`. The lint config (`eslint.config.mjs`) layers
the `eslint-plugin-office-addins` rules — which catch real Office.js
footguns like `context.sync()` in loops and reading a property before
calling `load()` — over `@typescript-eslint`'s recommended rules, plus
the project's stricter additions (`no-explicit-any`,
`consistent-type-imports`, `no-non-null-assertion`,
`no-inferrable-types`). Prettier owns formatting; settings live in
`.prettierrc.json`.

### Coverage

`npm test` collects v8 coverage scoped to `src/convert/**` and fails the
gate if any threshold is missed. The Office.js applier in
`src/taskpane/` is excluded because it requires the live Word host —
verify it via sideload, not vitest.

| metric     | threshold |
| ---------- | --------: |
| lines      |        90 |
| statements |        90 |
| functions  |       100 |
| branches   |        90 |

## Architecture

```
markdown text
     │
     ▼
markdown-it tokens
     │  parseMarkdown() walks tokens, dispatches heading/paragraph/
     │  list-item, and lets flattenInline() collapse inline children
     │  into Run[] with a shared InlineState (mark depths + link). List
     │  items get a synthetic listId per top-level Markdown scope.
     ▼
Block[]   ← stable seam between parsing and host integration
     │  ┌──────────────────────────────────────────────────────────┐
     │  │ paragraph  { runs: Run[], quoteDepth? }                  │
     │  │ heading    { level: 1..6, runs: Run[] }                  │
     │  │ listItem   { ordered, depth, listId, runs, checked? }   │
     │  │ codeBlock  { content: string, language? }                │
     │  │ thematicBreak  {}                                        │
     │  │ table      { header: Run[][], rows: Run[][][],           │
     │  │              alignments: Alignment[] }                   │
     │  │ Run        { text, bold?, italic?, strike?, code?, link? }│
     │  └──────────────────────────────────────────────────────────┘
     ▼
Office.js (Word object model)
     paragraph.styleBuiltIn = Heading1..6 | Normal | ListParagraph | Quote
     range.font.bold / italic / strikeThrough / name
     range.hyperlink
     paragraph.startNewList() + setLevelBullet/Numbering + listItem.level
     paragraph.insertHtml("<hr/>", Replace)  for thematic breaks
     paragraph.insertTable(rows, cols, Before) + cell.body + horizontalAlignment
```

The Block AST is the load-bearing abstraction. An earlier attempt (`pivot to
Office.js`, see PR #5) generated a Flat OPC OOXML fragment and called
`Range.insertOoxml`; Word interpreted our minimal style declarations as the
authoritative definitions of Heading1–6 and silently flattened the output.
The object-model path binds to the host document's real styles instead. OOXML
emission stays in the design as a future option for bulk content (tables,
images, footnotes), implemented as another emitter over the same `Block[]`.

## Dependency footprint

Markwright deliberately keeps its install thin so cloud containers, CI
cold runs, and contributor laptops aren't billed for tooling we don't
use. The current footprint:

| | size |
| --- | ---: |
| `node_modules` | ~237 MB |
| `dist/` (production, no sourcemaps) | ~180 KB |
| `dist/` (development, sourcemaps on) | ~912 KB |

Runtime deps are just `markdown-it`. Dev tooling is `vitest`, `webpack`,
`typescript`, `eslint` + plugins, and `prettier` — picked individually
rather than via `office-addin-lint` or `office-addin-debugging`
wrappers, both of which dragged in large transitive trees (the
debugging wrapper pulled ~290 MB of Azure ARM SDK via its
`@microsoft/m365agentstoolkit-cli` peer) without proportional value.

## Project layout

```
manifest.xml             XML add-in-only manifest (cross-platform)
src/
  convert/               markdown-it → Block[] AST (pure, unit-tested)
  taskpane/              Task pane UI + Office.js applier
  commands/              Ribbon function-file (reserved for future actions)
  assets/                Manifest icons
tests/                   Vitest suite for the converter
webpack.config.js        Dual-entry build, copies manifest, serves over HTTPS
eslint.config.mjs        Flat ESLint config extending office-addin-lint
tsconfig.json
```

## Roadmap

Remaining in **M3** (CommonMark coverage — minor polish):

- Visual indent scaling for nested blockquotes (`quoteDepth > 1`)
- Code blocks and thematic breaks inside blockquotes (currently dropped
  — rare in practice)

**M4** (GFM): tables (likely requires the Word Table API or revisiting OOXML),
task lists, table-of-contents-friendly headings.

**M5** (style binding): read the document's named styles, expose a mapping UI,
persist via the Office `Settings` API.

**M6** (stretch): images, footnotes, math.

**M7**: polish + distribution.

## License

MIT. See [LICENSE](./LICENSE).
