# Markwright

> Markwright crafts polished Word documents from raw markdown.

A cross-platform Word add-in (Windows, Mac, web) that pastes markdown into the
active document and maps it onto that document's own styles. Built on
Office.js + TypeScript.

## Status

**Basic Markdown is complete.** The task pane parses markdown into a
small `Block` AST and writes it into the document via the Word object
model (`paragraph.styleBuiltIn`, `range.font.*`, `paragraph.startNewList()`,
`paragraph.insertTable`, `paragraph.insertHtml` for `<hr>` and `<img>`),
so every CommonMark + GFM construct binds to the host document's own
styles. Next milestones: style binding UI (M5), then stretch
(footnotes, math).

**Supported — full CommonMark + GFM:**

- Paragraphs
- Headings (ATX `#`..`######` *and* Setext `===` / `---`) — bind to Word's
  built-in Heading 1–6 styles
- Inline: `**bold**`, `*italic*` / `_italic_`, `~~strike~~`, `` `inline code` ``
  (monospace font), `[label](url)`, autolinks `<https://…>`, bare-URL
  autolinks, reference-style links (`[label][ref]` + `[ref]: url`), hard
  line breaks (two trailing spaces), backslash escapes, named and numeric
  HTML entities
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
- GFM **tables** with per-column alignment (`:---`, `:---:`, `---:`),
  inline marks inside cells, and bold header rows — rendered via
  `paragraph.insertTable` against an empty anchor that we reuse for
  the next block
- **Images** (`![alt](src)`, reference-style `![alt][ref]`, optional
  `"title"`) — rendered as `<img>` via `paragraph.insertHtml`, so Word's
  HTML paste pipeline fetches the URL. Data URIs work directly; if the
  fetch fails Word falls back to the alt text (browser `<img>` behavior)

**Not yet** — footnotes, math, the style-mapping UI, and the OOXML fast
path for bulk insertion.

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
| lines      |        95 |
| statements |        90 |
| functions  |       100 |
| branches   |        90 |

## Architecture

```
markdown text
     │
     ▼
markdown-it tokens
     │  parseMarkdown() walks tokens and dispatches per block kind
     │  (heading / paragraph / list-item / code / hr / table). The
     │  shared flattenInline() collapses inline children into
     │  Inline[] = (Run | Image)[] with an InlineState struct
     │  tracking mark depths + the current link. List items get a
     │  synthetic listId per top-level Markdown scope so the applier
     │  knows when to start a new Word.List vs. extend the active one.
     ▼
Block[]   ← stable seam between parsing and host integration
     │  ┌──────────────────────────────────────────────────────────┐
     │  │ paragraph  { runs: Inline[], quoteDepth? }               │
     │  │ heading    { level: 1..6, runs: Inline[] }               │
     │  │ listItem   { ordered, depth, listId, runs, checked? }   │
     │  │ codeBlock  { content: string, language? }                │
     │  │ thematicBreak  {}                                        │
     │  │ table      { header: Inline[][], rows: Inline[][][],     │
     │  │              alignments: Alignment[] }                   │
     │  │ Inline = Run | Image                                     │
     │  │ Run        { text, bold?, italic?, strike?, code?, link? }│
     │  │ Image      { src, alt, title? }                          │
     │  └──────────────────────────────────────────────────────────┘
     ▼
Office.js (Word object model)
     paragraph.styleBuiltIn = Heading1..6 | Normal | ListParagraph | Quote
     range.font.bold / italic / strikeThrough / name
     range.hyperlink
     paragraph.startNewList() + setLevelBullet/Numbering + listItem.level
     paragraph.insertHtml("<hr/>", Replace)  for thematic breaks
     paragraph.insertHtml('<img src="…" alt="…"/>')  for images
     paragraph.insertTable(rows, cols, Before) + cell.body + horizontalAlignment
```

The `Block` AST is the load-bearing abstraction. An earlier attempt
(`pivot to Office.js`, see PR #5) generated a Flat OPC OOXML fragment
and called `Range.insertOoxml`; Word interpreted our minimal style
declarations as the authoritative definitions of Heading1–6 and
silently flattened the output. The object-model path binds to the host
document's real styles instead. OOXML emission stays in the design as
a future option for surfaces the object model can't reach (footnotes,
math, possibly a bulk-insertion fast path for very large documents),
implemented as another emitter over the same `Block[]`.

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
  taskpane/
    taskpane.ts          UI bootstrap + Convert button handler (DOM only)
    apply.ts             Office.js applier: walks Block[] into the doc
    taskpane.html/css    Task pane shell
  commands/              Ribbon function-file (reserved for future actions)
  assets/                Manifest icons
tests/                   Vitest suite for the converter
webpack.config.js        Dual-entry build, copies manifest, serves over HTTPS
eslint.config.mjs        Flat ESLint config: @typescript-eslint +
                         eslint-plugin-office-addins + strict additions
tsconfig.json
```

## Roadmap

**Done** — M3 (CommonMark) and M4 (GFM): paragraphs, headings (ATX +
Setext), inline marks, lists (bullet, ordered, nested, task), code
blocks, blockquotes, thematic breaks, tables with alignment, images,
and the supporting CommonMark verifications (reference links, HTML
entities).

**Open polish** — small lossy spots flagged in the code:

- Visual indent scaling for nested blockquotes (`quoteDepth > 1`)
- Code blocks, thematic breaks, and tables nested inside blockquotes
  (currently dropped — rare in practice)
- Tables inside list items (dropped — GFM doesn't formally allow this)

**M5 — style binding.** Read the host document's named styles, expose
a mapping UI so users can pick which Word style each Markdown
construct binds to, persist the mapping via the Office `Settings`
API.

**M6 — stretch.** Footnotes (Pandoc-style), math (LaTeX → OMML). Both
likely revisit OOXML emission since the object model doesn't expose
them.

**M7 — polish + distribution.** Manifest cleanup for store submission,
icons at additional sizes if needed, end-to-end sideload docs.

## License

MIT. See [LICENSE](./LICENSE).
