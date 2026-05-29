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

**Not yet** — thematic breaks (`---`), tables, task lists, images,
footnotes, math, the style-mapping UI, and the OOXML fast path for bulk
insertion.

## Prerequisites

- Node.js 18+ and npm
- Microsoft Word (desktop or web). For desktop sideloading, see Microsoft's
  [add-in sideloading docs](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins).

## Develop

```bash
npm install
npm start          # builds, starts the dev server, and sideloads into Word desktop
# or
npm run dev-server # just runs the webpack dev server on https://localhost:3000
npm run start:web  # sideload into Word on the web
```

On first run, `office-addin-dev-certs` will install a local CA so the dev
server can serve HTTPS — Word refuses to load add-ins over plain HTTP.

## Check before pushing

```bash
npm run typecheck   # tsc --noEmit
npm run lint        # office-addin-lint (ESLint + Prettier)
npm run lint:fix    # auto-fix formatting and trivially-fixable lint
npm test            # vitest run --coverage
npm run build       # production webpack build
```

All four (typecheck, lint, test, build) also run in CI on every PR via
`.github/workflows/ci.yml`. The lint config extends `office-addin-lint`
and additionally enforces `no-explicit-any`, `consistent-type-imports`,
`no-non-null-assertion`, and `no-inferrable-types` (`eslint.config.mjs`)
— Prettier remains the only source of formatting rules.

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
     │  │ listItem   { ordered, depth, listId, runs: Run[] }       │
     │  │ codeBlock  { content: string, language? }                │
     │  │ Run        { text, bold?, italic?, strike?, code?, link? }│
     │  └──────────────────────────────────────────────────────────┘
     ▼
Office.js (Word object model)
     paragraph.styleBuiltIn = Heading1..6 | Normal | ListParagraph | Quote
     range.font.bold / italic / strikeThrough / name
     range.hyperlink
     paragraph.startNewList() + setLevelBullet/Numbering + listItem.level
```

The Block AST is the load-bearing abstraction. An earlier attempt (`pivot to
Office.js`, see PR #5) generated a Flat OPC OOXML fragment and called
`Range.insertOoxml`; Word interpreted our minimal style declarations as the
authoritative definitions of Heading1–6 and silently flattened the output.
The object-model path binds to the host document's real styles instead. OOXML
emission stays in the design as a future option for bulk content (tables,
images, footnotes), implemented as another emitter over the same `Block[]`.

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

Remaining in **M3** (CommonMark coverage):

- Thematic breaks (`---`)
- Visual indent scaling for nested blockquotes (`quoteDepth > 1`)
- Code blocks inside blockquotes (currently dropped — rare in practice)

**M4** (GFM): tables (likely requires the Word Table API or revisiting OOXML),
task lists, table-of-contents-friendly headings.

**M5** (style binding): read the document's named styles, expose a mapping UI,
persist via the Office `Settings` API.

**M6** (stretch): images, footnotes, math.

**M7**: polish + distribution.

## License

MIT. See [LICENSE](./LICENSE).
