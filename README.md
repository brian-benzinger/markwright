# Markwright

> Markwright crafts polished Word documents from raw markdown.

A cross-platform Word add-in (Windows, Mac, web) that pastes markdown into the
active document and maps it onto that document's own styles. Built on
Office.js + TypeScript.

## Install

### Prerequisites

- Node.js 18+ and npm
- Microsoft Word (desktop or web)

### Easy path — one command (Word desktop, Win/Mac)

```bash
npm install
npm run sideload      # builds certs, starts the dev server, opens Word with Markwright loaded
```

`npm run sideload` shells out to `npx --yes office-addin-debugging` — it
installs the local HTTPS cert (so Word will load the add-in), starts the
dev server on `https://localhost:3000`, registers the manifest, and
launches Word with Markwright already in the ribbon. The `--yes` flag
means npx fetches the tool on demand into its cache; it is **not** a
project dependency, so it never bloats `node_modules` (see
[Dependency footprint](#dependency-footprint) for why that matters).

To stop and unregister:

```bash
npm run sideload:stop
```

> `office-addin-debugging` can only auto-sideload Word **desktop**. For
> Word on the web, use the manual upload below.

### Manual path

Use this if the easy path can't reach your host (Word web, locked-down
desktop, or you'd rather not run the helper). Start the dev server
yourself, then point Word at `dist/manifest.xml`.

```bash
npm install
npx office-addin-dev-certs install   # first time per machine — installs the local CA
npm run dev-server                   # serves the bundle on https://localhost:3000
```

Word refuses to load add-ins over plain HTTP, hence the cert install.
Then sideload the manifest **once** per host — the path differs per
surface.

<details>
<summary><b>Word desktop (Windows)</b> — trusted shared-folder catalog</summary>

There's no "Upload My Add-in" button on Windows desktop; sideload goes
through a trusted shared-folder catalog:

1. Create a folder somewhere (e.g. `C:\Users\<you>\WordAddins`) and copy
   `dist/manifest.xml` into it.
2. In Word, *File → Options → Trust Center → Trust Center Settings →
   Trusted Add-in Catalogs.*
3. Paste the folder path in **Catalog Url**, click **Add Catalog**, then
   check **Show in Menu** for that row. Click OK and restart Word.
4. *Insert → Add-ins → Shared Folder* tab → Markwright → **Add**.

</details>

<details>
<summary><b>Word desktop (Mac)</b> — the <code>wef</code> folder</summary>

Drop `dist/manifest.xml` into
`~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` (create the
`wef` folder if it doesn't exist, lowercase). Restart Word and the add-in
appears under *Insert → My Add-ins*.

</details>

<details>
<summary><b>Word web</b> — Upload My Add-in (work/school accounts only)</summary>

Available only on **work / school M365 accounts** (consumer Microsoft
accounts have sideload disabled). The path is *Home → Add-ins → More
Add-ins → MY ADD-INS tab → Upload My Add-in* (link at the bottom of the
dialog) → pick `dist/manifest.xml`. If your tenant admin has blocked
custom add-ins, the upload link is hidden.

</details>

### After sideload

Open the **Markwright** button on the Home tab, paste markdown into the
pane, and click **Convert** to write it into the active document.

Edits hot-reload through the dev server — just refresh the task pane.
When you change icon assets or any other ribbon-cached resource, bump
`<Version>` in `manifest.xml` and re-add the add-in. Word desktop caches
icons by manifest ID + version under `%LOCALAPPDATA%\Microsoft\Office\
16.0\Wef\`; without a version bump it'll keep showing the old icon.

## What's supported

Full CommonMark + GFM. The task pane parses markdown into a small `Block`
AST and writes it into the document via the Word object model, so every
construct binds to the host document's own styles.

- **Headings** — ATX (`#`..`######`) and Setext (`===` / `---`), bound to
  Word's built-in Heading 1–6 styles
- **Inline** — `**bold**`, `*italic*` / `_italic_`, `~~strike~~`,
  `` `inline code` `` (monospace), `[label](url)`, autolinks `<https://…>`,
  bare-URL autolinks, reference-style links (`[label][ref]`), hard line
  breaks (two trailing spaces), backslash escapes, HTML entities
- **Lists** — bullet, ordered, mixed/nested (continuous numbering per
  Markdown list scope), and GFM **task lists** (`- [ ]` / `- [x]`)
- **Blockquotes** — bound to Word's "Quote" style (nesting tracked in the
  AST; depth-2+ doesn't visually scale indent yet)
- **Code blocks** — fenced and 4-space indented, rendered as one
  monospaced (Consolas) paragraph
- **Thematic breaks** (`---`, `***`, `___`) — rendered as a bottom-bordered
  paragraph via `insertHtml("<hr/>")`
- **Tables** — GFM, with per-column alignment, inline marks in cells, and
  bold header rows
- **Images** (`![alt](src)`, reference-style, optional `"title"`) — rendered
  as `<img>` so Word fetches the URL; data URIs decode inline, failed
  fetches fall back to alt text

**Not yet** — footnotes, math, the style-mapping UI, and the OOXML fast
path for bulk insertion.

## Before 0.1.0 — release guardrails

This is a **local-sideload preview**, not a Store-distributable build.
Before tagging or publishing 0.1.0, walk this checklist:

- [ ] **Pick a hosting URL.** `manifest.xml` and `webpack.config.js`
      (`urlProd`) both point at `https://localhost:3000`. A production
      build is therefore only loadable on a machine running the dev
      server — it is **not** distributable as-is. Set `urlProd` to the
      real HTTPS host and rewrite the manifest URLs before shipping
      anything beyond local sideload.
- [ ] **Keep versions in lockstep.** `package.json` `version` and the
      manifest `<Version>` (4-part) must agree. Currently both read
      `0.1.0`. Bump them together each release.
- [ ] **Run the gates** — `npm run typecheck && npm run lint && npm test
      && npm run build`. The build is size-gated (200 KB/asset), so a
      regression there fails CI too.
- [ ] **Validate the manifest** — `npm run validate`.
- [ ] **`npm audit` clean** — no known high/critical advisories in the
      shipped dependency (`markdown-it`).
- [ ] **Image fetch is remote content.** Images are handed to Word's HTML
      paste pipeline, which fetches arbitrary URLs from the user's
      machine. `src`/`alt`/`title` are HTML-escaped before they hit the
      tag (untrusted input lands inside an attribute — don't skip this),
      but document the privacy implication: converting a doc with remote
      image URLs makes outbound requests.
- [ ] **Decide the support surface** — `SupportUrl` in the manifest points
      at the GitHub repo; confirm that's where you want issues to land.

## Develop

Run all four gates locally before pushing — CI runs the same set on every
PR via `.github/workflows/ci.yml`:

```bash
npm run typecheck   # tsc --noEmit
npm run lint        # eslint . && prettier --check ...
npm run lint:fix    # auto-fix formatting and trivially-fixable lint
npm test            # vitest run --coverage
npm run build       # production webpack build (no sourcemaps, size-gated)
```

The production build enforces a **moderate size budget** via webpack's
`performance.hints = "error"`: each JS asset must stay under **200 KB**,
each entrypoint under **250 KB**. Current `taskpane.js` is ~132 KB, so
there's ~35% headroom. The dev build is exempt (sourcemaps push it past
the limit by design).

The lint config (`eslint.config.mjs`) layers the
`eslint-plugin-office-addins` rules — which catch real Office.js footguns
like `context.sync()` in loops and reading a property before calling
`load()` — over `@typescript-eslint`'s recommended rules, plus the
project's stricter additions (`no-explicit-any`,
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
| `node_modules` | ~278 MB |
| `dist/` (production, no sourcemaps) | ~184 KB |
| `dist/` (development, sourcemaps on) | ~936 KB |

Runtime deps are just `markdown-it`. Dev tooling is `vitest`, `webpack`,
`typescript`, `eslint` + plugins, and `prettier` — picked individually
rather than via `office-addin-lint` or `office-addin-debugging`
wrappers, both of which dragged in large transitive trees (the
debugging wrapper pulled ~290 MB of Azure ARM SDK via its
`@microsoft/m365agentstoolkit-cli` peer) without proportional value.

That's why `npm run sideload` reaches `office-addin-debugging` through
`npx --yes` instead of listing it as a devDependency: contributors who
want the one-command sideload pay a one-time download into npx's cache,
while everyone else keeps the lean `node_modules`. Don't "fix" the
sideload script by re-adding the dependency — that's the 290 MB we
deliberately removed.

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

**M7 — polish + distribution.** Manifest cleanup for store submission
(starting with the hosting-URL guardrail above), icons at additional
sizes if needed, end-to-end sideload docs.

## License

MIT. See [LICENSE](./LICENSE).
</content>
</invoke>
