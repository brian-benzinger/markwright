# Markwright

> Markwright crafts polished Word documents from raw markdown.

A cross-platform Word add-in (Windows, Mac, web) that pastes markdown into the
active document and maps it onto that document's own styles. Built on
Office.js + TypeScript.

## Install

### Recommended — sideload the hosted add-in (no server, ever)

Markwright's task pane is hosted on GitHub Pages at
`https://brian-benzinger.github.io/markwright/`. The add-in is pure
client-side code — markdown is parsed and written into your document
entirely inside Word; nothing about your document leaves your machine.
That means there's **no server to run**: you sideload the manifest **once**
and it works forever, including after reboots.

1. Download the manifest:
   [`https://brian-benzinger.github.io/markwright/manifest.xml`](https://brian-benzinger.github.io/markwright/manifest.xml)
2. Sideload it once for your platform using the [steps below](#manual-path)
   — but **skip the cert install and dev server**; those are only for the
   source build. The hosted manifest already points at a public HTTPS
   origin, so Word trusts it as-is.

Updates ship automatically: each push to `main` redeploys `dist/` to
Pages, so a refreshed task pane reaches you without re-sideloading (bump
`<Version>` only when ribbon-cached resources like icons change).

### Offline / desktop one-shot — `npx markwright` (Win/Mac)

Prefer not to rely on the hosted origin (air-gapped, or you want the pane
served from your own machine)? `npx markwright` serves the prebuilt pane on
`https://localhost:3000` locally and sideloads it:

```bash
npx markwright          # serves the prebuilt pane locally and sideloads it into Word
npx markwright stop     # unregister when you're done
```

This downloads the published package, installs the local HTTPS cert (so
Word will trust the add-in), serves the task pane from **your** machine,
registers the manifest, and launches Word. The trade-off: the add-in only
runs while `npx markwright` is running.

> Word **on the web** can't be driven by a CLI (sideload there is a
> tenant feature). For web, use the hosted manifest above (or the
> [manual upload](#manual-path)).

### Build from source

For hacking on Markwright. Needs **Node 18+** and **Microsoft Word**.

```bash
npm install
npm run sideload        # certs + hot-reloading dev server + opens Word
npm run sideload:stop   # unregister
```

`npm run sideload` shells out to `npx --yes office-addin-debugging` — same
one-command sideload, but pointed at the webpack **dev server** so source
edits hot-reload. The `--yes` flag fetches the tool into npx's cache on
demand; it is **not** a project dependency, so it never bloats
`node_modules` (see [Dependency footprint](#dependency-footprint) for why
that matters).

### Manual path

Use this to sideload the manifest by hand — whether it's the **hosted**
`manifest.xml` you downloaded from
[GitHub Pages](#recommended--sideload-the-hosted-add-in-no-server-ever)
(nothing else to run) or a **local source build** you're hacking on.

If you're sideloading the hosted manifest, **skip the next block** — the
content is already served over HTTPS from Pages. The cert + dev server are
only needed when you want Word pointed at your own `dist/` over
`https://localhost:3000`:

```bash
npm install
npx office-addin-dev-certs install   # first time per machine — installs the local CA
npm run dev-server                   # serves the bundle on https://localhost:3000
```

Word refuses to load add-ins over plain HTTP, hence the cert install.
Either way, sideload the manifest **once** per host — the steps differ per
surface (below, "your `manifest.xml`" means the hosted download or your
local `dist/manifest.xml`).

<details>
<summary><b>Word desktop (Windows)</b> — trusted shared-folder catalog</summary>

There's no "Upload My Add-in" button on Windows desktop; sideload goes
through a trusted **shared-folder catalog**. The catalog URL must be a
**UNC network share** (`\\…`) — a plain `C:\…` path is rejected — and it
must point at the **folder that directly contains** `manifest.xml`, not a
parent or child of it.

1. Make a folder and put your downloaded `manifest.xml` **inside it**
   (e.g. `C:\Users\<you>\Downloads\markwright\manifest.xml`).
2. Turn that folder into a share. Easiest from an **admin PowerShell**:
   ```powershell
   New-SmbShare -Name WordAddins -Path C:\Users\<you>\Downloads\markwright -FullAccess "$env:USERNAME"
   ```
   …which gives you the catalog URL `\\localhost\WordAddins`. (GUI
   alternative: right-click the folder → *Properties → Sharing → Share…*
   → add yourself → **Share**; the share name defaults to the folder name.
   The hidden admin share also works without creating one:
   `\\localhost\c$\Users\<you>\Downloads\markwright`.)
3. In Word, *File → Options → Trust Center → Trust Center Settings →
   Trusted Add-in Catalogs.* Paste the `\\…` URL in **Catalog Url**, click
   **Add Catalog**, check **Show in Menu** for that row, **OK**, then
   **fully restart Word** (kill any lingering `WINWORD.EXE` in Task
   Manager).
4. Load it: *Insert → My Add-ins → **More Add-ins*** → **SHARED FOLDER**
   tab → **Markwright** → **Add**. (Even with "Show in Menu" ticked you
   still pick it from this dialog once; after that it's remembered.)

Troubleshooting:
- **"No add-ins available"** in the Shared Folder tab → the catalog folder
  doesn't actually contain `manifest.xml`. Open the exact `\\…` path in
  File Explorer and confirm the file sits in that folder's root (the
  classic miss is the manifest one level up, e.g. in `Downloads\` while
  the catalog points at `Downloads\markwright\`).
- **"Windows can't find \\\\localhost\\…"** → that name isn't a real share.
  `\\localhost\<name>` only resolves if `<name>` is a *shared* folder, not
  an arbitrary path. Create the share (step 2) or use the `c$` admin form.
- Still empty after fixing the path → Word cached the empty catalog. Clear
  it: delete everything inside `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
  and reopen Word.

</details>

> **Less hassle:** the shared-folder dance is Windows-desktop-only. **Word
> on the web** sideloads the same manifest with a plain file picker (no
> sharing) — see the Word web row below.

<details>
<summary><b>Word desktop (Mac)</b> — the <code>wef</code> folder</summary>

Drop your `manifest.xml` into
`~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` (create the
`wef` folder if it doesn't exist, lowercase). Restart Word and the add-in
appears under *Insert → My Add-ins*.

</details>

<details>
<summary><b>Word web</b> — Upload My Add-in (work/school accounts only)</summary>

Available only on **work / school M365 accounts** (consumer Microsoft
accounts have sideload disabled). The path is *Home → Add-ins → More
Add-ins → MY ADD-INS tab → Upload My Add-in* (link at the bottom of the
dialog) → pick your `manifest.xml`. No folder sharing — this is the
lowest-friction surface. If your tenant admin has blocked custom add-ins,
the upload link is hidden.

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

- **Style mapping** — a **Styles** panel lets you bind each block construct
  (Heading 1–6, paragraph, blockquote, code block) to any of the host
  document's paragraph styles; the choice persists per-document via the
  Office `Settings` API. Left unconfigured, everything uses the built-in
  defaults above.

**Not yet** — footnotes, math, and the OOXML fast path for bulk insertion.

## Before 0.1.0 — release guardrails

0.1.0 ships two ways: the **hosted manifest on GitHub Pages** (the
recommended, zero-server install) and the **`npx markwright`** local
fallback. It's not yet a Store-listed add-in. Before tagging or
`npm publish`, walk this checklist:

- [x] **Hosting model: GitHub Pages (static).** `dist/` is published to
      `https://brian-benzinger.github.io/markwright/` by
      `.github/workflows/deploy.yml`, and `urlProd` in `webpack.config.js`
      points the production manifest there. The add-in is pure client-side
      code, so the host is a plain static origin — no backend to operate.
      `npx markwright` remains as a localhost fallback for offline/air-gapped
      use. (A future *Store* listing builds on top of this same hosted
      origin — AppSource distributes the manifest but does not host the JS.)
- [ ] **npm publish readiness.** `private` is removed, `files` whitelists
      `bin`/`dist`/`manifest.xml`, and `prepack` runs the production build
      so the tarball always carries a fresh `dist/`. Before publishing:
      confirm the `markwright` name is free, you're authed (`npm whoami`),
      and `npm pack` shows the expected contents.
- [ ] **Keep versions in lockstep.** `package.json` `version` and the
      manifest `<Version>` (4-part) must agree. Currently both read
      `0.1.0`. Bump them together each release.
- [ ] **Run the gates** — `npm run typecheck && npm run lint && npm test
      && npm run build`. The build is size-gated (200 KB/asset), so a
      regression there fails CI too.
- [ ] **Validate the manifest** — `npm run validate`.
- [ ] **`npm audit` clean** — no known high/critical advisories.
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

Runtime deps are `markdown-it` (bundled into the task pane) and
`office-addin-dev-certs` (the `npx markwright` CLI uses it to serve `dist/`
over trusted HTTPS — see `bin/serve.mjs`). Dev tooling is `vitest`,
`webpack`, `typescript`, `eslint` + plugins, and `prettier` — picked
individually rather than via `office-addin-lint` or
`office-addin-debugging` wrappers, both of which dragged in large
transitive trees (the debugging wrapper pulled ~290 MB of Azure ARM SDK
via its `@microsoft/m365agentstoolkit-cli` peer) without proportional
value.

That's why both `npm run sideload` and the published `markwright` CLI
reach `office-addin-debugging` through `npx --yes` instead of listing it
as a dependency: whoever wants the one-command sideload pays a one-time
download into npx's cache, while everyone keeps the lean `node_modules`.
Don't "fix" the sideload path by adding the dependency — that's the
290 MB we deliberately removed.

## Project layout

```
manifest.xml             XML add-in-only manifest (cross-platform)
bin/
  markwright.mjs         `npx markwright` entry — sideloads via office-addin-debugging
  serve.mjs              HTTPS static server for the prebuilt dist/ (CLI dev server)
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

**M5 — style binding.** ✅ Done. Reads the host document's named
paragraph styles (`getStyles()`), exposes a **Styles** mapping panel so
users pick which Word style each block construct binds to, and persists
the mapping per-document via the Office `Settings` API. Defaults
reproduce the original built-in bindings, so the zero-config path is
unchanged. Pure map logic lives in `src/convert/styleMap.ts` (tested);
the Office.js glue is in `src/taskpane/styles.ts`. Remaining: mapping
list-item and inline (character) styles.

**M6 — stretch.** Footnotes (Pandoc-style), math (LaTeX → OMML). Both
likely revisit OOXML emission since the object model doesn't expose
them.

**M7 — polish + distribution.** npm distribution via `npx markwright` is
in (this milestone's first slice); remaining: manifest cleanup for Store
submission (which would revisit the hosting-URL guardrail above), icons at
additional sizes if needed, end-to-end sideload docs.

## License

MIT. See [LICENSE](./LICENSE).
</content>
</invoke>
