# Markwright

> Markwright crafts polished Word documents from raw markdown.

A cross-platform Word add-in (Windows, Mac, web) that pastes markdown into the
active document and maps it onto that document's own styles. Built on
Office.js + TypeScript.

This repository is at **Milestone 1**: a sideloadable skeleton with a task
pane, a textarea, and a Convert button that drops the raw input at the
selection. No markdown parsing yet — that arrives in Milestone 2.

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
npm test            # vitest run
npm run build       # production webpack build
```

All four also run in CI on every PR via `.github/workflows/ci.yml`.

## Project layout

```
manifest.xml             XML add-in-only manifest (cross-platform)
src/
  convert/               Markdown → OOXML emitter (pure functions, unit-tested)
  taskpane/              Task pane UI (textarea + Convert button)
  commands/              Ribbon function-file (reserved for future actions)
  assets/                Manifest icons
tests/                   Vitest suite for the converter
webpack.config.js        Dual-entry build, copies manifest, serves over HTTPS
eslint.config.mjs        Flat ESLint config extending office-addin-lint
tsconfig.json
```

## Roadmap

See the design brief for the full plan. Next milestones:

2. Parse + naive insert — wire up `markdown-it`, emit a trivial OOXML fragment.
3. MVP coverage — AST-to-OOXML emitter for headings, lists, emphasis, code,
   blockquotes, links, tables. Defaults to built-in Word styles.
4. Style binding — read the document's named styles, expose a mapping UI,
   persist via the Office `Settings` API.
5. Stretch — images, footnotes, math.
6. Polish + distribution.

## License

MIT. See [LICENSE](./LICENSE).
