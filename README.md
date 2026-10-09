# FreeWord

A free, offline, cross-platform Word-like document editor built with Electron + Tiptap —
with an optional **Notion mode**. No accounts, no cloud, no telemetry — your documents
stay on your machine.

## Features

- **Notion mode** (View → Notion Mode, or Ctrl+Shift+Space): a continuous, pageless
  canvas with a minimal chrome, on top of the exact same document model:
  - **Slash menu** — type `/` on an empty line for a keyboard-navigable block picker
  - **Bubble toolbar** — select text and the mark commands float above it
  - **Block handle** — ⋮⋮ on hover: turn into, duplicate, delete, move, drag to reorder
  - **Workspace sidebar** — a local page tree with nesting, drag-to-reparent,
    emoji icons, and **Ctrl+K quick switcher**; pages **autosave** (debounced) to
    `userData/workspace/`, no Ctrl+S needed; deleted pages land in a sidebar
    **Trash** section with Restore / Delete-forever / Empty-trash actions
  - **Page covers & titles** — a Notion-style header above the canvas for workspace
    pages: cover image (native picker), big emoji icon, and an inline-editable
    title (covers/icons/titles live in page metadata, never in the document)
  - **Duplicate page** — the sidebar row menu deep-copies a page and its whole
    subtree (fresh ids, same content and cover) beside the original
  - **Dark theme** (View → Theme → Light / Dark / Follow system), persisted
- **Rich text editing**: bold / italic / underline / strikethrough, text color,
  yellow & green highlighting, inline code (monospace), hyperlinks,
  10 font families, 15 font sizes
- **Paragraph styles**: Paragraph, Heading 1–3, alignment (left / center / right / justify),
  line spacing (1.0 / 1.15 / 1.5 / 2.0)
- **Blocks**: bulleted and numbered lists with nesting (Tab / Shift+Tab),
  to-do lists (checkboxes, click or Space), callouts (clickable icon),
  **toggle lists** (click the ▸ caret to collapse everything below the first
  line), a live **table of contents** (renders the document's headings;
  click an entry to jump to it), blockquotes, dividers and fenced code blocks —
  most markdown input rules work (`# `, `- `, `1. `, `[] `, `> `, `>> `, `---`, ` ` ```)
- **Soft line break**: Shift+Enter wraps inside a paragraph without starting a new one
- **Tables**: insert a 3×3 table with a header row, add / delete rows and columns,
  toggle the header row, **drag column borders to resize**, **merge and split cells**
  (drag across cells to select them). Tables survive DOCX round-trips, including
  the widths you set and the merges you made.
- **Images**: insert PNG / JPEG / GIF / BMP images (≤ 10 MB) at the cursor;
  images are embedded in the document and exported to DOCX with their size
  preserved. **Hover an image and drag a corner to resize it** (aspect ratio
  is kept; the size lands in the document and survives .docx round-trips).
- **Paged document**: true page model (A4 or Letter, adjustable margins via
  View → Page Setup), live pagination in the editor, page breaks (Ctrl+Enter),
  page & word count
- **File formats**:
  - Open / save **`.docx`** (round-trip via the `docx` library)
  - Open **`.txt`** and **`.html`**
  - Export **PDF** (print engine), **plain text** and **HTML**, and print
- **Right-click menu**: Cut / Copy / Paste and spell-check suggestions for the
  word under the cursor (Electron ships no context menu of its own)
- **Zoom** 50–200% (Ctrl+= / Ctrl+- step beyond the presets, 25–400%),
  built-in spell check, recent files list
- **Find & Replace** (Ctrl+F / Ctrl+H): highlighted matches with a
  next/prev counter, match-case toggle, Replace and Replace-All (one undo
  step reverts the whole batch)
- **Headers & footers** (Insert → Header & Footer…): text on every page with
  `{{page}}` / `{{pages}}` tokens; exported as real Word headers/footers with
  live PAGE/NUMPAGES fields
- **Footnotes** (Insert → Footnote, Ctrl+Alt+F): auto-numbered inline refs
  with the note list at the bottom of the page; exported to real
  `word/footnotes.xml` references
- **Comments** (select text → 💬): threads anchored by a mark, with a
  right-hand Comments panel (View → Comments Panel, Ctrl+Alt+C) for jump /
  edit / resolve / delete; exported as native Word comments
- **Track changes** (Review → Track Changes, Ctrl+Shift+M): insertions and
  deletions recorded with author + date, accept/reject per change or all at
  once; exported as native `w:ins`/`w:del` so Word's own Review pane shows
  them

## Getting started

```bash
npm install
npm run dev        # start the app in development mode (hot reload)
npm start          # run the real app through Electron — built automatically
                   # when the bundles are missing or stale, no dist step
```

### Scripts

| Script                  | What it does                                                                                                                                                                                                                        |
| ----------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `npm run dev`           | Run the app with hot reload                                                                                                                                                                                                         |
| `npm run dev:nosandbox` | Run with the Chromium sandbox disabled (see Troubleshooting)                                                                                                                                                                        |
| `npm start`             | Run the real app through Electron — builds first only when `out/` is missing or stale (no dist/packaging step); if Electron dies instantly (the sandbox-broker case below), it retries once automatically with the sandbox disabled |
| `npm start:nosandbox`   | Same, with the Chromium sandbox disabled                                                                                                                                                                                            |
| `npm run build`         | Build main / preload / renderer bundles to `out/`                                                                                                                                                                                   |
| `npm run typecheck`     | Type-check both the node and web TypeScript projects                                                                                                                                                                                |
| `npm test`              | Run the vitest suite (unit + integration)                                                                                                                                                                                           |
| `npm run lint`          | ESLint over `src/`, `test/` and `scripts/`                                                                                                                                                                                          |
| `npm run format`        | Prettier (also `format:check` for CI)                                                                                                                                                                                               |
| `npm run icons`         | Regenerate `build/icon.*` from `scripts/make-icon.py`                                                                                                                                                                               |
| `npm run dist:win`      | Package a Windows NSIS installer                                                                                                                                                                                                    |
| `npm run dist:mac`      | Package macOS DMGs (x64 + arm64)                                                                                                                                                                                                    |
| `npm run dist:linux`    | Package a Linux AppImage and .deb                                                                                                                                                                                                   |
| `npm run dist:all`      | All three targets at once (on the host that can build them)                                                                                                                                                                         |

Packaging for a platform must be run **on that platform** (or via CI, see below).

## How Notion mode works

Notion mode is a _presentation switch_, not a second document model. The custom
schema stays `doc → page → blocks` either way:

- Turning it on **pauses the pagination engine** (blocks stop being re-partitioned)
  and CSS collapses the pages into a single continuous column — so `.docx` save,
  PDF export and print are byte-identical in both modes. Turning it back off
  re-partitions the document in one pass.
- Print/PDF always renders the Word form first (`beginPrintAppearance()` in
  `features/appearance.ts`), even when triggered from Notion mode.
- Header/footer bands are Word-mode rendering chrome (hidden in Notion mode
  via CSS, like the ribbon); the page-end footnote areas re-flow to trail
  the continuous canvas instead. Find/replace, comments and tracked changes
  behave identically in both modes — they live in the same document model.
- Settings (`notionMode`, `theme`, `sidebarOpen`, `toolbarPinned`) persist to
  `userData/settings.json` via `settings:get` / `settings:set` IPC.
- Workspace pages live in `userData/workspace/`: `pages.json` (metadata:
  id / title / icon / cover / parentId / order) plus one `<id>.json` per page holding
  the same document JSON the .docx exporter consumes. Deleting a page moves its
  file to `workspace/.trash/`, where the sidebar Trash section offers Restore /
  Delete-forever. Page edits autosave 600 ms after the last
  keystroke; the quit handshake flushes before answering, and Ctrl+S is an
  immediate flush.
- To-do / callout / quote / divider / code-block round-trip through .docx via
  custom style ids (`FwTaskItem`, `FwCallout`, …) plus visual fallbacks
  (☐/☑ glyphs, shading, Consolas), so they survive re-saving in Word itself as
  readable, if plainer, paragraphs. Toggles export the same way (`FwToggle`
  summary with a ▸/▾ glyph that doubles as the collapsed flag, `FwToggleBody`
  for the rest); the table-of-contents node exports as `FwToc`-tagged heading
  paragraphs and re-hydrates into one live TOC node on import.

## Troubleshooting

**`npm run dev` dies with `TypeError: Cannot read properties of undefined (reading 'on')`**

If `ELECTRON_RUN_AS_NODE` is set in the environment (CI runners, agent shells,
and any parent process that was itself Electron set it), `electron.exe` does not
start the Electron runtime at all — it runs as a plain Node interpreter, so
`require("electron")` returns the path to the binary as a _string_ and every
named import is `undefined`. `npm run dev` and the `dist:*` scripts go through
`scripts/run-electron-safe.mjs`, which clears the variable for the child process.

**The dev server starts but no window ever appears**

Two distinct causes, both silent. Check them in this order:

1. **Chromium cannot start its sandbox broker.** `electron.exe` dies during
   startup — before any code in `src/main/index.ts` runs — so there is no error
   output at all. Typical on restricted tokens, AppLocker/WDAC, and agent or CI
   sandboxes. Confirm with Task Manager (no `electron.exe` from this project)
   and launch without the sandbox instead:

   ```bash
   npm run dev:nosandbox                 # passes --no-sandbox to Electron
   # or, for the packaged app / any entry point:
   set ELECTRON_DISABLE_SANDBOX=1 && npm run dev
   ```

   The main process logs a `WARNING: the renderer sandbox is DISABLED` line so
   the reduced isolation is never silent. Use it as a workaround, not a default —
   with it on, the renderer keeps its IPC bridge and loses its OS-level isolation.

2. **The single-instance lock could not be taken.** `src/main/index.ts` exits
   when `app.requestSingleInstanceLock()` fails — including when another copy is
   already running, or the userData directory is not writable (locked ACL,
   restricted profile). It now logs the userData path it could not lock. To rule
   it out, run against a writable directory:
   `npm run dev -- -- --user-data-dir=<path>`.

If a window appears but stays blank, the main process reports why
(`did-fail-load`, `render-process-gone`, `unresponsive`) on stderr, and
`electron . --smoke-test` prints a one-line JSON report of whether the chrome
actually painted.

## Security model

- `contextIsolation` on, `nodeIntegration` off, Chromium sandbox on.
- The preload bridge (`window.fw`) exposes a fixed list of named operations and
  nothing else: no `require`, no `ipcRenderer`, no arbitrary channels.
- Every file-writing handler answers with a result object rather than rejecting,
  so a failed save can never hang the quit handshake or vanish silently.
- `file:open` honours an explicit path **only** when it is already in the recent
  files the user picked through a dialog; anything else falls back to the dialog.
- Every IPC handler checks that the calling frame is the app window's main frame.
- Navigation and `window.open` are denied at the `web-contents` level; the only
  way out to the desktop is `shell:openExternal`, which re-validates the URL
  scheme (`http`/`https`/`mailto`/`tel`) in the main process. Link clicks are
  Ctrl/Cmd+click, never a plain click.
- Permission requests (camera, clipboard, notifications, …) are refused.

## Cross-platform packaging / CI

`electron-builder` is configured for all three targets in `electron-builder.yml`:

- **Windows**: NSIS installer (`dist:win`), icon from `build/icon.png`
- **macOS**: DMG for x64 and arm64, unsigned (`dist:mac`)
- **Linux**: AppImage + deb (`dist:linux`), hicolor icon set in `build/icons/`

The GitHub Actions workflow in `.github/workflows/build.yml` runs
lint + format check + typecheck + tests on every push/PR, then packages each
target on its own native runner and uploads the installers as workflow artifacts:

```
┌────────────┐     ┌──────────────────────────────────────────────────┐
│  check     │ ──▶ │ dist (matrix: windows / macos / ubuntu)          │
│  lint +    │     │   npm ci → npm run dist:<target>                 │
│  format +  │     │   → smoke test the built app (xvfb on Linux)     │
│  typecheck │     │   → upload dist/* as an artifact                 │
│  + tests   │     └──────────────────────────────────────────────────┘
└────────────┘
```

The smoke step (`test/smoke-launch.test.ts`) launches the real Electron binary
with `--smoke-test`, which asserts the editor, toolbar, status bar and at least
one page are in the DOM, then exits — the mocked-Electron suites cannot catch a
build that packages green and never paints.

No code signing is configured; unsigned mac builds will show a Gatekeeper
warning on first launch.

## Project structure

```
src/
  main/            Electron main process
    index.ts       app lifecycle, window creation, navigation guards, --smoke-test
    menu.ts        native menu (File / Edit / View / Insert / Review / Help)
                   + right-click menu
    ipc.ts         IPC handlers (open / save / save-raw / PDF / print / image /
                   external URL / settings / workspace pages)
    docx-export.ts FreeWord JSON → .docx (docx library)
    recent.ts      recent-files persistence
    settings.ts    Notion mode / theme / sidebar persistence (userData/settings.json)
    workspace.ts   page tree store (userData/workspace/)
  preload/         contextBridge API exposed as window.fw
  shared/          types + the URL scheme policy shared across processes
  renderer/
    src/
      app.ts       document state, menu command dispatch, print mode,
                   workspace pages + autosave
      editor/
        extensions.ts  Tiptap schema (custom doc → page → blocks)
        schema.ts      doc / page / pageBreak nodes (doc attrs: page setup,
                       header/footer text, trackChanges flag, comment store)
        blocks.ts      task lists, callout, quote, divider, code block,
                       toggle list, table of contents
        footnotes.ts   footnoteRef inline atom + page-end note area
        review.ts      tracked-insert / tracked-delete marks + capture plugin
        comments.ts    the comment mark + thread store helpers
        mark-edit.ts   per-mark removal (accept/reject, delete comment)
        block-utils.ts shared block catalog (slash menu / turn-into / style select)
        code.ts, link.ts, hard-break.ts   inline primitives (hand-rolled)
        cursors.ts     drop + gap cursors
        pagination.ts  live pagination engine (blocks never split; pausable;
                       reserves the header/footer bands' height)
        create-editor.ts
      features/      toolbar, slash menu, bubble toolbar, block handle,
                     sidebar, page hero (cover/icon/title), quick switcher,
                     status bar, page-setup modal, find-replace panel,
                     header-footer modal, comments panel,
                     prompt/alert/confirm modals, zoom, image insert, link clicks,
                     appearance (Notion mode / theme)
      io/
        import-docx.ts  .docx → DOM (docx-preview) → FreeWord JSON
        import-other.ts .txt / .html → FreeWord JSON
        serialize.ts    editor state → JSON / text / HTML
test/              vitest suite (pagination, docx export/import, IPC, smoke launch)
scripts/           run-electron-safe.mjs, make-icon.py
build/             generated icons (see `npm run icons`)
```

### How .docx import works

`docx-preview` renders the binary .docx into a DOM. A normalization pass
converts its output (heading paragraphs, CSS-counter list paragraphs,
`<section class="docx">` pages) into plain HTML matching FreeWord's schema,
and ProseMirror's `DOMParser` turns that into the document JSON. Links are
kept when their scheme is one the app trusts and unwrapped (text preserved,
payload dropped) otherwise. Embedded images are inlined as base64 data URLs so
they round-trip back through export.

## Known limitations (v1.x)

- Legacy `.doc` (pre-2007) files are not supported — save them as `.docx` first.
- WebP images are not supported (use PNG / JPEG / GIF / BMP).
- A single block taller than one page (e.g. a very large table) is kept whole
  and may extend past the page in PDF export.
- Autosave / crash recovery exists for **workspace pages** only: an unsaved
  `.docx` file document still lives only in the window.
- To-do / callout / code-block fidelity in `.docx` is best-effort: Word itself
  sees checkbox glyphs, shaded paragraphs and monospace runs rather than the
  rich nodes (they re-hydrate when FreeWord re-opens its own export).
- **Import** of the review features is best-effort: header/footer text comes
  back with page-number fields rendered as their first-page number, footnotes
  re-hydrate from FreeWord's own export, and comments / Word-native revisions
  from other editors import as plain text. Files authored in Word or Google
  Docs keep all their body content either way.
- Track changes are character-level: deleting an image or a whole table
  (while tracking is on) is an immediate delete, not a revision.
- The note area at the bottom of a page is an overlay; a page filled edge to
  edge may need the margins widened (View → Page Setup) for the notes to sit
  clear of the text.

## License

MIT — see [LICENSE](LICENSE).
