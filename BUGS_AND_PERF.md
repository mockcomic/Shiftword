# Shiftword — Bug & Efficiency Audit

Full-codebase review (main process, preload, editor, features, io).
**Fixed** items ship with this change; **Open** items are recommendations with suggested fixes.
Verified via `npm run typecheck && npm run lint && npm test` (290 tests green).

## Fixed

### Main process
| # | Severity | Where | Bug → Fix |
|---|----------|-------|-----------|
| 1 | **High** | `src/main/workspace.ts` | Lost-update race: every mutator did `readStore()` → await → `writeStore()` from a stale snapshot, so an autosave flush in flight could silently delete a page a concurrent create/rename/move had added (and report success). All operations now serialize behind a module promise chain; regression test added (`does not lose updates when an autosave races a create/rename`). |
| 2 | **High** | `src/main/ipc.ts` | `settings:set` whitelist omitted `authorName` → Edit → Author Name… was a silent no-op (comments/track-changes always stamped the OS user). Now validated + length-capped. |
| 3 | Med (security) | `src/main/ipc.ts` | `docx:save` / `file:saveRaw` trusted a renderer-supplied `path` and wrote anywhere on disk. Renderer paths are now accepted only if main issued them (dialog result / opened file / recents); test added. |
| 4 | Med | `src/main/ipc.ts` | The 100 MB open cap was checked *after* `readFile` (a 2 GB file was fully buffered + IPC-cloned, then rejected); `file:pickImage` had the same shape. Now `stat` first, and the `stat` result feeds `upsertRecent` (one fewer round-trip). |
| 5 | Med (security) | `src/main/index.ts` | Packaged `will-navigate` guard was inert: `file:` URLs have opaque origins, so `origin` comparison let *any* local file pass, and `appUrl` was built as malformed `file://C:\…`. Now `pathToFileURL` + exact-href comparison; `will-frame-navigate` also guarded; `setPermissionCheckHandler` added (was only Request). |
| 6 | Med (security) | `src/main/index.ts` | `recent:get` was the only handler without the `trustedSender` main-frame check. Guarded. |
| 7 | Med (security) | `src/main/workspace.ts` | Page-id `"pages"` exclusion was case-sensitive — `restoreTrash("PAGES-t.json")` could clobber `pages.json` on case-insensitive Windows volumes. Now case-insensitive + basename belt-and-braces. |
| 8 | Med (bug+perf) | `src/main/workspace.ts` | `listTrash` truncated 200 entries in *readdir* order before sorting (newest deletions could be dropped), then read each file serially. Orders by the name timestamp first and reads with `Promise.all`. |
| 9 | Low | `src/main/workspace.ts` | `deletePage`: pages whose doc file was missing silently vanished (no trash entry); the catch-fallback `rename` could overwrite the wrapped trash entry with the bare doc; same-ms trash-name collisions. All fixed; size-cap now compares UTF-8 bytes (`Buffer.byteLength`), not UTF-16 units. |
| 10 | Low | `src/main/index.ts` | Quit handshake: a rejecting `showMessageBox` escaped as an unhandled rejection leaving `awaitingResponse` stuck `true` → the app could never be quit again. Wrapped in try/catch/finally + reentrancy guard + destroyed-window check before phase 2. |
| 11 | Low | `src/main/index.ts` | `second-instance` dropped `argv` — opening a .docx while running did nothing. Now routes the file through recents + `menu:command`. |
| 12 | Low | `src/main/menu.ts` | Recent-file labels passed raw to `buildFromTemplate`: `&` is a mnemonic ("Q&A draft.txt" rendered "QA draft.txt"). Escaped as `&&`. |
| 13 | Low | `src/main/menu.ts` | About dialog: `showMessageBox(..., detail)` is not an option (version never shown) and the floating promise was uncaught. Uses `message` + `.catch`. |
| 14 | Low | `src/main/recent.ts` / `settings.ts` | `upsertRecent` deduped case-*sensitively* (duplicate entries on Windows when Explorer casing differed) — now shares the case-insensitive identity. All three JSON stores (workspace/recent/settings) use tmp+`rename` atomic writes, and recent/settings serialize read-modify-write. `os.userInfo()` memoized (was a sync sysinfo call on every settings read). |
| 15 | Low | `src/main/docx-export.ts` | Per-export waste + hardening: the base64 regex captured the whole payload (extra copy); `imageSizeFromBase64` decoded the *entire* image just for header dims — now a 64 KB header slice. `walkPm` got a depth cap so a cyclic `payload.json` produces a clear error instead of a stack overflow. |

### Renderer
| # | Severity | Where | Bug → Fix |
|---|----------|-------|-----------|
| 16 | **High** | `src/renderer/src/io/import-docx.ts` | `replaceRunWithList` captured `first.nextSibling` **after** removing the run → null → imported lists jumped to the END of the section, reordering content after every list. Anchor now captured before removal. |
| 17 | **High (perf)** | `src/renderer/src/editor/pagination.ts` | Every keystroke re-measured EVERY block of EVERY page (`view.nodeDOM` + `getBoundingClientRect`), and the settle loop re-measured the full document after every dispatch. Block heights are now cached in a `WeakMap` keyed by the immutable PM node — only new/changed nodes measure; cache invalidates on zoom, theme/mode, font-load and window resize. |
| 18 | **High (bug)** | `src/renderer/src/editor/pagination.ts` | Measured height ignored CSS margins (`getBoundingClientRect` excludes them) while every block carries `margin: 0 0 6pt` → pages over-filled ~8 px/block past the 2 px tolerance. Margins now included. |
| 19 | High | `src/renderer/src/editor/footnotes.ts` | Footnote numbers went stale: NodeView `paint()` runs only when *its own* node updates, so deleting footnote 1 left the rest showing `[2]`, `[3]`… Live painters now register in a set the plugin's `view.update` re-runs whenever the collected footnote set changes. |
| 20 | High (perf) | `src/renderer/src/editor/footnotes.ts` | The page-end note-area widget was rebuilt on **every transaction** (fresh `toDOM` closure — prosemirror-view compares widgets by reference) → widget DOM churn per keystroke. The `DecorationSet` is now cached until the footnote set actually changes. |
| 21 | High | `src/renderer/src/editor/review.ts` | Cut in track-changes mode went through `beforeinput("cut")` → preventDefault suppressed the browser cut, and nothing wrote the clipboard: Ctrl+X *and* lost the text. The selection text is now copied before the soft delete. |
| 22 | Med | `src/renderer/src/app.ts` | `flushPageSave` had no serialization (blur / quit / switch can fire at once; two `writePage` calls could land out of order and the older snapshot win) — flushes now chain, and `pageUnsaved` clears only if the doc node is still the one that was written (no lost keystrokes if the user typed mid-write). |
| 23 | Med | `src/renderer/src/app.ts` | No re-entry guard on document switches: double-clicks / concurrent menu + sidebar switches interleaved (content of one page + title of another). One `switchBusy` guard across `newDoc`/`openFile`/`openPage`. |
| 24 | Med (bug/leak) | `src/renderer/src/features/modal.ts` | A second modal of the same kind orphaned the first's capture-phase keydown listener and left its Promise unresolved forever (caller hung; orphan handler stole Escape/Enter from the new dialog). Disposer registry resolves + detaches the preempted dialog. |
| 25 | Med (XSS/markup) | `modal.ts`, `quick-switcher.ts` | `title`/`message`/filenames/error strings and page titles/icons were interpolated into `innerHTML` unescaped (CSP blunts scripts, but markup could break the dialog). `escapeHtml` applied; quick-switcher now too. |
| 26 | Med (IME) | `modal.ts`, `quick-switcher.ts`, `find-replace.ts`, `slash-menu.ts` | Enter/Escape had no `isComposing`/`keyCode===229` guard — committing an IME candidate accepted dialogs, picked switcher rows, jumped find matches, inserted slash blocks. Guarded everywhere. |
| 27 | Med (bug) | `src/renderer/src/editor/blocks.ts` | Enter inside a task item rebuilt paragraphs from `textBetween` — silently destroying bold/italic/links/comment marks and inline atoms (hard breaks, footnote refs) at the split. The paragraph node is now cut at the caret, marks and all. |
| 28 | Med (bug) | `src/renderer/src/editor/block-utils.ts` | The `paragraph` and `heading` block defs passed *paragraph nodes* into `inline*` content via un-`.createChecked()` `.create()` → schema-invalid `p>p` / `h1>p` trees (broke HTML round-trips and docx export of that content). Now one real paragraph/heading per line with `schema.text`. |
| 29 | Med (perf) | `src/renderer/src/features/comments-panel.ts` | The rail re-rendered `innerHTML` + re-attached every card's listeners on **every keystroke** (and reset scroll). A store-signature guard skips unchanged re-renders; listeners are one delegated handler installed once. |
| 30 | Med (perf) | `src/renderer/src/features/bubble-toolbar.ts` | Rebuilt its popup (innerHTML + 16 listeners + two forced `coordsAtPos` reflows) on every transaction while a selection existed. Now diffs a (selection-range + active-marks) signature and no-ops when unchanged. |
| 31 | Med (perf) | `src/renderer/src/features/block-handle.ts` | `blockPosForDom` walked every block of every page calling `view.nodeDOM` on **every raw mousemove** — O(doc) per pixel. Per-element `WeakMap` cache (validated against `nodeDOM` on hit) reduces it to one scan per document state. |
| 32 | Low | `src/renderer/src/editor/extensions.ts` | `lineSpacing` parseHTML `parseFloat` turned an imported `line-height: 18px` into `line-height: 18` (18× spacing). px/em/rem/pt/% values and absurd ratios now fall back to 1. |
| 33 | Low | `src/renderer/src/editor/schema.ts` | The doc `comments` attr's `{ default: {} }` shared one mutable object across all docs created without the attr. Now a function default. |
| 34 | Low | `src/renderer/src/features/sidebar.ts` | Auto-expand tested `expanded.size === 0` *inside* its loop — after the first expansion every other root stayed collapsed. Flag hoisted before the loop. |
| 35 | Low | `src/renderer/src/features/page-setup.ts` | No keyboard handling at all: Escape did not close it and Enter did not confirm. Capture-phase handlers added (with IME guard). |
| 36 | Low | `src/renderer/src/features/find-replace.ts` | The panel's ✕ button left focus on `<body>` (Escape path refocused the editor; ✕ didn't), so editor keymaps looked dead. ✕ now returns focus. |

## Open (recommendations, not yet applied)

1. **`sleep(100)` print/PDF race** (`ipc.ts`) — main fixed-delays instead of awaiting a renderer ack; a slow machine can capture pre-print-mode layout. Suggest a `print:ready` IPC handshake.
2. **Undo-history batching contract** (`review.ts rejectAllChanges`, `comments.ts updateComment`) — dispatch 2 transactions where the file's own comment promises one undo per operation. Thread one `tr` through.
3. **Pagination selection mapping** (`pagination.ts makeMapper`) — maps by page-slot, not block identity: a selection on a block that moves pages can land in a sibling (or at doc start via the catch fallback). Map via the layout result's moves.
4. **Store caching** — `readRecent()` runs twice per open (dialog path check + upsert), `rebuildMenu()` re-reads both stores on every settings toggle, `readPage` has no size cap on the file it parses. In-memory caches + `stat` cap in `readPage`.
5. **`blocks.ts posOfRenderedNode`** stops descending on same-type matches (nested callouts unfindable) and scans O(doc) per click; prefer `view.posAtDOM`.
6. **`import-docx.ts`** — footnote refs pair to notes **positionally** (`notes[i]`), not by the `sup` number (wrong text when refs are out of order or a foreign `<ol>` is picked); `normalizeImages` silently drops WebP/SVG/relative-src images — keep a `[image missing]` placeholder instead.
7. **`import-other.ts` `htmlToDoc`** assigns untrusted HTML to `container.innerHTML` (image fetches start at parse time); use `DOMParser` + strip `script`/`object`/`iframe`.
8. **Latent listener accumulation** — `sidebar.ts:328` (document mousedown), `find-replace.ts:331` (module keydown), `appearance.ts:60` (matchMedia change): all fine with the current single init paths, but they leak if those are ever called more than once (e.g. a second `initAppearance`).
9. **`toolbar.ts syncToolbar`** runs ~35 selectors + `can()` dry-runs per debounced tick; cache element refs from `createToolbar`. `statusbar.ts updateStatus` re-runs doc-wide `textBetween` for the word count on every keystroke — cache per doc node like pagination heights do.
10. **`serialize.ts docJson`** `JSON.parse(JSON.stringify())` deep-copies the whole doc per autosave — reuse `editor.getJSON()`'s fresh copy (it is already a fresh tree) or serialize lazily.
11. **`schedulePaginate` module-global rAF handle** — harmless single-editor today; a second editor instance would starve on the first's schedule (use `WeakMap<Editor, number>`).
12. **`replaceCurrent`** (`find-replace.ts`) does not advance to the next match after replacing (Word does) — the next F3 can re-show the just-replaced match when the replacement contains the query.
13. **`pdf:export`** — after the save dialog `showSaveDialog` path selection, printToPDF runs silently; consider honoring `printBackground:false` for text-only docs (smaller PDFs).
14. **`menu.ts`** has no macOS `role: "appMenu"` — on darwin the standard app menu is missing.
15. **`trustedSender`** compares `win.webContents.mainFrame` — also verify `event.sender === win.webContents` (iframe-in-detached-devtools edge).
16. **No `process.on("unhandledRejection")`** in main — a single logger would surface any future floating rejection instead of silent loss.

## Not issues (verified)

- Link/URL policy (`shared/urls.ts`, `editor/link.ts`): protocol allow-list applied at import **and** render; `file:`/`javascript:` rejected.
- `mark-edit.ts` descending-order rewrites; pagination's `preventUpdate`/`addToHistory`/`createChecked`-abort guards; the `hasTrailingEmptyPage` heuristic; the toolbar dropdown and block-handle menu lifecycles; preload (minimal surface, working unsubscribers).
- `replaceAll` iterates matches back-to-front — safe (agent-reported stale-range concern disproven in review).
