import { app, BrowserWindow, dialog, ipcMain, session } from "electron";
import { stat } from "node:fs/promises";
import { basename, join } from "node:path";
import { pathToFileURL } from "node:url";
import { registerIpc } from "./ipc";
import { upsertRecent } from "./recent";
import { buildContextMenu, configureMenu, rebuildMenu } from "./menu";

let mainWindow: BrowserWindow | null = null;

/**
 * `electron . --smoke-test` boots the *real* app, asserts the window painted,
 * and exits. The unit/e2e suites mock Electron (`test/e2e-main-ipc.test.ts`)
 * and the preload API (`test/e2e-app.test.ts`), so nothing else proves the
 * packaged binary starts, loads the bundle, and lays out at least one page —
 * which is exactly the failure class the troubleshooting section documents.
 */
const SMOKE_TEST = process.argv.includes("--smoke-test");
const SMOKE_TIMEOUT_MS = 20_000;

/** The URL this app loaded; only navigations that stay on it are allowed. */
let appUrl: string | null = null;

// ---------------------------------------------------------------------------
// Sandbox policy
//
// On locked-down Windows environments (restricted tokens, AppLocker/WDAC, some
// agent and CI sandboxes) Chromium cannot create its sandbox broker. The
// process then dies during startup — before any code in this file runs — with
// no output at all, which is indistinguishable from "the app never opens a
// window". Because it dies that early, only Electron's own pre-startup
// ELECTRON_DISABLE_SANDBOX escape hatch helps; a project-local flag read here
// would be too late to matter. It is off by default, so the sandbox — which
// this app relies on together with contextIsolation — stays on.
// ---------------------------------------------------------------------------
const sandboxDisabled = process.env.ELECTRON_DISABLE_SANDBOX === "1";

if (sandboxDisabled) {
  app.commandLine.appendSwitch("no-sandbox");
  console.warn(
    "[FreeWord] WARNING: the renderer sandbox is DISABLED (ELECTRON_DISABLE_SANDBOX=1)." +
      " Only use this to work around an environment that cannot start Chromium's sandbox broker.",
  );
}

// ---------------------------------------------------------------------------
// Quit handshake (unsaved changes): main asks the renderer, the renderer
// answers with its dirty state; main decides Save / Discard / Cancel.
// ---------------------------------------------------------------------------
let isQuitting = false;
let awaitingResponse = false;
let pendingSaveQuit = false;
let quitTimer: NodeJS.Timeout | null = null;

function clearQuitTimer(): void {
  if (quitTimer) {
    clearTimeout(quitTimer);
    quitTimer = null;
  }
}

function forceQuit(): void {
  isQuitting = true;
  awaitingResponse = false;
  pendingSaveQuit = false;
  clearQuitTimer();
  app.quit();
}

function requestQuit(): void {
  if (isQuitting) {
    app.quit();
    return;
  }
  const win = mainWindow;
  if (!win || win.isDestroyed() || win.webContents.isDestroyed()) {
    forceQuit();
    return;
  }
  if (awaitingResponse) return;
  awaitingResponse = true;
  pendingSaveQuit = false;
  win.webContents.send("app:will-quit");
  // Safety net: if the renderer crashed we must not hang the app.
  quitTimer = setTimeout(() => {
    if (awaitingResponse && !pendingSaveQuit) forceQuit();
  }, 8000);
}

async function handleQuitResponse(
  payload: { dirty: boolean; name?: string } | null | undefined,
): Promise<void> {
  if (!awaitingResponse || handlingQuitResponse) return;
  handlingQuitResponse = true;
  try {
    await handleQuitResponseInner(payload);
  } catch (err) {
    // A rejecting dialog (window destroyed mid-dialog) used to escape as an
    // unhandled rejection with `awaitingResponse` stuck true — after that the
    // app could never be quit again from the close button or the menu.
    console.error("[FreeWord] quit handshake failed:", err);
    awaitingResponse = false;
    pendingSaveQuit = false;
    clearQuitTimer();
  } finally {
    handlingQuitResponse = false;
  }
}

async function handleQuitResponseInner(
  payload: { dirty: boolean; name?: string } | null | undefined,
): Promise<void> {
  clearQuitTimer();

  // Phase 2: response to `app:save-then-quit`.
  if (pendingSaveQuit) {
    pendingSaveQuit = false;
    awaitingResponse = false;
    if (!payload?.dirty) forceQuit();
    // Still dirty (save was cancelled) → stay alive.
    return;
  }

  // Phase 1: response to `app:will-quit`.
  if (!payload?.dirty) {
    forceQuit();
    return;
  }
  const win = mainWindow;
  if (!win || win.isDestroyed()) {
    forceQuit();
    return;
  }
  const { response } = await dialog.showMessageBox(win, {
    type: "warning",
    title: "Unsaved changes",
    message: `"${payload?.name ?? "Document"}" has unsaved changes.`,
    detail: "Do you want to save the changes before closing FreeWord?",
    buttons: ["Save", "Discard", "Cancel"],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
  });
  if (response === 0) {
    // Save → tell the renderer to save, then respond again.
    if (!mainWindow || mainWindow.isDestroyed() || mainWindow.webContents.isDestroyed()) {
      awaitingResponse = false;
      return;
    }
    pendingSaveQuit = true;
    mainWindow.webContents.send("app:save-then-quit");
    // The user may take a while in the Save As dialog: be patient, but not
    // forever — an ack timeout returns the app to a state where quit works.
    quitTimer = setTimeout(() => {
      if (awaitingResponse && pendingSaveQuit) {
        awaitingResponse = false;
        pendingSaveQuit = false;
        // Stay alive; the user can quit again later.
      }
    }, 300000);
  } else if (response === 1) {
    forceQuit();
  } else {
    awaitingResponse = false;
    // Cancel → stay alive.
  }
}

let handlingQuitResponse = false;

ipcMain.on("app:quit-response", (_event, payload: { dirty: boolean; name?: string }) => {
  void handleQuitResponse(payload);
});

// ---------------------------------------------------------------------------
// Window + lifecycle
// ---------------------------------------------------------------------------

/** What `--smoke-test` prints on stdout, one JSON object per line. */
const SMOKE_PROBE = `(() => {
  const q = (s) => document.querySelector(s);
  const bar = q("#statusbar");
  return {
    editor: !!q(".ProseMirror"),
    toolbar: !!q("#toolbar .tb-group"),
    statusbar: !!bar && /Page\\s+\\d+\\s+\\/\\s+\\d+/.test(bar.textContent || ""),
    pages: document.querySelectorAll(".fw-page").length,
    words: !!bar && /\\bwords\\b/.test(bar.textContent || ""),
  };
})()`;

async function runSmokeCheck(win: BrowserWindow): Promise<void> {
  try {
    const state = (await win.webContents.executeJavaScript(SMOKE_PROBE, true)) as Record<string, unknown>;
    const ok =
      Boolean(state.editor) && Boolean(state.toolbar) && Boolean(state.statusbar) && Number(state.pages) >= 1;
    console.log(`SMOKE ${JSON.stringify({ ok, ...state })}`);
    // `app.exit` bypasses the unsaved-changes handshake entirely.
    app.exit(ok ? 0 : 1);
  } catch (err) {
    console.log(
      `SMOKE ${JSON.stringify({ ok: false, reason: err instanceof Error ? err.message : String(err) })}`,
    );
    app.exit(1);
  }
}

function blockHostileContents(): void {
  app.on("web-contents-created", (_event, contents) => {
    // Nothing in this app opens a window: a document-embedded redirect, a
    // stray `window.open` or a link click that escaped the renderer's own
    // handler must never produce a second, un-guarded renderer.
    contents.setWindowOpenHandler(({ url }) => {
      console.warn(`[FreeWord] blocked window.open → ${url}`);
      return { action: "deny" as const };
    });

    contents.on("will-navigate", (event, url) => {
      // Content is loaded exactly once (dev server or bundled file). A
      // navigation away from that URL would replace the editor with a page.
      if (!isAllowedNavigation(url)) {
        event.preventDefault();
        console.warn(`[FreeWord] blocked navigation to ${url}`);
      }
    });
    // Nested frames navigate on their own channel; guard it the same way.
    contents.on("will-frame-navigate", (event) => {
      if (!isAllowedNavigation(event.url)) {
        event.preventDefault();
        console.warn(`[FreeWord] blocked frame navigation to ${event.url}`);
      }
    });

    contents.on("will-attach-webview", (event) => event.preventDefault());
  });

  // An offline editor needs no camera, microphone, clipboard read, geolocation
  // or notification: every permission request is refused, not silently granted
  // by a default that was written for a browser.
  session.defaultSession.setPermissionRequestHandler((_wc, permission, callback) => {
    console.warn(`[FreeWord] denied permission request: ${permission}`);
    callback(false);
  });
  // The *check* channel is separate from the *request* one — without this,
  // synchronous checks (alert(), clipboard-read, …) still fall through to
  // Electron's allow-by-default answer.
  session.defaultSession.setPermissionCheckHandler((_wc, permission) => {
    console.warn(`[FreeWord] denied permission check: ${permission}`);
    return false;
  });
}

/**
 * Is this navigation staying on the app's own document?
 *
 * `file:` URLs have *opaque* origins — `new URL(a).origin === "null"` for
 * every one of them — so an origin comparison would let ANY local file pass
 * as "same origin". For the packaged app the only allowed target is the
 * exact index.html the app itself loaded.
 */
function isAllowedNavigation(url: string): boolean {
  if (appUrl === null) return url.startsWith("file:");
  try {
    return isSameOrigin(appUrl, url);
  } catch {
    return false;
  }
}

function isSameOrigin(base: string, url: string): boolean {
  const a = new URL(base);
  const b = new URL(url);
  if (a.protocol === "file:" || b.protocol === "file:") return a.href === b.href;
  return a.origin === b.origin;
}

function createWindow(): void {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 900,
    title: "Untitled — FreeWord",
    webPreferences: {
      preload: join(__dirname, "../preload/index.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: !sandboxDisabled,
      spellcheck: true,
    },
  });

  // A window that is created but never paints looks exactly like a window that
  // never appeared, so record what actually happened to it.
  mainWindow.webContents.on("did-fail-load", (_event, code, description, url) => {
    console.error(`[FreeWord] page load failed (${code} ${description}) for ${url}`);
  });
  mainWindow.webContents.on("render-process-gone", (_event, details) => {
    console.error(`[FreeWord] renderer process gone: ${details.reason} (exit ${details.exitCode})`);
  });
  mainWindow.on("unresponsive", () => {
    console.error("[FreeWord] window is not responding");
  });

  // Route every window close through the quit handshake.
  mainWindow.on("close", (event) => {
    if (isQuitting) return;
    event.preventDefault();
    requestQuit();
  });

  // Right-click: Electron installs no menu of its own, so Cut/Copy/Paste and —
  // more importantly — the spell-checker suggestions have to come from here.
  mainWindow.webContents.on("context-menu", (event, params) => {
    event.preventDefault();
    const win = mainWindow;
    if (!win || win.isDestroyed()) return;
    buildContextMenu(win, params).popup({ window: win });
  });

  // `--smoke-test` is the packaged-app boot check the CI job runs: prove the
  // window actually paints, then exit before the developer/user sees anything.
  if (SMOKE_TEST) {
    mainWindow.webContents.once("did-finish-load", () => {
      void runSmokeCheck(mainWindow as BrowserWindow);
    });
    setTimeout(() => {
      console.error('SMOKE {"ok":false,"reason":"timed out waiting for the window to load"}');
      app.exit(1);
    }, SMOKE_TIMEOUT_MS).unref();
  }

  mainWindow.on("closed", () => {
    mainWindow = null;
  });

  if (process.env.ELECTRON_RENDERER_URL) {
    appUrl = process.env.ELECTRON_RENDERER_URL;
    void mainWindow.loadURL(appUrl);
  } else {
    const htmlPath = join(__dirname, "../renderer/index.html");
    // `file://${join(…)}` produces `file://C:\Users\…` on Windows — a
    // drive-name-as-authority URL that URL parsing treats as an empty
    // authority, so it never compared equal to real file URLs.
    appUrl = pathToFileURL(htmlPath).href;
    void mainWindow.loadFile(htmlPath);
  }
}

if (!app.requestSingleInstanceLock()) {
  // A non-writable userData directory (locked ACL, restricted profile) fails
  // here too, and without this line the app exits silently — indistinguishable
  // from "the app refuses to launch".
  console.warn(
    "[FreeWord] another instance already holds the single-instance lock — exiting." +
      ` (userData: ${app.getPath("userData")})`,
  );
  app.quit();
} else {
  app.on("second-instance", (_event, argv) => {
    const win = mainWindow;
    if (!win) return;
    if (win.isMinimized()) win.restore();
    win.focus();
    // Route the file the user double-clicked (or opened from a second
    // launch) to the running instance instead of dropping argv on the floor.
    const filePath = argv
      .slice(1)
      .find((a) => /\.(docx|txt|html|htm)$/i.test(a));
    if (!filePath) return;
    void (async () => {
      try {
        const st = await stat(filePath);
        await upsertRecent({ path: filePath, name: basename(filePath), mtimeMs: st.mtimeMs });
      } catch {
        return; // not a readable file — keep it out of the recent store
      }
      if (!win.isDestroyed() && !win.webContents.isDestroyed()) {
        win.webContents.send("menu:command", { id: "recent:open", path: filePath });
      }
    })();
  });

  void app.whenReady().then(() => {
    blockHostileContents();
    createWindow();
    registerIpc(
      () => mainWindow,
      () => {
        void rebuildMenu();
      },
    );
    configureMenu(
      () => mainWindow,
      () => requestQuit(),
    );
    void rebuildMenu();
    app.on("activate", () => {
      if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
  });

  app.on("before-quit", (event) => {
    if (isQuitting) return;
    event.preventDefault();
    requestQuit();
  });

  app.on("window-all-closed", () => {
    // macOS keeps the app resident with no windows — that is what the
    // `activate` handler above exists for, and quitting here made reopening
    // from the Dock impossible.
    if (process.platform === "darwin") return;
    requestQuit();
  });
}
