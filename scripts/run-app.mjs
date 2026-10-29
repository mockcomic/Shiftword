#!/usr/bin/env node
/**
 * Run Shiftword with Electron — no packaging, no manual build step.
 *
 * `npm run dev` serves the renderer through the Vite dev server (hot reload);
 * `dist:*` scripts run electron-builder and produce installers. Both are not
 * "just run the app". This script boots the *real* app from `out/` (the same
 * bundles the packaged app loads — `main` in package.json points there), and
 * takes care of the footguns on the way:
 *
 * 1. `out/` may be missing or stale (edited src since the last build), so it
 *    runs `electron-vite build` first — but only when something is actually
 *    out of date. A fresh checkout or CI artifact with current bundles skips
 *    the rebuild entirely. The build uses the workspace `.tmp/` as the temp
 *    directory (confined sandboxes deny the esbuild binary's own cleanup in
 *    the system temp — see `useWorkspaceTemp`).
 * 2. An inherited ELECTRON_RUN_AS_NODE makes electron.exe behave as a plain
 *    Node interpreter (the failure documented in the README troubleshooting
 *    section), so it is cleared for the child like run-electron-safe.mjs does
 *    for the dev/packaging scripts.
 *
 * Usage:
 *   node scripts/run-app.mjs                  # build only when stale, then run
 *   node scripts/run-app.mjs --build          # force a rebuild first
 *   node scripts/run-app.mjs --skip-build     # never build (run out/ as-is)
 *   node scripts/run-app.mjs --smoke-test     # any other flag goes to Electron
 *                                             # and the app (e.g. --no-sandbox)
 */
import { spawn, spawnSync } from "node:child_process";
import { existsSync, mkdirSync, readdirSync, statSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

// Must be deleted before any spawn: the child inherits the (cleaned) env.
delete process.env.ELECTRON_RUN_AS_NODE;

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);

/** Files whose presence defines "a build exists" — mirrors electron.vite.config.ts. */
const OUTPUTS = ["out/main/index.js", "out/preload/index.js", "out/renderer/index.html"];

/** Inputs that invalidate that build when newer than it. */
const SOURCES = ["src", "electron.vite.config.ts", "package.json"];

function mtimeMs(target) {
  return statSync(target).mtimeMs;
}

/** Newest mtime inside `target` (recursive, bounded; missing paths score -1). */
function newestMtime(target, depth = 0) {
  if (!existsSync(target)) return -1;
  const st = statSync(target);
  if (!st.isDirectory()) return st.mtimeMs;
  if (depth > 12) return st.mtimeMs;
  let newest = 0;
  for (const entry of readdirSync(target, { withFileTypes: true })) {
    if (entry.name === "node_modules" || entry.name.startsWith(".")) continue;
    newest = Math.max(newest, newestMtime(path.join(target, entry.name), depth + 1));
  }
  return newest;
}

function buildNeeded() {
  if (OUTPUTS.some((o) => !existsSync(path.join(ROOT, o)))) return true;
  const built = Math.max(...OUTPUTS.map((o) => mtimeMs(path.join(ROOT, o))));
  const source = Math.max(...SOURCES.map((s) => newestMtime(path.join(ROOT, s))));
  return source > built;
}

/**
 * Build inside `.tmp/` (gitignored) instead of the system temp directory.
 * Confined agent/CI sandboxes let the parent Node process touch %TEMP% but
 * deny the esbuild binary it spawns — cleanup of esbuild's own scratch dirs
 * then aborts the renderer build ("remove …: Access is denied"). A workspace
 * temp is always writable in those environments and costs nothing in normal
 * ones (vite/esbuild happily build against any TMPDIR; leftover cleanup stays
 * best-effort).
 */
function useWorkspaceTemp() {
  const dir = path.join(ROOT, ".tmp");
  try {
    mkdirSync(dir, { recursive: true });
    process.env.TMP = process.env.TEMP = process.env.TMPDIR = dir;
  } catch {
    /* keep the system temp — the build may still work there */
  }
}

const argv = process.argv.slice(2);
const flags = new Set(["--build", "--skip-build"]);
const force = argv.includes("--build");
const skip = argv.includes("--skip-build") && !force;
const passThrough = argv.filter((a) => !flags.has(a));

if (!skip && (force || buildNeeded())) {
  console.log("[run-app] building main/preload/renderer with electron-vite…");
  useWorkspaceTemp();
  const build = spawnSync(
    process.execPath,
    [path.join(ROOT, "scripts", "run-electron-safe.mjs"), "electron-vite", "build"],
    {
      cwd: ROOT,
      stdio: "inherit",
    },
  );
  if (build.status !== 0) {
    console.error(`[run-app] build failed${build.error ? `: ${build.error.message}` : ""}`);
    process.exit(build.status ?? 1);
  }
}

// Outside the Electron runtime the `electron` package's main export is the
// absolute path to the real binary — exactly what spawn needs here.
let electronPath;
try {
  electronPath = require("electron");
} catch (err) {
  console.error(`[run-app] could not resolve the Electron binary: ${err.message}`);
  process.exit(1);
}
if (typeof electronPath !== "string" || !existsSync(electronPath)) {
  console.error("[run-app] Electron is not installed — run `npm install` first.");
  process.exit(1);
}

/**
 * Launch Electron once; resolves with its outcome (exit status + lifetime).
 * stderr is teed (shown AND scanned) so exit causes that print one line and
 * leave — like the userData lock denial — can be recognised and reported.
 */
function launch(args) {
  return new Promise((resolve) => {
    const startedAt = Date.now();
    const child = spawn(electronPath, [ROOT, ...args], {
      cwd: ROOT,
      stdio: ["inherit", "inherit", "pipe"],
      env: process.env,
    });
    let stderr = "";
    child.stderr.setEncoding("utf8");
    child.stderr.on("data", (chunk) => {
      stderr += chunk;
      process.stderr.write(chunk);
    });
    for (const signal of ["SIGINT", "SIGTERM", "SIGHUP"]) {
      process.on(signal, () => child.kill(signal));
    }
    child.on("error", (error) => {
      console.error(`[run-app] failed to start Electron: ${error.message}`);
      resolve({ status: 1, ms: Date.now() - startedAt, stderr });
    });
    child.on("exit", (code, signal) => {
      resolve({ status: signal ? 1 : (code ?? 0), ms: Date.now() - startedAt, signal, stderr });
    });
  });
}

const sandboxRequested = process.env.ELECTRON_DISABLE_SANDBOX === "1" || passThrough.includes("--no-sandbox");
const userDataRequested = passThrough.some((a) => a.startsWith("--user-data-dir"));

const args = passThrough.slice();
let result = await launch(args);

// The README's first troubleshooting case: Chromium cannot create its sandbox
// broker on locked-down environments, and electron.exe then dies during
// startup — before any project code prints anything. The user sees a silent
// return to the prompt. Recognise the signature (nonzero exit within a few
// seconds) and retry once with the documented escape hatch, so `npm start`
// works everywhere — with the sandbox warning the app itself prints.
if (result.status !== 0 && result.ms < 5000 && !sandboxRequested) {
  console.log(
    "[run-app] Electron exited immediately — the Chromium sandbox broker likely could not " +
      "start here (see README Troubleshooting). Retrying with the renderer sandbox disabled…",
  );
  process.env.ELECTRON_DISABLE_SANDBOX = "1";
  result = await launch(args);
}

// The README's second silent case: %APPDATA% is not writable here, so the
// singleton lock file cannot be created and the main process exits cleanly as
// a phantom "second instance" (the exit signature that matches a real second
// instance — the Chromium ERROR line tells them apart). For the unwritable
// profile, relaunch against a writable project-local one and say so loudly.
if (
  result.status === 0 &&
  result.ms < 5000 &&
  !userDataRequested &&
  /Lock file can not be created/i.test(result.stderr)
) {
  const dir = path.join(ROOT, ".tmp", "user-data");
  console.log(
    `[run-app] the default Electron profile is not writable here. Retrying with a ` +
      `project-local one: ${dir}\n` +
      "            (settings and workspace pages created in this run live there, not in %APPDATA%).",
  );
  args.push(`--user-data-dir=${dir}`);
  result = await launch(args);
}

process.exit(result.status);
