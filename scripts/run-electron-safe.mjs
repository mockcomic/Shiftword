#!/usr/bin/env node
/**
 * Runs a local CLI package with ELECTRON_RUN_AS_NODE cleared.
 *
 * Why this exists: when ELECTRON_RUN_AS_NODE is set (npm inherits it from CI
 * runners, agent shells, and any parent Node process that was itself Electron),
 * `electron.exe` does not start the Electron runtime at all — it behaves as a
 * plain Node interpreter. In that mode `require("electron")` resolves the npm
 * stub, which exports the *path to the binary* as a string, so every named
 * import is undefined and the main process dies on its first top-level Electron
 * API call:
 *
 *   TypeError: Cannot read properties of undefined (reading 'on')
 *       at electron.ipcMain.on("app:quit-response", ...)
 *
 * The variable is not something this project sets, and requiring every
 * developer / runner to unset it by hand is not a fix. Clearing it here makes
 * `npm run dev` and the packaging scripts behave the same everywhere.
 *
 * Usage: node scripts/run-electron-safe.mjs <package-name> [...args]
 */
import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, join, resolve } from "node:path";
import process from "node:process";

delete process.env.ELECTRON_RUN_AS_NODE;

const [target, ...args] = process.argv.slice(2);

if (!target) {
  console.error("usage: node scripts/run-electron-safe.mjs <package-name> [...args]");
  process.exit(2);
}

const require = createRequire(import.meta.url);

/**
 * Resolve an installed package's bin entry to an absolute .js path.
 * Accepts the package name when its bin key matches, or a single-entry bin.
 */
function resolveBin(name) {
  const pkgJsonPath = require.resolve(`${name}/package.json`);
  const { bin } = require(pkgJsonPath);

  if (typeof bin === "string") return join(dirname(pkgJsonPath), bin);

  const keys = Object.keys(bin ?? {});
  const preferred = name.split("/").pop();
  const key = bin?.[name] ? name : bin?.[preferred] ? preferred : keys.length === 1 ? keys[0] : null;

  if (!key) {
    throw new Error(
      `package "${name}" has multiple bins (${keys.join(", ")}) — pass the file to run instead`,
    );
  }
  return join(dirname(pkgJsonPath), bin[key]);
}

let binPath;
try {
  binPath = existsSync(target) ? resolve(target) : resolveBin(target);
} catch (error) {
  console.error(error.message);
  process.exit(1);
}

const child = spawn(process.execPath, [binPath, ...args], {
  cwd: process.cwd(),
  env: process.env,
  stdio: "inherit",
});

for (const signal of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(signal, () => child.kill(signal));
}

child.on("error", (error) => {
  console.error(`failed to start ${target}: ${error.message}`);
  process.exit(1);
});

child.on("exit", (code, signal) => {
  process.exit(signal ? 1 : (code ?? 0));
});
