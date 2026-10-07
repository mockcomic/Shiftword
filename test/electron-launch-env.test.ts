/**
 * Regression guard for the `npm run dev` launch path.
 *
 * With ELECTRON_RUN_AS_NODE inherited, electron.exe runs as a plain Node
 * interpreter: require("electron") returns the binary *path string*, every
 * named import is undefined, and the main process dies at its first top-level
 * Electron API call. scripts/run-electron-safe.mjs exists to clear it.
 */
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const wrapper = fileURLToPath(new URL("../scripts/run-electron-safe.mjs", import.meta.url));
const probe = fileURLToPath(new URL("./fixtures/run-as-node-probe.js", import.meta.url));

function runWrapper(args: string[], env: Record<string, string>) {
  const result = spawnSync(process.execPath, [wrapper, ...args], {
    encoding: "utf8",
    env: { ...process.env, ...env },
  });
  return { status: result.status, stdout: result.stdout.trim(), stderr: result.stderr.trim() };
}

describe("run-electron-safe", () => {
  it("clears ELECTRON_RUN_AS_NODE for the child process", () => {
    const { status, stdout } = runWrapper([probe], { ELECTRON_RUN_AS_NODE: "1" });

    expect(status).toBe(0);
    expect(JSON.parse(stdout).runAsNode).toBe("unset");
  });

  it("forwards CLI arguments to the child", () => {
    const { stdout } = runWrapper([probe, "dev", "--no-sandbox"], { ELECTRON_RUN_AS_NODE: "1" });

    expect(JSON.parse(stdout).args).toEqual(["dev", "--no-sandbox"]);
  });

  it("resolves an installed package bin by name", () => {
    // This is the exact call `npm run dev` makes, so it proves the wrapper
    // still reaches the electron-vite CLI (not just an explicit file path).
    const { status, stdout } = runWrapper(["electron-vite", "--version"], {});

    expect(status).toBe(0);
    expect(stdout).toMatch(/^electron-vite\/\d+\./);
  });

  it("rejects a package whose bin name is ambiguous, with a usable message", () => {
    const { status, stderr } = runWrapper(["typescript", "--version"], {});

    expect(status).toBe(1);
    expect(stderr).toMatch(/multiple bins \(tsc, tsserver\)/);
  });

  it("fails loudly for an unknown package", () => {
    const { status, stderr } = runWrapper(["@this/package-does-not-exist"], {});

    expect(status).toBe(1);
    expect(stderr).toMatch(/Cannot find module|does not declare a bin entry/);
  });
});
