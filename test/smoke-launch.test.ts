/**
 * Boot the *real* Electron app and prove it paints.
 *
 * The rest of the suite fakes the platform: `test/e2e-main-ipc.test.ts` mocks
 * the `electron` module, `test/e2e-app.test.ts` mocks `window.fw`. Nothing
 * there can catch the failure class the README documents at length — a build
 * that packages, type-checks and unit-tests green while the window never
 * paints, or a renderer bundle that throws on boot.
 *
 * This test drives the actual binary with `--smoke-test` (see
 * `src/main/index.ts`), which loads the window, asserts the editor / toolbar /
 * status bar / page DOM exist, prints one JSON line, and exits.
 *
 * It needs the built bundles in `out/`, so in CI it runs in the packaging job
 * after `npm run dist:<target>` — on Linux under `xvfb-run`.
 */
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { beforeAll, describe, expect, it } from "vitest";

const root = fileURLToPath(new URL("..", import.meta.url));
const wrapper = path.join(root, "scripts", "run-electron-safe.mjs");
const built = existsSync(path.join(root, "out", "main", "index.js"));

interface Run {
  status: number | null;
  stdout: string;
  stderr: string;
  error?: NodeJS.ErrnoException;
}

function runElectron(args: string[], env: Record<string, string> = {}): Run {
  const result = spawnSync(process.execPath, [wrapper, "electron", ...args], {
    cwd: root,
    encoding: "utf8",
    timeout: 120_000,
    env: { ...process.env, ...env },
  });
  return {
    status: result.status,
    stdout: (result.stdout ?? "").trim(),
    stderr: (result.stderr ?? "").trim(),
    error: (result as { error?: NodeJS.ErrnoException }).error,
  };
}

// Chromium cannot start in every environment (restricted tokens, no display,
// sandboxes without user namespaces). Gate on the cheapest probe: if the
// binary cannot even report its version, the platform is not available here
// and a failure would be noise rather than signal.
let canLaunch = false;
let probeDetail = "";
let userDataDir = "";

beforeAll(async () => {
  if (!built) {
    probeDetail = "no out/ build — run `npm run build` (or the dist job) first";
    return;
  }
  userDataDir = await mkdtemp(path.join(tmpdir(), "fw-smoke-"));
  const probe = runElectron(["--version"]);
  canLaunch = probe.status === 0 && /electron|v?\d+\.\d+/.test(probe.stdout + probe.stderr);
  if (!canLaunch) {
    probeDetail = `electron --version failed (${probe.status}): ${probe.stderr || probe.stdout || probe.error?.message}`;
  }
});

describe("packaged app smoke test", () => {
  it("reports why it could not run", () => {
    // `it.skipIf` cannot await the probe, so the skip reason is reported from a
    // test that always runs: a silent green here would hide an environment that
    // quietly never launched the app.
    if (built && canLaunch) {
      console.log("[smoke] electron launched; the boot assertion runs in the next test");
      expect(probeDetail).toBe("");
      return;
    }
    expect(probeDetail, "skip reason must be stated").toBeTruthy();
    console.log(`[smoke] skipped: ${probeDetail}`);
  });

  it.skipIf(!built || !canLaunch)("boots, paints the editor, and exits clean", () => {
    const run = runElectron([".", "--smoke-test", "--no-sandbox", `--user-data-dir=${userDataDir}`]);

    const line = run.stdout.split(/\r?\n/).find((l) => l.startsWith("SMOKE "));
    expect(line, `no SMOKE line on stdout.\nstdout: ${run.stdout}\nstderr: ${run.stderr}`).toBeTruthy();

    const report = JSON.parse(line!.slice("SMOKE ".length)) as {
      ok?: boolean;
      editor?: boolean;
      toolbar?: boolean;
      statusbar?: boolean;
      pages?: number;
      reason?: string;
    };

    // Every part of the chrome must be present, not merely "some" DOM.
    expect(report, `smoke report: ${JSON.stringify(report)}`).toMatchObject({
      ok: true,
      editor: true,
      toolbar: true,
      statusbar: true,
    });
    expect(Number(report.pages)).toBeGreaterThanOrEqual(1);
    expect(run.status, `exit ${run.status}; stderr: ${run.stderr}`).toBe(0);
  });
});
