import { fileURLToPath } from "node:url";
import { defineConfig } from "electron-vite";

const rel = (p: string): string => fileURLToPath(new URL(p, import.meta.url));

export default defineConfig({
  main: {
    build: {
      outDir: "out/main",
      rollupOptions: { input: rel("./src/main/index.ts") },
    },
  },
  preload: {
    build: {
      outDir: "out/preload",
      rollupOptions: { input: rel("./src/preload/index.ts") },
    },
  },
  renderer: {
    root: "src/renderer",
    build: {
      outDir: "out/renderer",
    },
  },
});
