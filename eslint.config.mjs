import js from "@eslint/js";
import tseslint from "typescript-eslint";
import globals from "globals";

export default tseslint.config(
  {
    ignores: ["out/**", "dist/**", "node_modules/**", "coverage/**", ".tmp/**", "test/fixtures/**"],
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    // The renderer touches the DOM; main/preload do not, and a stray `window`
    // reference in the main process is the bug this split catches.
    files: ["src/renderer/**/*.ts"],
    languageOptions: {
      globals: { ...globals.browser },
    },
  },
  {
    files: ["src/main/**/*.ts", "src/preload/**/*.ts", "scripts/**/*.mjs"],
    languageOptions: {
      globals: { ...globals.node },
    },
  },
  {
    files: ["test/**/*.ts"],
    languageOptions: {
      globals: { ...globals.node, ...globals.browser },
    },
    rules: {
      // Tests cast freely to reach internals; the type gymnastics are the point.
      "@typescript-eslint/no-explicit-any": "off",
      "@typescript-eslint/no-non-null-assertion": "off",
    },
  },
  {
    rules: {
      // `console` is this app's only diagnostic channel: the main process logs
      // sandbox/quit/load failures there, and the renderer reports save errors.
      "no-console": "off",
      eqeqeq: ["error", "always", { null: "ignore" }],
      "no-var": "error",
      "prefer-const": "error",
      "@typescript-eslint/consistent-type-imports": ["error", { prefer: "type-imports" }],
      "@typescript-eslint/no-unused-vars": [
        "error",
        {
          // `_schema`, `_event`, …: deliberately unused parameters are how the
          // code says "this position matters, the value does not".
          argsIgnorePattern: "^_",
          varsIgnorePattern: "^_",
          caughtErrorsIgnorePattern: "^_",
          destructuredArrayIgnorePattern: "^_",
        },
      ],
    },
  },
);
