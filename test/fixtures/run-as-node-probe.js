// Probe child for scripts/run-electron-safe.mjs.
// Prints whether ELECTRON_RUN_AS_NODE survived into the child environment,
// followed by any extra CLI arguments the wrapper forwarded.
const value = process.env.ELECTRON_RUN_AS_NODE;
console.log(
  JSON.stringify({ runAsNode: value === undefined ? "unset" : `set:${value}`, args: process.argv.slice(2) }),
);
