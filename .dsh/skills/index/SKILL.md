---
name: index
description: Index the current workspace with the codebase-memory MCP (cbm) so graph-first tools (search_graph, trace_path, get_architecture, query_graph) have fresh data. Use when asked to /index, (re)build the codebase index, or refresh a stale index.
---

# Index Workspace with Codebase-Memory MCP

Index the current workspace repository using the `cbm` (codebase-memory-mcp) MCP server.

## Steps

1. **Resolve the workspace root.** Use the session working directory; the project root is the nearest ancestor containing `.git` (the same rule the harness uses). Pass the absolute native path (e.g. `C:\Users\you\Documents\GitHub\Proj`).

2. **Check for an existing project first.** Call `list_projects` (server `cbm`). If a project already exists for this repository, note its name and current stats so you can report what changed after re-indexing.

3. **Run the index.** Call `index_repository` with:
   - `repo_path`: the workspace root absolute path.
   - `mode`: `full` (default; nodes + edges + semantic extraction). Use `fast` (structure only, no semantic layer) only when the user asks for a quick structural pass, or for very large repositories where semantic extraction is too slow. Use `moderate` when the user asks to filter noise. Use `cross-repo-intelligence` only when the user explicitly wants to link this repo to other indexed services (`target_projects`).
   - `name`: pass only if the user requested a specific project name.

4. **Verify the result.** Call `index_status` for that project with `diagnostics: summary`. Report:
   - total nodes/edges and languages detected;
   - any coverage gaps it lists (files not indexed);
   - whether the index root matches the expected repository path (use `verbose: true` if the user is debugging index locations).

5. **Spot-check queryability.** Run one cheap graph query to prove the index works, e.g. `get_architecture` with `aspects: ["overview"]` or a small `search_graph` query for a symbol the user cares about.

## Reporting

Summarize in a few lines: project name, mode used, node/edge counts, languages, and any coverage gaps or warnings. If indexing failed, relay the exact error from the tool (e.g. invalid path, permission denied) — do not silently retry with a different mode.

## Notes

- Re-indexing the same path refreshes the existing project entry in place.
- If the user mentions a different repository than the workspace root, ask before indexing.
- Do not call `delete_project` as part of a normal index run; only when the user asks to reset a corrupt index.
