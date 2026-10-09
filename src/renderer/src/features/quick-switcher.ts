import type { PageMeta } from "../../../shared/types";
import { filterPages } from "./sidebar";
import { escapeHtml } from "./modal";

/**
 * Ctrl+K quick switcher: a centered search box over the page list, Notion's
 * other signature keystroke. Enter opens, Esc / backdrop dismisses — same
 * capture-phase key handling as features/modal.ts, so the editor keymaps never
 * see these keys.
 */
export function openQuickSwitcher(pages: PageMeta[], onPick: (id: string) => void): void {
  document.getElementById("fw-quick-switcher")?.remove();
  if (pages.length === 0) return;

  const overlay = document.createElement("div");
  overlay.id = "fw-quick-switcher";
  overlay.className = "fw-modal-overlay quick-overlay";
  overlay.innerHTML = `
    <div class="fw-modal quick-modal" role="dialog" aria-modal="true" aria-label="Quick switcher">
      <input id="quick-input" type="text" placeholder="Search pages…" autocomplete="off" />
      <div id="quick-list" class="quick-list"></div>
    </div>
  `;
  document.body.appendChild(overlay);

  const input = overlay.querySelector<HTMLInputElement>("#quick-input")!;
  const list = overlay.querySelector<HTMLElement>("#quick-list")!;
  let current: PageMeta[] = [];
  let active = 0;

  const renderList = (): void => {
    if (current.length === 0) {
      list.innerHTML = `<div class="quick-empty">No matching pages</div>`;
      return;
    }
    // Titles come from user-created page names — they must not steer markup.
    list.innerHTML = current
      .map(
        (p, i) =>
          `<button type="button" class="quick-item${i === active ? " active" : ""}" data-id="${escapeHtml(p.id)}">` +
          `<span class="side-icon">${escapeHtml(p.icon || "📄")}</span>` +
          `<span class="quick-title">${escapeHtml(p.title)}</span></button>`,
      )
      .join("");
    list.querySelectorAll<HTMLButtonElement>(".quick-item").forEach((btn) => {
      btn.addEventListener("mousedown", (e) => e.preventDefault());
      btn.addEventListener("click", () => finish(btn.dataset.id ?? ""));
    });
  };

  const setQuery = (q: string): void => {
    current = filterPages(pages, q);
    active = 0;
    renderList();
  };

  let done = false;
  const finish = (id: string | null): void => {
    if (done) return;
    done = true;
    document.removeEventListener("keydown", onKey, true);
    overlay.remove();
    if (id) onPick(id);
  };

  const onKey = (e: KeyboardEvent): void => {
    // While an IME composition is open, Enter commits candidates and Escape
    // cancels the composition — neither may pick or close the switcher.
    if (e.isComposing || e.keyCode === 229) return;
    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      finish(null);
    } else if (e.key === "ArrowDown") {
      e.preventDefault();
      active = Math.min(current.length - 1, active + 1);
      renderList();
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      active = Math.max(0, active - 1);
      renderList();
    } else if (e.key === "Enter") {
      e.preventDefault();
      e.stopPropagation();
      finish(current[active]?.id ?? null);
    }
  };
  document.addEventListener("keydown", onKey, true);

  input.addEventListener("input", () => setQuery(input.value));
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) finish(null);
  });

  setQuery("");
  input.focus();
}
