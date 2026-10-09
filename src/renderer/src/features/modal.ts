/**
 * Small modal primitives shared by the editor chrome.
 *
 * These exist so the editor never has to fall back on `window.alert` /
 * `window.confirm`: Electron runs those as *blocking* Chromium dialogs, which
 * freeze the renderer (and the quit handshake) while open, cannot be styled to
 * match the app, and cannot be driven from a test.
 */

export interface PromptOptions {
  title: string;
  label: string;
  /** Initial input value. */
  value?: string;
  placeholder?: string;
  /** Label of the confirm button. */
  okLabel?: string;
  /**
   * Render a `<textarea>` instead of a one-line input (footnote text,
   * comments, header/footer editing). Enter then inserts a newline;
   * Ctrl/Cmd+Enter accepts.
   */
  multiline?: boolean;
  /**
   * Return an error string to keep the dialog open and show it instead of
   * accepting the value.
   */
  validate?: (value: string) => string | null;
}

/** Escape text for interpolation into the template `innerHTML` below. */
export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Only one dialog of each kind is live at a time. A second `alertDialog`
 * preempting the first used to `remove()` the old overlay but leave its
 * capture-phase keydown handler attached and its Promise unresolved: the
 * awaiting caller hung forever and the orphan handler stole Escape/Enter
 * from the NEW dialog (it was registered first, so it ran first). Registering
 * a disposer makes preemption resolve the old Promise and detach its listener.
 */
const activeDisposers = new Map<string, () => void>();

function preempt(id: string): void {
  activeDisposers.get(id)?.();
  activeDisposers.delete(id);
}

function arm(id: string, dispose: () => void): void {
  activeDisposers.set(id, dispose);
}

function disarm(id: string, dispose: () => void): void {
  if (activeDisposers.get(id) === dispose) activeDisposers.delete(id);
}

/** True while an IME composition owns the keyboard: Enter then commits candidates, not dialogs. */
function isImeComposing(e: KeyboardEvent): boolean {
  return e.isComposing || e.keyCode === 229;
}

/**
 * Ask the user for a line (or a block, with `multiline`) of text. Resolves
 * with the entered value, or null when dismissed (Cancel, Escape, click on
 * the backdrop, or a newer dialog of the same kind).
 */
export function promptModal(opts: PromptOptions): Promise<string | null> {
  const { title, label, value = "", placeholder = "", okLabel = "OK", multiline = false, validate } = opts;

  return new Promise((resolve) => {
    preempt("fw-prompt");

    const field = multiline ? "textarea" : "input";
    const overlay = document.createElement("div");
    overlay.id = "fw-prompt";
    overlay.className = "fw-modal-overlay";
    overlay.innerHTML = `
      <div class="fw-modal" role="dialog" aria-modal="true" aria-label="${escapeHtml(title)}">
        <h2>${escapeHtml(title)}</h2>
        <div class="row">
          <label for="fw-prompt-input">${escapeHtml(label)}</label>
          <${field} id="fw-prompt-input"${multiline ? ' rows="4"' : ' type="text"'}></${field}>
        </div>
        <p class="fw-prompt-error" id="fw-prompt-error" role="alert" hidden></p>
        <div class="actions">
          <button id="fw-prompt-cancel" type="button">Cancel</button>
          <button id="fw-prompt-ok" type="button" class="primary">${escapeHtml(okLabel)}</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const input = overlay.querySelector<HTMLInputElement | HTMLTextAreaElement>("#fw-prompt-input")!;
    const error = overlay.querySelector<HTMLElement>("#fw-prompt-error")!;
    input.value = value;
    if (placeholder) input.placeholder = placeholder;

    let done = false;
    const dispose = (): void => {
      if (done) return;
      done = true;
      document.removeEventListener("keydown", onKey, true);
      disarm("fw-prompt", dispose);
      overlay.remove();
    };
    const finish = (result: string | null): void => {
      dispose();
      resolve(result);
    };

    const accept = (): void => {
      // Multiline keeps the user's line breaks; a single line never had any.
      const v = multiline ? input.value.replace(/\s+$/, "") : input.value.trim();
      const problem = validate ? validate(v) : null;
      if (problem) {
        error.textContent = problem;
        error.hidden = false;
        input.focus();
        return;
      }
      finish(v);
    };

    const onKey = (e: KeyboardEvent): void => {
      if (isImeComposing(e)) return;
      // Capture phase: the editor's own keymaps must not see the keys that
      // drive the dialog.
      if (e.key === "Escape") {
        e.preventDefault();
        e.stopPropagation();
        finish(null);
      } else if (e.key === "Enter" && (!multiline || e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        e.stopPropagation();
        accept();
      }
    };
    document.addEventListener("keydown", onKey, true);
    arm("fw-prompt", () => finish(null));

    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) finish(null);
    });
    overlay.querySelector("#fw-prompt-cancel")!.addEventListener("click", () => finish(null));
    overlay.querySelector("#fw-prompt-ok")!.addEventListener("click", accept);
    input.focus();
    input.select();
  });
}

/** Show a one-button message that does not block the renderer process. */
export function alertDialog(message: string, title = "FreeWord"): Promise<void> {
  return new Promise((resolve) => {
    preempt("fw-alert");

    const overlay = document.createElement("div");
    overlay.id = "fw-alert";
    overlay.className = "fw-modal-overlay";
    overlay.innerHTML = `
      <div class="fw-modal" role="alertdialog" aria-modal="true" aria-label="${escapeHtml(title)}">
        <h2>${escapeHtml(title)}</h2>
        <p class="fw-alert-message">${escapeHtml(message)}</p>
        <div class="actions">
          <button id="fw-alert-ok" type="button" class="primary">OK</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const ok = overlay.querySelector<HTMLButtonElement>("#fw-alert-ok")!;
    const dispose = (): void => {
      if (disposed) return;
      disposed = true;
      document.removeEventListener("keydown", onKey, true);
      disarm("fw-alert", dispose);
      overlay.remove();
    };
    let disposed = false;
    const finish = (): void => {
      dispose();
      resolve();
    };
    const onKey = (e: KeyboardEvent): void => {
      if (isImeComposing(e)) return;
      if (e.key === "Escape" || e.key === "Enter") {
        e.preventDefault();
        e.stopPropagation();
        finish();
      }
    };
    document.addEventListener("keydown", onKey, true);
    arm("fw-alert", finish);
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) finish();
    });
    ok.addEventListener("click", finish);
    ok.focus();
  });
}

/**
 * The unsaved-changes prompt (Save / Don't save / Cancel), Word-style.
 *
 * Replaces the old blocking `window.confirm("Discard unsaved changes…")`:
 * the native dialog froze the renderer (quit handshake included), could not
 * be styled to match the app, and offered no way to keep the work without
 * losing the operation the user was in the middle of.
 */
export function saveChangesModal(opts: {
  title: string;
  message: string;
  saveLabel?: string;
  discardLabel?: string;
}): Promise<"save" | "discard" | "cancel"> {
  const { title, message, saveLabel = "Save", discardLabel = "Don\u2019t save" } = opts;
  return new Promise((resolve) => {
    preempt("fw-save-changes");

    const overlay = document.createElement("div");
    overlay.id = "fw-save-changes";
    overlay.className = "fw-modal-overlay";
    overlay.innerHTML = `
      <div class="fw-modal" role="alertdialog" aria-modal="true" aria-label="${escapeHtml(title)}">
        <h2>${escapeHtml(title)}</h2>
        <p class="fw-alert-message">${escapeHtml(message)}</p>
        <div class="actions">
          <button id="fw-save-changes-cancel" type="button">Cancel</button>
          <button id="fw-save-changes-discard" type="button">${escapeHtml(discardLabel)}</button>
          <button id="fw-save-changes-save" type="button" class="primary">${escapeHtml(saveLabel)}</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    let done = false;
    const dispose = (): void => {
      if (done) return;
      done = true;
      document.removeEventListener("keydown", onKey, true);
      disarm("fw-save-changes", dispose);
      overlay.remove();
    };
    const finish = (value: "save" | "discard" | "cancel"): void => {
      dispose();
      resolve(value);
    };
    const onKey = (e: KeyboardEvent): void => {
      if (isImeComposing(e)) return;
      if (e.key === "Escape") {
        e.preventDefault();
        e.stopPropagation();
        finish("cancel");
      } else if (e.key === "Enter") {
        // Enter picks the safe option: saving can never lose work.
        e.preventDefault();
        e.stopPropagation();
        finish("save");
      }
    };
    document.addEventListener("keydown", onKey, true);
    arm("fw-save-changes", () => finish("cancel"));
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) finish("cancel");
    });
    overlay.querySelector("#fw-save-changes-cancel")!.addEventListener("click", () => finish("cancel"));
    overlay.querySelector("#fw-save-changes-discard")!.addEventListener("click", () => finish("discard"));
    const ok = overlay.querySelector<HTMLButtonElement>("#fw-save-changes-save")!;
    ok.addEventListener("click", () => finish("save"));
    ok.focus();
  });
}

/**
 * Yes/no without the blocking native dialog: resolves true on the primary
 * button, false on Escape / backdrop / the secondary button.
 */
export function confirmModal(opts: {
  title: string;
  message: string;
  okLabel?: string;
  danger?: boolean;
}): Promise<boolean> {
  const { title, message, okLabel = "OK", danger = false } = opts;
  return new Promise((resolve) => {
    preempt("fw-confirm");

    const overlay = document.createElement("div");
    overlay.id = "fw-confirm";
    overlay.className = "fw-modal-overlay";
    overlay.innerHTML = `
      <div class="fw-modal" role="alertdialog" aria-modal="true" aria-label="${escapeHtml(title)}">
        <h2>${escapeHtml(title)}</h2>
        <p class="fw-alert-message">${escapeHtml(message)}</p>
        <div class="actions">
          <button id="fw-confirm-cancel" type="button">Cancel</button>
          <button id="fw-confirm-ok" type="button" class="primary${danger ? " danger" : ""}">${escapeHtml(okLabel)}</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    let done = false;
    const dispose = (): void => {
      if (done) return;
      done = true;
      document.removeEventListener("keydown", onKey, true);
      disarm("fw-confirm", dispose);
      overlay.remove();
    };
    const finish = (value: boolean): void => {
      dispose();
      resolve(value);
    };
    const onKey = (e: KeyboardEvent): void => {
      if (isImeComposing(e)) return;
      if (e.key === "Escape") {
        e.preventDefault();
        e.stopPropagation();
        finish(false);
      } else if (e.key === "Enter") {
        e.preventDefault();
        e.stopPropagation();
        finish(true);
      }
    };
    document.addEventListener("keydown", onKey, true);
    arm("fw-confirm", () => finish(false));
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) finish(false);
    });
    overlay.querySelector("#fw-confirm-cancel")!.addEventListener("click", () => finish(false));
    const ok = overlay.querySelector<HTMLButtonElement>("#fw-confirm-ok")!;
    ok.addEventListener("click", () => finish(true));
    ok.focus();
  });
}
