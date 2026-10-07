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
   * Return an error string to keep the dialog open and show it instead of
   * accepting the value.
   */
  validate?: (value: string) => string | null;
}

/**
 * Ask the user for one line of text. Resolves with the entered value, or null
 * when dismissed (Cancel, Escape, click on the backdrop).
 */
export function promptModal(opts: PromptOptions): Promise<string | null> {
  const { title, label, value = "", placeholder = "", okLabel = "OK", validate } = opts;

  return new Promise((resolve) => {
    document.getElementById("fw-prompt")?.remove();

    const overlay = document.createElement("div");
    overlay.id = "fw-prompt";
    overlay.className = "fw-modal-overlay";
    overlay.innerHTML = `
      <div class="fw-modal" role="dialog" aria-modal="true" aria-label="${title}">
        <h2>${title}</h2>
        <div class="row">
          <label for="fw-prompt-input">${label}</label>
          <input id="fw-prompt-input" type="text">
        </div>
        <p class="fw-prompt-error" id="fw-prompt-error" role="alert" hidden></p>
        <div class="actions">
          <button id="fw-prompt-cancel" type="button">Cancel</button>
          <button id="fw-prompt-ok" type="button" class="primary">${okLabel}</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const input = overlay.querySelector<HTMLInputElement>("#fw-prompt-input")!;
    const error = overlay.querySelector<HTMLElement>("#fw-prompt-error")!;
    input.value = value;
    if (placeholder) input.placeholder = placeholder;

    let done = false;
    const finish = (result: string | null): void => {
      if (done) return;
      done = true;
      document.removeEventListener("keydown", onKey, true);
      overlay.remove();
      resolve(result);
    };

    const accept = (): void => {
      const v = input.value.trim();
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
      // Capture phase: the editor's own keymaps must not see the keys that
      // drive the dialog.
      if (e.key === "Escape") {
        e.preventDefault();
        e.stopPropagation();
        finish(null);
      } else if (e.key === "Enter") {
        e.preventDefault();
        e.stopPropagation();
        accept();
      }
    };
    document.addEventListener("keydown", onKey, true);

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
    document.getElementById("fw-alert")?.remove();

    const overlay = document.createElement("div");
    overlay.id = "fw-alert";
    overlay.className = "fw-modal-overlay";
    overlay.innerHTML = `
      <div class="fw-modal" role="alertdialog" aria-modal="true" aria-label="${title}">
        <h2>${title}</h2>
        <p class="fw-alert-message">${message}</p>
        <div class="actions">
          <button id="fw-alert-ok" type="button" class="primary">OK</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const ok = overlay.querySelector<HTMLButtonElement>("#fw-alert-ok")!;
    const finish = (): void => {
      document.removeEventListener("keydown", onKey, true);
      overlay.remove();
      resolve();
    };
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === "Escape" || e.key === "Enter") {
        e.preventDefault();
        e.stopPropagation();
        finish();
      }
    };
    document.addEventListener("keydown", onKey, true);
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) finish();
    });
    ok.addEventListener("click", finish);
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
    document.getElementById("fw-confirm")?.remove();

    const overlay = document.createElement("div");
    overlay.id = "fw-confirm";
    overlay.className = "fw-modal-overlay";
    overlay.innerHTML = `
      <div class="fw-modal" role="alertdialog" aria-modal="true" aria-label="${title}">
        <h2>${title}</h2>
        <p class="fw-alert-message">${message}</p>
        <div class="actions">
          <button id="fw-confirm-cancel" type="button">Cancel</button>
          <button id="fw-confirm-ok" type="button" class="primary${danger ? " danger" : ""}">${okLabel}</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    let done = false;
    const finish = (value: boolean): void => {
      if (done) return;
      done = true;
      document.removeEventListener("keydown", onKey, true);
      overlay.remove();
      resolve(value);
    };
    const onKey = (e: KeyboardEvent): void => {
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
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) finish(false);
    });
    overlay.querySelector("#fw-confirm-cancel")!.addEventListener("click", () => finish(false));
    const ok = overlay.querySelector<HTMLButtonElement>("#fw-confirm-ok")!;
    ok.addEventListener("click", () => finish(true));
    ok.focus();
  });
}
