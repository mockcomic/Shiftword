import type { PageSettings, PageSize } from "../../../shared/types";

const clampMm = (v: number): number => Math.min(100, Math.max(5, Number.isFinite(v) ? v : 25.4));

/**
 * Show the Page Setup modal (size + 4 margins in mm) and invoke `onApply`
 * with the chosen settings. Cancel closes without calling the callback.
 */
export function openPageSetup(current: PageSettings, onApply: (settings: PageSettings) => void): void {
  document.getElementById("fw-page-setup")?.remove();

  const overlay = document.createElement("div");
  overlay.id = "fw-page-setup";
  overlay.className = "fw-modal-overlay";
  overlay.innerHTML = `
    <div class="fw-modal">
      <h2>Page Setup</h2>
      <div class="row">
        <label for="fw-ps-size">Page size</label>
        <select id="fw-ps-size">
          <option value="a4">A4 (210 × 297 mm)</option>
          <option value="letter">Letter (215.9 × 279.4 mm)</option>
        </select>
      </div>
      <div class="row">
        <label for="fw-ps-mt">Top margin (mm)</label>
        <input id="fw-ps-mt" type="number" min="5" max="100" step="0.1">
      </div>
      <div class="row">
        <label for="fw-ps-mr">Right margin (mm)</label>
        <input id="fw-ps-mr" type="number" min="5" max="100" step="0.1">
      </div>
      <div class="row">
        <label for="fw-ps-mb">Bottom margin (mm)</label>
        <input id="fw-ps-mb" type="number" min="5" max="100" step="0.1">
      </div>
      <div class="row">
        <label for="fw-ps-ml">Left margin (mm)</label>
        <input id="fw-ps-ml" type="number" min="5" max="100" step="0.1">
      </div>
      <div class="actions">
        <button id="fw-ps-cancel" type="button">Cancel</button>
        <button id="fw-ps-ok" type="button" class="primary">OK</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  const close = (): void => {
    overlay.remove();
    document.removeEventListener("keydown", onKey, true);
  };
  const onKey = (e: KeyboardEvent): void => {
    if (e.isComposing || e.keyCode === 229) return;
    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      close();
    } else if (e.key === "Enter") {
      // Enter confirms from any field, matching the modal.ts dialogs.
      e.preventDefault();
      e.stopPropagation();
      okBtn.click();
    }
  };
  document.addEventListener("keydown", onKey, true);

  const sel = overlay.querySelector<HTMLSelectElement>("#fw-ps-size");
  const inputs = {
    mt: overlay.querySelector<HTMLInputElement>("#fw-ps-mt"),
    mr: overlay.querySelector<HTMLInputElement>("#fw-ps-mr"),
    mb: overlay.querySelector<HTMLInputElement>("#fw-ps-mb"),
    ml: overlay.querySelector<HTMLInputElement>("#fw-ps-ml"),
  };

  if (sel) sel.value = current.size;
  if (inputs.mt) inputs.mt.value = String(current.mt);
  if (inputs.mr) inputs.mr.value = String(current.mr);
  if (inputs.mb) inputs.mb.value = String(current.mb);
  if (inputs.ml) inputs.ml.value = String(current.ml);

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) close();
  });
  const okBtn = overlay.querySelector("#fw-ps-ok") as HTMLButtonElement;
  (overlay.querySelector("#fw-ps-cancel") as HTMLButtonElement).addEventListener("click", close);
  okBtn.addEventListener("click", () => {
    const settings: PageSettings = {
      size: (sel?.value ?? "a4") as PageSize,
      mt: clampMm(parseFloat(inputs.mt?.value ?? "")),
      mr: clampMm(parseFloat(inputs.mr?.value ?? "")),
      mb: clampMm(parseFloat(inputs.mb?.value ?? "")),
      ml: clampMm(parseFloat(inputs.ml?.value ?? "")),
    };
    close();
    onApply(settings);
  });
}
