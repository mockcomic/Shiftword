import type { PageMeta } from "../../../shared/types";

/**
 * The Notion-style page header (the "hero") for a workspace page: a cover
 * banner, a big emoji icon, and an inline-editable title sitting above the
 * document on the continuous canvas.
 *
 * Everything it shows is page *metadata* (PageMeta title/icon/cover), never
 * document content — the same posture as the sidebar rows, which is what
 * keeps the .docx export byte-identical whether or not a page has a cover.
 * Visibility is CSS (`body:not(.notion-mode)`), and print/PDF hides it like
 * the rest of the chrome.
 */

export interface PageHeroHandlers {
  /** Ask the user to pick a cover image for the page (native picker). */
  onPickCover(id: string): void;
  onRemoveCover(id: string): void;
  onSetIcon(id: string): void;
  /** Commit a new title typed into the hero. */
  onRename(id: string, title: string): void;
}

export interface PageHero {
  /** Show the header for `meta`, or hide the whole element for null. */
  setMeta(meta: PageMeta | null): void;
  /** Re-render the title field if `id` is currently shown (sidebar renames). */
  refreshTitle(title: string): void;
  element: HTMLElement;
}

export function createPageHero(handlers: PageHeroHandlers): PageHero {
  const root = document.createElement("div");
  root.id = "fw-page-hero";
  root.hidden = true;

  // Mount above the editor inside the canvas (notion-mode centres both).
  const canvas = document.querySelector<HTMLElement>("#canvas");
  const editorEl = document.querySelector<HTMLElement>("#editor");
  if (canvas) canvas.insertBefore(root, editorEl ?? canvas.firstChild);

  // --- cover ---------------------------------------------------------------
  const coverBox = document.createElement("div");
  coverBox.className = "hero-cover";
  coverBox.hidden = true;
  const coverImg = document.createElement("img");
  coverImg.className = "hero-cover-img";
  coverImg.alt = "";
  const coverActions = document.createElement("div");
  coverActions.className = "hero-cover-actions";
  const changeCoverBtn = document.createElement("button");
  changeCoverBtn.type = "button";
  changeCoverBtn.textContent = "Change cover";
  const removeCoverBtn = document.createElement("button");
  removeCoverBtn.type = "button";
  removeCoverBtn.textContent = "Remove cover";
  coverActions.append(changeCoverBtn, removeCoverBtn);
  coverBox.append(coverImg, coverActions);

  // --- body: icon + title + add-cover affordance ---------------------------
  const body = document.createElement("div");
  body.className = "hero-body";

  const iconBtn = document.createElement("button");
  iconBtn.type = "button";
  iconBtn.className = "hero-icon";
  iconBtn.title = "Set page icon";
  iconBtn.textContent = "📄";

  const titleInput = document.createElement("input");
  titleInput.type = "text";
  titleInput.className = "hero-title";
  titleInput.maxLength = 200;
  titleInput.placeholder = "Untitled";
  titleInput.spellcheck = false;

  const addCoverBtn = document.createElement("button");
  addCoverBtn.type = "button";
  addCoverBtn.className = "hero-add-cover";
  addCoverBtn.textContent = "🖼 Add cover";

  const metaRow = document.createElement("div");
  metaRow.className = "hero-actions";
  metaRow.append(iconBtn, addCoverBtn);
  body.append(metaRow, titleInput);
  root.append(coverBox, body);

  // The hero never steals focus from the editor by accident.
  root.addEventListener("mousedown", (e) => {
    if ((e.target as HTMLElement).closest("button")) e.preventDefault();
  });

  let currentId: string | null = null;

  iconBtn.addEventListener("click", () => {
    if (currentId) handlers.onSetIcon(currentId);
  });
  addCoverBtn.addEventListener("click", () => {
    if (currentId) handlers.onPickCover(currentId);
  });
  changeCoverBtn.addEventListener("click", () => {
    if (currentId) handlers.onPickCover(currentId);
  });
  removeCoverBtn.addEventListener("click", () => {
    if (currentId) handlers.onRemoveCover(currentId);
  });

  // Commit the typed title on blur/Enter (Notion's own behaviour); Escape
  // restores the stored value instead of committing.
  const commitTitle = (): void => {
    if (currentId === null) return;
    const value = titleInput.value.trim();
    if (!value) {
      // An empty title is not a rename — put the old one back.
      titleInput.value = currentTitle;
      return;
    }
    if (value !== currentTitle) handlers.onRename(currentId, value);
  };
  let currentTitle = "Untitled";
  titleInput.addEventListener("blur", commitTitle);
  titleInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      titleInput.blur();
    } else if (e.key === "Escape") {
      titleInput.value = currentTitle;
      titleInput.blur();
    }
  });

  return {
    element: root,
    setMeta(meta) {
      if (!meta) {
        currentId = null;
        root.hidden = true;
        return;
      }
      currentId = meta.id;
      root.hidden = false;
      const cover = typeof meta.cover === "string" && meta.cover ? meta.cover : null;
      coverBox.hidden = !cover;
      addCoverBtn.style.visibility = cover ? "hidden" : "";
      if (cover) coverImg.src = cover;
      else coverImg.removeAttribute("src");
      iconBtn.textContent = meta.icon || "📄";
      iconBtn.style.visibility = "visible";
      // Never overwrite what the user is mid-way through typing.
      if (document.activeElement !== titleInput) {
        currentTitle = meta.title || "Untitled";
        titleInput.value = meta.title || "";
      }
    },
    refreshTitle(title) {
      if (currentId === null) return;
      currentTitle = title || "Untitled";
      if (document.activeElement !== titleInput) titleInput.value = title || "";
    },
  };
}
