/**
 * URL policy shared by the renderer (what may be stored on a link) and the
 * main process (what may be handed to the OS shell). Both sides check, because
 * a compromised renderer is the threat model here, not the user.
 */

/** Schemes FreeWord will store on a link and will open in the desktop browser. */
const ALLOWED_SCHEMES = ["http:", "https:", "mailto:", "tel:"];

/** Absolute URL with an allowed scheme, or null when it has none / is unsafe. */
export function safeExternalUrl(url: unknown): string | null {
  if (typeof url !== "string") return null;
  const trimmed = url.trim();
  if (!trimmed) return null;
  let parsed: URL;
  try {
    parsed = new URL(trimmed);
  } catch {
    return null;
  }
  // `#anchor` links are relative and never leave the document; the editor uses
  // them for internal references, and the shell must never see them.
  if (parsed.protocol === "file:" || parsed.protocol === "javascript:") return null;
  return ALLOWED_SCHEMES.includes(parsed.protocol) ? trimmed : null;
}

/** True when `url` may be stored on a link or opened externally. */
export function isSafeExternalUrl(url: unknown): boolean {
  return safeExternalUrl(url) !== null;
}
