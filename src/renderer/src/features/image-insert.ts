import type { Editor } from "@tiptap/core";

// `window.fw` is declared globally in app.ts (same compilation unit).

const MAX_WIDTH_PX = 900;

/**
 * Open the native image picker, read the chosen file as a data URL (via the
 * main process), measure its natural size, and insert an `image` node at the
 * cursor. Widths larger than MAX_WIDTH_PX are scaled down proportionally so
 * the embedded document JSON stays lean.
 */
export async function insertImageAtCursor(editor: Editor): Promise<void> {
  const fw = window.fw;
  if (!fw) return;
  const picked = await fw.pickImage();
  if (!picked) return;
  await new Promise<void>((resolve) => {
    const img = new Image();
    img.onload = () => {
      let w = img.naturalWidth > 0 ? img.naturalWidth : MAX_WIDTH_PX;
      let h = img.naturalHeight > 0 ? img.naturalHeight : Math.round((MAX_WIDTH_PX * 3) / 4);
      if (w > MAX_WIDTH_PX) {
        h = Math.max(1, Math.round((h * MAX_WIDTH_PX) / w));
        w = MAX_WIDTH_PX;
      }
      editor.chain().focus().setImage({ src: picked.dataUrl, alt: picked.name, width: w, height: h }).run();
      resolve();
    };
    img.onerror = () => {
      editor.chain().focus().setImage({ src: picked.dataUrl, alt: picked.name }).run();
      resolve();
    };
    img.src = picked.dataUrl;
  });
}
