import { app, dialog, Menu, shell } from "electron";
import type { BrowserWindow, MessageBoxOptions, MenuItemConstructorOptions } from "electron";
import { readRecent } from "./recent";
import { readSettings } from "./settings";
import { safeExternalUrl } from "../shared/urls";

type GetWindow = () => BrowserWindow | null;

let getWindow: GetWindow = () => null;
let onQuitRequest: () => void = () => {};

export function configureMenu(getWin: GetWindow, onQuit: () => void): void {
  getWindow = getWin;
  onQuitRequest = onQuit;
}

function send(id: string, filePath?: string): void {
  const win = getWindow();
  if (win && !win.isDestroyed()) {
    win.webContents.send("menu:command", filePath ? { id, path: filePath } : { id });
  }
}

function showAbout(): void {
  const options: MessageBoxOptions = {
    type: "info",
    title: "About FreeWord",
    message: `FreeWord ${app.getVersion()}`,
    detail:
      "An offline Word-like document editor.\n\n" +
      "Known limitations:\n" +
      "• WebP images are not supported (use PNG, JPEG, GIF, or BMP).\n" +
      "• Images cannot be resized after inserting them.\n" +
      "• A block taller than one page may be clipped in PDF export.\n" +
      "• Forced page breaks in imported .docx files may be re-paginated.\n" +
      "• No find & replace, headers/footers, footnotes or track changes yet.",
    buttons: ["OK"],
  };
  const win = getWindow();
  if (win) dialog.showMessageBox(win, options);
  else dialog.showMessageBox(options);
}

/**
 * Right-click menu for the editor.
 *
 * Electron does not install one: with no `context-menu` handler the renderer
 * shows nothing at all, which cost two things — mouse Cut/Copy/Paste, and any
 * visible result of `spellcheck: true`, since the Chromium suggestions only
 * arrive as `params.dictionarySuggestions` on this event.
 */
export function buildContextMenu(win: BrowserWindow, params: Electron.ContextMenuParams): Menu {
  const items: MenuItemConstructorOptions[] = [];
  const wc = win.webContents;

  if (params.isEditable) {
    items.push({ role: "undo" }, { role: "redo" }, { type: "separator" });
    items.push(
      { role: "cut", enabled: params.editFlags.canCut },
      { role: "copy", enabled: params.editFlags.canCopy },
      { role: "paste", enabled: params.editFlags.canPaste },
      { type: "separator" },
      { role: "selectAll" },
    );
  } else if (params.selectionText) {
    items.push({ role: "copy" }, { type: "separator" });
  }

  if (params.isEditable && params.misspelledWord) {
    // The misspelled word is shown as a disabled header so the list of
    // alternatives has an obvious subject.
    items.push({ type: "separator" }, { label: `Spelling: “${params.misspelledWord}”`, enabled: false });
    const suggestions = params.dictionarySuggestions.slice(0, 6);
    if (suggestions.length === 0) {
      items.push({ label: "No suggestions", enabled: false });
    }
    for (const word of suggestions) {
      items.push({
        label: word,
        click: () => {
          void wc.replaceMisspelling(word);
        },
      });
    }
    items.push({
      label: "Add to dictionary",
      click: () => {
        void wc.session.addWordToSpellCheckerDictionary(params.misspelledWord);
      },
    });
  }

  if (params.linkURL) {
    items.push({ type: "separator" });
    if (safeExternalUrl(params.linkURL)) {
      items.push({
        label: "Open Link in Browser",
        click: () => {
          void shell.openExternal(safeExternalUrl(params.linkURL) as string).catch(() => undefined);
        },
      });
    }
    items.push({
      label: "Copy Link Address",
      enabled: false, // no clipboard write without the selection API handshake
    });
  }

  return Menu.buildFromTemplate(items);
}

export async function rebuildMenu(): Promise<void> {
  const recent = await readRecent();
  const settings = await readSettings();
  const recentItems: MenuItemConstructorOptions[] =
    recent.length === 0
      ? [{ label: "(none)", enabled: false }]
      : recent.slice(0, 8).map((entry) => ({
          label: entry.name,
          click: () => send("recent:open", entry.path),
        }));

  const template: MenuItemConstructorOptions[] = [
    {
      label: "File",
      submenu: [
        { label: "New", accelerator: "CmdOrCtrl+N", click: () => send("file:new") },
        {
          label: "New Page",
          accelerator: "CmdOrCtrl+Alt+N",
          click: () => send("page:new"),
        },
        { label: "Open…", accelerator: "CmdOrCtrl+O", click: () => send("file:open") },
        { label: "Save", accelerator: "CmdOrCtrl+S", click: () => send("file:save") },
        { label: "Save As…", accelerator: "Shift+CmdOrCtrl+S", click: () => send("file:saveAs") },
        { type: "separator" },
        { label: "Export PDF…", accelerator: "CmdOrCtrl+F12", click: () => send("file:exportPdf") },
        { label: "Save As Plain Text…", click: () => send("file:exportTxt") },
        { label: "Save As HTML…", click: () => send("file:exportHtml") },
        { label: "Print…", accelerator: "CmdOrCtrl+P", click: () => send("file:print") },
        { type: "separator" },
        { label: "Recent", submenu: recentItems },
        { type: "separator" },
        { label: "Quit", accelerator: "CmdOrCtrl+Q", click: () => onQuitRequest() },
      ],
    },
    {
      label: "Edit",
      submenu: [
        { role: "undo" },
        { role: "redo" },
        { type: "separator" },
        { role: "cut" },
        { role: "copy" },
        { role: "paste" },
        { role: "selectAll" },
      ],
    },
    {
      label: "View",
      submenu: [
        { label: "Quick Switcher…", accelerator: "CmdOrCtrl+K", click: () => send("view:quickSwitcher") },
        {
          label: "Sidebar",
          type: "checkbox",
          checked: settings.sidebarOpen,
          accelerator: "CmdOrCtrl+Shift+B",
          click: () => send("view:toggleSidebar"),
        },
        { type: "separator" },
        { label: "Zoom In", accelerator: "CmdOrCtrl+=", click: () => send("view:zoomIn") },
        { label: "Zoom Out", accelerator: "CmdOrCtrl+-", click: () => send("view:zoomOut") },
        { label: "Reset Zoom", accelerator: "CmdOrCtrl+0", click: () => send("view:zoomReset") },
        { type: "separator" },
        {
          label: "Notion Mode",
          type: "checkbox",
          checked: settings.notionMode,
          accelerator: "CmdOrCtrl+Shift+Space",
          click: () => send("view:notionMode"),
        },
        {
          label: "Show Toolbar",
          type: "checkbox",
          checked: settings.toolbarPinned,
          enabled: settings.notionMode,
          click: () => send("view:toggleToolbarPin"),
        },
        {
          label: "Theme",
          submenu: (["light", "dark", "system"] as const).map((t) => ({
            label: t === "light" ? "Light" : t === "dark" ? "Dark" : "Follow system",
            type: "radio" as const,
            checked: settings.theme === t,
            click: () => send(`view:theme:${t}`),
          })),
        },
        { type: "separator" },
        { label: "Page Setup…", click: () => send("view:pageSetup") },
      ],
    },
    {
      label: "Insert",
      submenu: [
        { label: "Table", accelerator: "CmdOrCtrl+Alt+T", click: () => send("insert:table") },
        { label: "Image…", click: () => send("insert:image") },
        { type: "separator" },
        { label: "Page Break", accelerator: "CmdOrCtrl+Enter", click: () => send("insert:pageBreak") },
      ],
    },
    {
      label: "Help",
      submenu: [{ label: "About FreeWord", click: showAbout }],
    },
  ];

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}
