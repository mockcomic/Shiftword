import { contextBridge, ipcRenderer } from "electron";
import type {
  AppSettings,
  FwApi,
  MenuCommand,
  OpenFileResult,
  PageMeta,
  PickImageResult,
  PrintResult,
  RecentEntry,
  SaveDocPayload,
  SaveRawPayload,
  WorkspacePage,
  WriteFileResult,
} from "../shared/types";

const api: FwApi = {
  openFile: (path?: string) => ipcRenderer.invoke("file:open", path) as Promise<OpenFileResult | null>,
  saveDoc: (payload: SaveDocPayload) =>
    ipcRenderer.invoke("docx:save", payload) as Promise<WriteFileResult | null>,
  saveRaw: (payload: SaveRawPayload) =>
    ipcRenderer.invoke("file:saveRaw", payload) as Promise<WriteFileResult | null>,
  pickImage: () => ipcRenderer.invoke("file:pickImage") as Promise<PickImageResult | null>,
  exportPdf: (defaultName: string) =>
    ipcRenderer.invoke("pdf:export", defaultName) as Promise<WriteFileResult | null>,
  print: () => ipcRenderer.invoke("print") as Promise<PrintResult | null>,
  getRecent: () => ipcRenderer.invoke("recent:get") as Promise<RecentEntry[]>,
  getSettings: () => ipcRenderer.invoke("settings:get") as Promise<AppSettings>,
  setSettings: (patch: Partial<AppSettings>) =>
    ipcRenderer.invoke("settings:set", patch) as Promise<AppSettings>,
  listPages: () => ipcRenderer.invoke("page:list") as Promise<PageMeta[]>,
  readPage: (id: string) => ipcRenderer.invoke("page:read", id) as Promise<WorkspacePage | null>,
  writePage: (id: string, json: Record<string, unknown>) =>
    ipcRenderer.invoke("page:write", id, json) as Promise<{ error?: string }>,
  createPage: (opts?: { title?: string; parentId?: string | null }) =>
    ipcRenderer.invoke("page:create", opts) as Promise<PageMeta | null>,
  renamePage: (id: string, patch: { title?: string; icon?: string }) =>
    ipcRenderer.invoke("page:rename", id, patch) as Promise<PageMeta | null>,
  movePage: (id: string, parentId: string | null, order: number) =>
    ipcRenderer.invoke("page:move", id, parentId, order) as Promise<{ error?: string }>,
  deletePage: (id: string) => ipcRenderer.invoke("page:delete", id) as Promise<{ error?: string }>,
  openExternal: (url: string) => ipcRenderer.invoke("shell:openExternal", url) as Promise<boolean>,
  onCommand: (cb: (cmd: MenuCommand) => void) => {
    const listener = (_e: Electron.IpcRendererEvent, cmd: MenuCommand): void => cb(cmd);
    ipcRenderer.on("menu:command", listener);
    return () => {
      ipcRenderer.removeListener("menu:command", listener);
    };
  },
  sendQuitResponse: (payload: { dirty: boolean; name: string; path: string | null }) => {
    ipcRenderer.send("app:quit-response", payload);
  },
  onWillQuit: (cb: () => void) => {
    const listener = (): void => cb();
    ipcRenderer.on("app:will-quit", listener);
    return () => {
      ipcRenderer.removeListener("app:will-quit", listener);
    };
  },
  onSaveThenQuit: (cb: () => void) => {
    const listener = (): void => cb();
    ipcRenderer.on("app:save-then-quit", listener);
    return () => {
      ipcRenderer.removeListener("app:save-then-quit", listener);
    };
  },
};

contextBridge.exposeInMainWorld("fw", api);

declare global {
  interface Window {
    fw: FwApi;
  }
}
