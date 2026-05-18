import { contextBridge, ipcRenderer, type IpcRendererEvent } from "electron";

export interface LogEvent {
  level: "info" | "success" | "warning" | "error" | "step";
  ts: string;
  msg: string;
  step?: string;
  progress?: number;
  data?: Record<string, unknown>;
}

export interface ScriptDone {
  exitCode: number;
  spreadsheetPath: string | null;
  durationMs: number;
}

export type IpcOk<T = unknown> = { ok: true } & T;
export type IpcErr = { ok: false; reason: string };
export type IpcResult<T = unknown> = IpcOk<T> | IpcErr;

const api = {
  getAppVersion: (): Promise<string> => ipcRenderer.invoke("app:version"),
  runScript: (): Promise<IpcResult> => ipcRenderer.invoke("script:run"),
  cancelScript: (): Promise<IpcResult> => ipcRenderer.invoke("script:cancel"),
  openFile: (path: string): Promise<IpcResult> => ipcRenderer.invoke("file:open", path),
  chooseBaseSpreadsheet: (): Promise<IpcResult<{ destino?: string; origem?: string; renamed?: boolean }>> =>
    ipcRenderer.invoke("dialog:choose-base"),

  onLog: (listener: (ev: LogEvent) => void): (() => void) => {
    const handler = (_e: IpcRendererEvent, payload: LogEvent) => listener(payload);
    ipcRenderer.on("script:log", handler);
    return () => ipcRenderer.off("script:log", handler);
  },
  onDone: (listener: (result: ScriptDone) => void): (() => void) => {
    const handler = (_e: IpcRendererEvent, payload: ScriptDone) => listener(payload);
    ipcRenderer.on("script:done", handler);
    return () => ipcRenderer.off("script:done", handler);
  },
};

contextBridge.exposeInMainWorld("electronAPI", api);

export type ElectronAPI = typeof api;
