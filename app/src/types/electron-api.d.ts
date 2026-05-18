import type { LogEvent, ScriptDone } from "./log";

type IpcResult<T = unknown> = ({ ok: true } & T) | { ok: false; reason: string };

export interface ElectronAPI {
  getAppVersion: () => Promise<string>;
  runScript: () => Promise<IpcResult>;
  cancelScript: () => Promise<IpcResult>;
  openFile: (path: string) => Promise<IpcResult>;
  chooseBaseSpreadsheet: () => Promise<
    IpcResult<{ destino?: string; origem?: string; renamed?: boolean }>
  >;
  onLog: (listener: (ev: LogEvent) => void) => () => void;
  onDone: (listener: (result: ScriptDone) => void) => () => void;
}

declare global {
  interface Window {
    electronAPI: ElectronAPI;
  }
}

export {};
