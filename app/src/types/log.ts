export type LogLevel = "info" | "success" | "warning" | "error" | "step";

export interface LogEvent {
  level: LogLevel;
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

export type RunStatus = "idle" | "running" | "success" | "warning" | "error";
