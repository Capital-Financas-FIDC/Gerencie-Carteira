import { useCallback, useEffect, useReducer, useRef } from "react";
import type { LogEvent, RunStatus, ScriptDone } from "../types/log";

const EXIT_BASE_NEEDS_USER = 4;

interface State {
  status: RunStatus;
  logs: LogEvent[];
  currentStep?: string;
  spreadsheetPath: string | null;
  durationMs?: number;
}

type Action =
  | { type: "run" }
  | { type: "log"; ev: LogEvent }
  | { type: "done"; result: ScriptDone }
  | { type: "system"; ev: LogEvent }
  | { type: "clear" };

const initialState: State = {
  status: "idle",
  logs: [],
  spreadsheetPath: null,
};

function reducer(state: State, action: Action): State {
  switch (action.type) {
    case "run":
      return { ...initialState, status: "running", logs: [] };
    case "log":
    case "system": {
      const ev = action.ev;
      const resultData = (ev.data as any)?.result as
        | { status?: string; spreadsheet_path?: string | null }
        | undefined;
      return {
        ...state,
        logs: [...state.logs, ev],
        currentStep: ev.step ?? state.currentStep,
        spreadsheetPath: resultData?.spreadsheet_path ?? state.spreadsheetPath,
      };
    }
    case "done": {
      const lastResult = [...state.logs].reverse().find((l) => (l.data as any)?.result);
      const resultStatus = (lastResult?.data as any)?.result?.status as string | undefined;
      const final: RunStatus =
        action.result.exitCode !== 0
          ? "error"
          : resultStatus === "warning"
            ? "warning"
            : resultStatus === "error"
              ? "error"
              : "success";
      return {
        ...state,
        status: final,
        spreadsheetPath: action.result.spreadsheetPath ?? state.spreadsheetPath,
        durationMs: action.result.durationMs,
        currentStep: undefined,
      };
    }
    case "clear":
      return { ...initialState };
    default:
      return state;
  }
}

function systemEvent(msg: string, level: LogEvent["level"] = "info", step?: string): LogEvent {
  return { level, ts: new Date().toISOString(), msg, step };
}

export function useScriptRunner() {
  const [state, dispatch] = useReducer(reducer, initialState);
  const lastStepRef = useRef<string | undefined>(undefined);

  useEffect(() => {
    const offLog = window.electronAPI.onLog((ev) => {
      if (ev.step) lastStepRef.current = ev.step;
      dispatch({ type: "log", ev });
    });
    const offDone = window.electronAPI.onDone(async (result) => {
      dispatch({ type: "done", result });

      // Auto-prompt quando cascata Python falhou
      if (
        result.exitCode === EXIT_BASE_NEEDS_USER ||
        lastStepRef.current === "excel.base.needs_user"
      ) {
        dispatch({
          type: "system",
          ev: systemEvent(
            "Abrindo seletor para escolher a planilha base...",
            "info",
            "dialog.choose-base",
          ),
        });

        const choice = await window.electronAPI.chooseBaseSpreadsheet();

        if (!choice.ok) {
          const reason = (choice as { reason: string }).reason;
          dispatch({
            type: "system",
            ev: systemEvent(
              reason === "canceled"
                ? "Selecao de planilha cancelada. Clique em Executar para tentar novamente."
                : `Falha ao copiar planilha base: ${reason}`,
              reason === "canceled" ? "warning" : "error",
              "dialog.cancel",
            ),
          });
          return;
        }

        const data = choice as { ok: true; destino?: string; origem?: string; renamed?: boolean };
        dispatch({
          type: "system",
          ev: systemEvent(
            data.renamed
              ? `Planilha copiada e renomeada como base: ${data.destino}`
              : `Planilha copiada como base: ${data.destino}`,
            "success",
            "dialog.copied",
          ),
        });

        // Auto-rerun
        dispatch({
          type: "system",
          ev: systemEvent("Reiniciando automacao com a nova planilha base...", "info", "auto-rerun"),
        });

        // Pequeno delay para o usuario ver a mensagem antes do clear
        setTimeout(() => {
          dispatch({ type: "run" });
          window.electronAPI.runScript();
        }, 600);
      }
    });
    return () => {
      offLog();
      offDone();
    };
  }, []);

  const run = useCallback(async () => {
    lastStepRef.current = undefined;
    dispatch({ type: "run" });
    const result = await window.electronAPI.runScript();
    if (!result.ok) {
      dispatch({
        type: "log",
        ev: {
          level: "error",
          ts: new Date().toISOString(),
          msg: `Nao foi possivel iniciar: ${result.reason}`,
          step: "runner.error",
        },
      });
      dispatch({ type: "done", result: { exitCode: -1, spreadsheetPath: null, durationMs: 0 } });
    }
  }, []);

  const cancel = useCallback(async () => {
    await window.electronAPI.cancelScript();
  }, []);

  const clear = useCallback(() => dispatch({ type: "clear" }), []);

  const openSpreadsheet = useCallback(async () => {
    if (!state.spreadsheetPath) return;
    await window.electronAPI.openFile(state.spreadsheetPath);
  }, [state.spreadsheetPath]);

  return {
    status: state.status,
    logs: state.logs,
    currentStep: state.currentStep,
    spreadsheetPath: state.spreadsheetPath,
    durationMs: state.durationMs,
    run,
    cancel,
    clear,
    openSpreadsheet,
  };
}
