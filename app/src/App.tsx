import { useEffect, useState } from "react";
import { useScriptRunner } from "./hooks/useScriptRunner";
import { RunButton } from "./components/RunButton";
import { StatusBadge } from "./components/StatusBadge";
import { LogViewer } from "./components/LogViewer";
import { ActionBar } from "./components/ActionBar";
import { GerentesOrfaosForm } from "./components/GerentesOrfaosForm";

export default function App() {
  const {
    status,
    logs,
    currentStep,
    spreadsheetPath,
    durationMs,
    pendingInput,
    run,
    clear,
    openSpreadsheet,
    submitGerentes,
  } = useScriptRunner();

  // Versao exibida vem da fonte unica (app/package.json via IPC), nunca hardcoded.
  const [version, setVersion] = useState("");
  useEffect(() => {
    window.electronAPI.getAppVersion().then(setVersion).catch(() => setVersion(""));
  }, []);

  const isRunning = status === "running";

  return (
    <div className="app-shell">
      <header className="app-header">
        <div className="app-title">
          <h1>Gerencie Carteira</h1>
          {version && <span className="app-version">v{version}</span>}
        </div>
        <StatusBadge status={status} step={currentStep} durationMs={durationMs} />
      </header>

      <main className="app-main">
        <div className="run-panel">
          <RunButton running={isRunning} onClick={run} disabled={isRunning} />
          <p className="run-hint">
            {isRunning
              ? "Buscando emails do Serasa e atualizando planilha..."
              : "Clique em Executar para iniciar a automacao."}
          </p>
        </div>
        <LogViewer logs={logs} />
      </main>

      <ActionBar
        spreadsheetPath={spreadsheetPath}
        onOpen={openSpreadsheet}
        onClear={clear}
        disabled={isRunning}
      />

      {pendingInput && (
        <GerentesOrfaosForm orfaos={pendingInput.orfaos} onSubmit={submitGerentes} />
      )}
    </div>
  );
}
