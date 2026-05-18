import type { RunStatus } from "../types/log";

interface StatusBadgeProps {
  status: RunStatus;
  step?: string;
  durationMs?: number;
}

const LABEL: Record<RunStatus, string> = {
  idle: "Ocioso",
  running: "Executando",
  success: "Sucesso",
  warning: "Concluido com avisos",
  error: "Erro",
};

export function StatusBadge({ status, step, durationMs }: StatusBadgeProps) {
  const detail = status === "running" && step
    ? step
    : durationMs
      ? `${(durationMs / 1000).toFixed(1)}s`
      : null;

  return (
    <div className="status-badge" data-status={status} role="status" aria-live="polite">
      <span className="status-dot" />
      <div className="status-text">
        <strong>{LABEL[status]}</strong>
        {detail && <small>{detail}</small>}
      </div>
    </div>
  );
}
