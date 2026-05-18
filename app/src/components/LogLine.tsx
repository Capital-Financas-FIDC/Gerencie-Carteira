import type { LogEvent } from "../types/log";

interface LogLineProps {
  event: LogEvent;
}

const ICON: Record<string, string> = {
  info: "i",
  success: "OK",
  warning: "!",
  error: "X",
  step: ">",
};

function formatTime(iso: string): string {
  try {
    const d = new Date(iso);
    return d.toLocaleTimeString("pt-BR", { hour12: false });
  } catch {
    return iso;
  }
}

export function LogLine({ event }: LogLineProps) {
  const icon = ICON[event.level] ?? "?";
  const isBootstrap = event.step === "workspace.bootstrap";
  const isNeedsUser = event.step === "excel.base.needs_user";
  const created = Array.isArray((event.data as any)?.created) ? ((event.data as any).created as string[]) : null;
  const searched = (event.data as any)?.searched as Record<string, string | null> | undefined;
  const showBootstrapCard = isBootstrap && created && created.length > 0;
  const showSearchedCard = isNeedsUser && searched;

  return (
    <div className={`log-line log-line-${event.level}`} role="listitem">
      <span className="log-line-time">{formatTime(event.ts)}</span>
      <span className="log-line-icon" aria-hidden="true">{icon}</span>
      <div className="log-line-body">
        <span className="log-line-msg">{event.msg}</span>
        {event.step && <span className="log-line-step">{event.step}</span>}
        {typeof event.progress === "number" && (
          <span className="log-line-progress">{event.progress}%</span>
        )}
        {showBootstrapCard && (
          <div className="bootstrap-card">
            <strong>Pastas criadas</strong>
            <ul>
              {created!.map((p) => <li key={p}><code>{p}</code></li>)}
            </ul>
          </div>
        )}
        {showSearchedCard && (
          <div className="bootstrap-card bootstrap-card-warn">
            <strong>Locais pesquisados</strong>
            <ul>
              {Object.entries(searched!).map(([k, v]) => (
                <li key={k}><em>{k}:</em> <code>{v ?? "nao encontrado"}</code></li>
              ))}
            </ul>
          </div>
        )}
      </div>
    </div>
  );
}
