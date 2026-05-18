import { useEffect, useRef, useState } from "react";
import type { LogEvent } from "../types/log";
import { LogLine } from "./LogLine";

interface LogViewerProps {
  logs: LogEvent[];
}

export function LogViewer({ logs }: LogViewerProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [autoScroll, setAutoScroll] = useState(true);

  useEffect(() => {
    if (!autoScroll || !containerRef.current) return;
    const el = containerRef.current;
    el.scrollTop = el.scrollHeight;
  }, [logs, autoScroll]);

  const handleScroll = () => {
    const el = containerRef.current;
    if (!el) return;
    const atBottom = el.scrollHeight - el.scrollTop - el.clientHeight < 20;
    setAutoScroll(atBottom);
  };

  return (
    <section className="log-viewer" aria-label="Log de execucao">
      <div className="log-viewer-header">
        <h2>Log</h2>
        <span className="log-viewer-count">{logs.length} evento(s)</span>
      </div>
      <div
        className="log-viewer-body"
        ref={containerRef}
        onScroll={handleScroll}
      >
        {logs.length === 0 ? (
          <p className="log-viewer-empty">Aguardando execucao. O log aparece aqui em tempo real.</p>
        ) : (
          logs.map((ev, i) => <LogLine key={i} event={ev} />)
        )}
      </div>
    </section>
  );
}
