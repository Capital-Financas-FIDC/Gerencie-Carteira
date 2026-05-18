interface ActionBarProps {
  spreadsheetPath: string | null;
  onOpen: () => void;
  onClear: () => void;
  disabled: boolean;
}

export function ActionBar({ spreadsheetPath, onOpen, onClear, disabled }: ActionBarProps) {
  return (
    <footer className="action-bar">
      <button
        type="button"
        className="btn btn-primary"
        onClick={onOpen}
        disabled={disabled || !spreadsheetPath}
        title={spreadsheetPath ?? "Execute a automacao para habilitar"}
      >
        <FolderIcon /> Abrir planilha
      </button>
      <button
        type="button"
        className="btn btn-secondary"
        onClick={onClear}
        disabled={disabled}
      >
        Limpar log
      </button>
    </footer>
  );
}

function FolderIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M10 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2h-8l-2-2z" fill="currentColor" />
    </svg>
  );
}
