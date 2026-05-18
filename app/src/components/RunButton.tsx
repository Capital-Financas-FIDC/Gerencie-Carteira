interface RunButtonProps {
  running: boolean;
  disabled?: boolean;
  onClick: () => void;
}

export function RunButton({ running, disabled, onClick }: RunButtonProps) {
  return (
    <button
      type="button"
      className="run-button"
      onClick={onClick}
      disabled={disabled}
      data-running={running}
      aria-live="polite"
    >
      {running ? (
        <>
          <span className="spinner" aria-hidden="true" />
          <span>Executando...</span>
        </>
      ) : (
        <>
          <PlayIcon />
          <span>Executar</span>
        </>
      )}
    </button>
  );
}

function PlayIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M8 5v14l11-7L8 5z" fill="currentColor" />
    </svg>
  );
}
