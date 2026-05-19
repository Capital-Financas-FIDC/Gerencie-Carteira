import { useMemo, useState } from "react";
import type { OrfaoEntry } from "../types/log";

interface GerentesOrfaosFormProps {
  orfaos: OrfaoEntry[];
  onSubmit: (mapping: Record<string, string>) => void;
}

/**
 * Formulario em runtime para resolver CNPJs sem gerente no PROCX.
 * O botao Confirmar so habilita quando TODOS os campos estao preenchidos.
 * Ao confirmar, devolve o mapeamento { cnpj: gerente } ao core via IPC.
 */
export function GerentesOrfaosForm({ orfaos, onSubmit }: GerentesOrfaosFormProps) {
  const [valores, setValores] = useState<Record<string, string>>({});

  const todosPreenchidos = useMemo(
    () => orfaos.length > 0 && orfaos.every((o) => (valores[o.cnpj] ?? "").trim() !== ""),
    [orfaos, valores],
  );

  const handleChange = (cnpj: string, valor: string) => {
    setValores((prev) => ({ ...prev, [cnpj]: valor }));
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!todosPreenchidos) return;
    const mapping: Record<string, string> = {};
    for (const o of orfaos) mapping[o.cnpj] = valores[o.cnpj].trim();
    onSubmit(mapping);
  };

  return (
    <div className="orfaos-overlay" role="dialog" aria-modal="true" aria-label="Gerentes orfaos">
      <form className="orfaos-modal" onSubmit={handleSubmit}>
        <header className="orfaos-header">
          <h2>Gerentes nao cadastrados</h2>
          <p>
            {orfaos.length} CNPJ(s) sem gerente no PROCX. Informe o gerente de cada
            um para continuar — os dados ainda nao foram gravados.
          </p>
        </header>

        <div className="orfaos-list">
          {orfaos.map((o) => (
            <div className="orfaos-row" key={o.cnpj}>
              <div className="orfaos-info">
                <span className="orfaos-cnpj">{o.cnpj}</span>
                <span className="orfaos-razao">{o.razao_social || "—"}</span>
              </div>
              <input
                type="text"
                className="orfaos-input"
                placeholder="Nome do gerente"
                value={valores[o.cnpj] ?? ""}
                onChange={(e) => handleChange(o.cnpj, e.target.value)}
                autoComplete="off"
                spellCheck={false}
              />
            </div>
          ))}
        </div>

        <footer className="orfaos-footer">
          <button type="submit" className="btn btn-primary" disabled={!todosPreenchidos}>
            Confirmar e continuar
          </button>
        </footer>
      </form>
    </div>
  );
}
