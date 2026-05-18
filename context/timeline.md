# Timeline — Gerencie Carteira
> Historia evolutiva. 3 fases | 2025 → 2026-05. Historico git nao disponivel (sem `.git` na raiz); fases derivadas do versionamento manual por diretorios.

## Fase 0: CLI single-file (v1.0.0 → v2.13.0)

- Script Python unico, Windows-only: Outlook (COM) → parse HTML do Serasa → atualiza planilha `.xlsm`
- Pipeline modular em ~9 funcoes; interface de console (Rich)
- Estabilizacao da automacao Excel via xlwings preservando macros VBA
- ~30 releases manuais acumuladas em `versions/main/MAJOR.MINOR/PATCH/`

## Fase 1: Endurecimento da CLI (v2.14.x)

- Fix do VLOOKUP: formula da coluna de verificacao copiada da linha anterior para evitar `#REF`
- Empacotamento via PyInstaller; v2.14.1 consolidada como release operacional estavel
- v2.14.1 mantida intocada como fallback paralelo ao migrar para v3

## Fase 2: Migracao Electron (v3.0.0 → v3.0.1) — atual

- MAJOR v3.0.0: CLI Rich substituida por protocolo JSON Lines (stdout) consumido por shell Electron + React
- Bootstrap idempotente do workspace `%USERPROFILE%\Documents\Gerencie_Carteira\`; defaults universais via `%USERPROFILE%`
- Cascata da planilha base + auto-rerun via dialog nativo (exit code 4)
- PATCH v3.0.1: base resolvida ANTES de marcar e-mails lidos (corrige perda de dados); recalculo forcado do Excel antes da checagem `#N/D`
- Inicializacao da estrutura de contexto (CONTEXT_SPEC, metaspec, index, timeline); `CLAUDE.md` passa a delegar para `AGENTS.md`
- Consolidacao em repo unico (fim da arvore `versions/`) + governanca SemVer: fonte unica `app/package.json`, entrypoint renomeado para nome estavel `gerencie_carteira.py`, versao propagada para UI (IPC) e Python (env)
- `git init` + enxerto sobre o historico legacy: v3.0.1 vira filho de v2.13.0 (linear, sem perda de commits/tags) no remote da org `Capital-Financas-FIDC/Gerencie-Carteira`

## Metricas Snapshot (2026-05-18)

| Metrica | Valor |
|---------|-------|
| Versao atual | v3.0.1 |
| Versionamento | SemVer; fonte unica `app/package.json`; repo unico |
| Releases historicas | ~30 (v1.0.0 → v3.0.1, agora linear no mesmo git) |
| Linguagens | Python (core), TypeScript/React (UI) |
| Testes backend | passing (~3 suites pytest) |
| Testes UI | inexistentes |
| Historico git | disponivel — remote `Capital-Financas-FIDC/Gerencie-Carteira`, branch `main` |
