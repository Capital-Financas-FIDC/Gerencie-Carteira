# Context Index — Gerencie Carteira
> Mapa de artefatos. Atualizado: 2026-06-08 (v1.3)

## Navegacao Rapida

| Artefato | Caminho | Descricao |
|----------|---------|-----------|
| CONTEXT_SPEC | `context/CONTEXT_SPEC.md` | Spec canonico imutavel dos docs de contexto |
| MetaSpec | `context/metaspec.md` | Identidade, stack, arquitetura, regras, estado atual |
| Index | `context/index.md` | Este mapa de navegacao e arquivos criticos |
| Timeline | `context/timeline.md` | Historia evolutiva por fases |
| AGENTS.md | `AGENTS.md` | Sistema operacional do repo: comandos, arquitetura, versionamento SemVer (CLAUDE.md delega para ele) |

## Artefatos Ativos

### Analises — context/analysis/
Nenhum artefato. (Historico em `context/analysis/old/`.)

### Walkthroughs — context/walkthroughs/
Nenhum artefato.

### Plans — plans/
| Arquivo | Data | Descricao | Status |
|---------|------|-----------|--------|
| `2026-06-08_Plan_Performance_Fase1.md` | 2026-06-08 | Fase de performance: instrumentacao + Outlook Restrict (Fases 1-3 = v4.2.8); Fase 4 (calculo manual/rebuild condicional = v4.2.9) e Fase 5 (launcher local) | Fases 1-5 implementadas; aguarda validacao na mesa |

## Arquivos Criticos

### Backend core
| Arquivo | Responsabilidade |
|---------|------------------|
| `backend/src/gerencie_carteira.py` | Entrypoint do pipeline (nome estavel, sem versao); `main()` define a invariante de ordem |
| `backend/src/log_emitter.py` | Contrato JSON Lines (`emit`, `emit_result`) — unica saida sancionada |
| `backend/src/directory_bootstrap.py` | Workspace idempotente, verificacao de share, deteccao de pasta legada |

### Electron / UI
| Arquivo | Responsabilidade |
|---------|------------------|
| `app/electron/main.ts` | spawn do core, readline IPC, dialog da cascata, whitelist de `shell.openPath` |
| `app/electron/preload.ts` | contextBridge — superficie `window.electronAPI` |
| `app/src/hooks/useScriptRunner.ts` | Reducer de estado; deteccao de exit 4 e auto-rerun |
| `app/src/App.tsx` | Composicao da UI |

### Config / Build
| Arquivo | Responsabilidade |
|---------|------------------|
| `config/config.ini` | Pastas de trabalho na rede, aba/coluna do Excel, assunto do e-mail, `[Retencao]` |
| `app/package.json` | Versao (fonte unica SemVer), scripts dev/build, config electron-builder |
| `backend/build_core.ps1` | PyInstaller → `app/resources/gerencie_carteira_core.exe` |
| `build-app.ps1` | Build completo (raiz) → entrega o app em `./Aplicativo/` |
| `publicar.ps1` | Build + espelha p/ `A:\...\Software` (deploy na rede); versao no nome do `.exe` |
| `Gerencie Carteira.cmd` | Atalho de um clique p/ `Aplicativo\Gerencie Carteira.exe` (build local) |

## Testes

| Camada | Diretorio | Status |
|--------|-----------|--------|
| Backend (unit) | `backend/tests/` | passing |
| UI / IPC | — | inexistente (Vitest planejado) |

## Documentacao Tecnica

| Arquivo | Descricao |
|---------|-----------|
| `README.md` | Layout do projeto, comandos dev/build/deploy, protocolo JSON Lines |
| `AGENTS.md` | Guia operacional para agentes + governanca de versionamento SemVer |
| `CLAUDE.md` | Stub que delega integralmente para `AGENTS.md` |
