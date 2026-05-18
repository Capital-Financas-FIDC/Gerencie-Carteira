# MetaSpec — Gerencie Carteira
> Contexto para agentes AI. Versao: 1.1 | Atualizado: 2026-05-18

## IDENTIDADE

- **Nome:** Gerencie Carteira (`gerencie-carteira-app`)
- **Dominio:** Automacao de monitoramento empresarial (FIDC / credito)
- **Proposito:** Extrair dados de monitoramento do Serasa via e-mails Outlook, parsear HTML e atualizar planilhas Excel macro-habilitadas (.xlsm), com UI desktop para operacao
- **Usuarios-alvo:** Equipe comercial de FIDC (gestores de carteira de credito)
- **Idioma:** Portugues Brasileiro (codigo, docs, UI)

## STACK

```
Runtime (backend):  Python 3.x (Windows-only)
Runtime (UI):       Electron 33 + React 18 + Vite 5 + TypeScript 5
COM:                pywin32 (Outlook MAPI) + xlwings (Excel .xlsm com macros VBA)
Parsing:            BeautifulSoup4 (HTML do Serasa)
Dados:              pandas (DataFrames)
IPC:                child_process.spawn (Node) + JSON Lines (stdout)
Tema:               prefers-color-scheme (CSS vars)
Build:              PyInstaller (core.exe) + electron-builder (NSIS / portable)
Testes:             pytest (backend); UI sem testes automatizados
Versionamento:      SemVer; fonte unica = app/package.json; repo unico
```

## ARQUITETURA

```
Serasa (email) → Outlook (COM) → Python core (script ou PyInstaller .exe)
                                      │  stdout: 1 objeto JSON por linha
                              Electron main.ts (spawn + readline)
                                      │  webContents.send('script:log' | 'script:done')
                               React renderer (useScriptRunner reducer)
                                      │  xlwings (COM) controla o Excel
                                   ↙           ↘
                          Diario .xlsm    Copia publica .xlsm
```

| Camada | Diretorio | Responsabilidade |
|--------|-----------|------------------|
| Backend core | `backend/src/` | Pipeline Python (config → bootstrap → Outlook → parse → Excel) |
| Protocolo | `backend/src/log_emitter.py` | Unica forma sancionada de output Python→UI |
| Bootstrap | `backend/src/directory_bootstrap.py` | Workspace idempotente + verificacao de share |
| Testes | `backend/tests/` | pytest: log_emitter, directory_bootstrap, cascata_base |
| Electron main | `app/electron/` | spawn, readline IPC, dialog, shell.openPath (whitelist) |
| UI React | `app/src/` | Componentes, `useScriptRunner` (reducer de estado) |
| Config | `config/config.ini` | Defaults universais via `%USERPROFILE%` |
| Build | `backend/build_core.ps1`, `app/package.json` | PyInstaller + electron-builder |

Pipeline v3.0.1:

```
boot → config.loaded → workspace.bootstrap → outlook.fetch
     → excel.base.* (cascata)  → html.parse → excel.open → excel.insert
     → excel.save → publico.copy → done (emit_result)
```

## DADOS

- **Fonte:** E-mails do Serasa no Outlook (assunto exato em `config.ini [Email]`)
- **Entrada:** Anexos HTML com tabela (CNPJ, Razao Social, Alteracao)
- **Saidas:** 2 planilhas `.xlsm` — diaria local (`\planilhas`) + copia publica em rede
- **Workspace:** `%USERPROFILE%\Documents\Gerencie_Carteira\{planilhas,html,logs}` — criado idempotentemente a cada run (`ensure_workspace`)
- **Pasta publica:** `A:\PUBLICA\GERENCIE CARTEIRA PUBLICA` — rota de rede pre-existente, apenas VERIFICADA, nunca criada; se offline a etapa e pulada com warning
- **Config:** `config.ini` usa `%USERPROFILE%` literal (lido com `interpolation=None`); `main.ts` tambem parseia o arquivo via regex para a whitelist do `shell.openPath` e para resolver a pasta local no dialog

## AUTH

Nao aplicavel. Roda com a sessao Windows logada; Outlook COM usa o perfil ativo do usuario.

## REGRAS DE NEGOCIO CRITICAS

### Protocolo JSON Lines (Python → UI)
- `emit()` escreve EXATAMENTE 1 linha JSON valida em stdout (flush imediato); `emit_result()` e o evento terminal (`step: "done"`, `data.result.{status, spreadsheet_path}`)
- Qualquer `print()` solto corrompe o stream — todo output passa por `emit()`
- Linhas nao-parseaveis viram `level: warning` na UI; stderr vira `level: error`. NUNCA derrubam a UI

### Cascata da planilha base
- `encontrar_arquivo_base_excel()` busca em ordem: (1) `planilhas` local, (2) pasta legada v2.14.1 `Documents\Gerencie Carteira\Diário` (com espaco), (3) share `A:\PUBLICA\...`
- Base encontrada fora do local e COPIADA para o local antes de processar
- Cascata esgotada → emite `excel.base.needs_user` e sai com **EXIT_BASE_NEEDS_USER=4**
- `main.ts`/`useScriptRunner` detecta exit 4 → dialog nativo → copia (renomeando para `Gerencie Carteira_{ontem}.xlsm` se nao bater o pattern) → **auto-rerun completo**
- Pattern `Gerencie Carteira_YYYY_MM_DD.xlsm` esta duplicado em `main.ts` (`BASE_FILENAME_PATTERN`) e no core Python — manter em sincronia

### Invariante de ordem (a razao da v3.0.1)
- Em `main()`, a base DEVE ser resolvida ANTES de `extrair_dados_dos_anexos()` — esta ultima marca e-mails `UnRead=False`. Se a cascata saisse apos consumir e-mails, o auto-rerun nao acharia e-mails nao lidos e perderia o dia. NAO reordenar; **esta ordem nao tem teste**

### Planilhas Excel
- Arquivos DEVEM ser `.xlsm` (preservar macros VBA)
- Formula VLOOKUP da coluna de verificacao (`config [Excel] coluna_verificacao`, default `E`) e copiada da linha anterior para novas linhas (evita `#REF`)
- `#N/D` = CNPJ sem cadastro → status "pendencias" (warning + abre `executavel_direciona` se existir)
- v3.0.1 forca `app.calculate()` antes de ler a coluna (evita falso `#N/D` por recalculo assincrono)
- Copia publica: arquivos `Gerencie*.xls*` antigos sao deletados antes de salvar o novo

### Versionamento (SemVer)
- Fonte unica da versao: campo `version` de `app/package.json`. NUNCA hardcodar versao em outro lugar (JSX, Python, nome de arquivo, config, README)
- UI le via IPC `app:version`; Python recebe via env `APP_VERSION` (fallback package.json → `"dev"`)
- Entrypoint `backend/src/gerencie_carteira.py` tem nome estavel — nao re-versionar arquivos
- Toda alteracao funcional bumpa a versao (MAJOR/MINOR/PATCH). Checklist completo em `AGENTS.md > Versionamento`

## ESTADO ATUAL (v3.0.1 — 18/05/2026)

Repo unico com versionamento SemVer (fonte unica `app/package.json`). Release v3.0.1 em validacao inicial.

**Pronto:**
- Backend Python v3 com pipeline JSON Lines funcional (smoke OK)
- Cascata da planilha base + auto-rerun via dialog Electron (exit 4)
- Resolucao da base ANTES de marcar e-mails lidos (fix de perda de dados v3.0.1)
- Recalculo forcado do Excel antes da checagem de `#N/D` (v3.0.1)
- Bootstrap idempotente do workspace; verificacao graciosa do share publico
- Deteccao informativa da pasta legada v2.14.1
- pytest passing (log_emitter + directory_bootstrap + cascata_base)
- Electron 33 + React 18 + Vite 5 com contextIsolation; UI completa (RunButton, StatusBadge, LogViewer, ActionBar)
- Tema responsivo (prefers-color-scheme); build Vite e PyInstaller onefile funcionais

**Dividas tecnicas:**
- Risco residual: e-mails marcados como lidos ANTES de `wb.save()` — se o save falha, o dia e perdido (fix exige adiar `UnRead=False` para apos o save)
- Invariante de ordem em `main()` (base antes de marcar lidos) sem teste — regressao passaria despercebida
- Sem testes de UI (Vitest planejado) nem de integracao IPC
- Smoke E2E com Outlook real ainda nao executado
- `atualizar_planilha_excel()` com ~6 responsabilidades, sem refatorar; type hints incompletos no pipeline
- Sem Git na raiz — `git init` pendente (release/rollback SemVer dependem de tags `vX.Y.Z`)
- Installer NSIS gerado mas nao validado em maquina limpa
