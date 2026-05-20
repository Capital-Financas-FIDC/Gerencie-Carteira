# MetaSpec — Gerencie Carteira
> Contexto para agentes AI. Versao: 1.7 | Atualizado: 2026-05-20

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
IPC:                child_process.spawn (Node); JSON Lines (stdout) + canal stdin (input UI->Python)
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
- **Pasta publica:** `A:\PUBLICA\GERENCIE CARTEIRA PÚBLICA` — rota de rede pre-existente, apenas VERIFICADA, nunca criada; se offline a etapa e pulada com warning
- **PROCX:** aba de cadastro CNPJ->gerente em `config [Excel] sheet_procx` (default `PROCX GERENTES`), colunas `col_procx_gerente` (B) e `col_procx_cnpj` (C)
- **Config:** `config.ini` usa `%USERPROFILE%` literal (lido com `interpolation=None`); `main.ts` tambem parseia o arquivo via regex para a whitelist do `shell.openPath`, a pasta local do dialog e a pasta publica do sweep

## AUTH

Nao aplicavel. Roda com a sessao Windows logada; Outlook COM usa o perfil ativo do usuario.

## REGRAS DE NEGOCIO CRITICAS

### Protocolo JSON Lines (Python → UI)
- `emit()` escreve EXATAMENTE 1 linha JSON valida em stdout (flush imediato); `emit_result()` e o evento terminal (`step: "done"`, `data.result.{status, spreadsheet_path}`)
- Qualquer `print()` solto corrompe o stream — todo output passa por `emit()`
- Linhas nao-parseaveis viram `level: warning` na UI; stderr vira `level: error`. NUNCA derrubam a UI
- Canal reverso (v4.x): `input_bridge.request_input()` emite o request via `emit()` e bloqueia em `sys.stdin.readline()`; a UI escreve 1 linha JSON no `child.stdin` (IPC `script:provideInput`). stdin do spawn deixou de ser `ignore`

### Cascata da planilha base
- `encontrar_arquivo_base_excel()` busca em ordem: (1) `planilhas` local, (2) pasta legada v2.14.1 `Documents\Gerencie Carteira\Diário` (com espaco), (3) share `A:\PUBLICA\...`
- Base encontrada fora do local e COPIADA para o local antes de processar
- Cascata esgotada → emite `excel.base.needs_user` e sai com **EXIT_BASE_NEEDS_USER=4**
- `main.ts`/`useScriptRunner` detecta exit 4 → dialog nativo → copia (renomeando para `Gerencie Carteira_{ontem}.xlsm` se nao bater o pattern) → **auto-rerun completo**
- Pattern `Gerencie Carteira_YYYY_MM_DD.xlsm` esta duplicado em `main.ts` (`BASE_FILENAME_PATTERN`) e no core Python — manter em sincronia

### Invariante de ordem
- Em `main()`, a base DEVE ser resolvida ANTES de `extrair_dados_dos_anexos()`. NAO reordenar; **esta ordem nao tem teste de integracao**
- v4.x: a marcacao `UnRead=False` saiu de `extrair_dados_dos_anexos()` e virou `marcar_emails_lidos()`, chamada SOMENTE apos o save ser promovido com sucesso (reduz a janela de perda; input de orfaos ocorre antes da marcacao)

### Captura de gerentes orfaos (v4.x)
- Antes de colar em `E-Mail BD`, le a aba PROCX → mapa CNPJ→gerente; CNPJs do df ausentes no mapa sao "orfaos" (dedupe)
- Havendo orfaos, request via stdin (ver Protocolo); `{"cancel":true}`/EOF → `CancelExecucao`: nada gravado, e-mails NAO marcados (reprocessaveis)
- Gerentes resolvidos reinjetados no PROCX. A aba e uma Tabela formatada (ListObject): a linha entra via `ListRows.Add` (Tabela expande, colunas calculadas autopreenchem) p/ a referencia estruturada do PROCX em 'E-Mail BD' cobrir a nova linha. Fallback p/ aba sem Tabela: celulas + replica linha anterior. Aba PROCX inexistente → **EXIT_PROCX_MISSING=5**
- DIRECIONA removido: `#N/D` residual e apenas warning — nenhum executavel e aberto

### Planilhas Excel / escrita transacional
- Arquivos DEVEM ser `.xlsm` (preservar macros VBA); `app.calculate()` forcado antes de ler a coluna de verificacao
- Formula VLOOKUP da coluna de verificacao (`config [Excel] coluna_verificacao`, default `E`) copiada da linha anterior (evita `#REF`)
- Base nunca sobrescrita in-place (zero-copy backup). Duas fases (o Excel mantem lock enquanto aberto): `escrever_parcial(wb, destino)` salva `<dest>.partial.<ext>` com o Excel ABERTO; `promover(parcial, destino)` faz `os.replace` (com retry) SO apos `wb.close()`/`app.quit()`. Colisao usa `<dest>.bak.<ext>`
- Publico: `escrever_parcial` (Excel aberto) → fecha → `limpar_publico_antigos()` remove `Gerencie*.xls*` → `promover()` (sem janela de destruicao)
- Recuperacao de kill forcado e do supervisor Electron: varre `.partial`/`.bak` orfaos no boot e no fechamento confirmado

### Versionamento (SemVer)
- Fonte unica da versao: campo `version` de `app/package.json`. NUNCA hardcodar versao em outro lugar (JSX, Python, nome de arquivo, config, README)
- UI le via IPC `app:version`; Python recebe via env `APP_VERSION` (fallback package.json → `"dev"`)
- Entrypoint `backend/src/gerencie_carteira.py` tem nome estavel — nao re-versionar arquivos
- Toda alteracao funcional bumpa a versao (MAJOR/MINOR/PATCH). Checklist completo em `AGENTS.md > Versionamento`

## ESTADO ATUAL (v4.1.0 — 20/05/2026)

Repo unico em git (remote `Capital-Financas-FIDC/Gerencie-Carteira`, branch `main`), SemVer com fonte unica `app/package.json`. **Linha v4** (atual: v4.1.0): captura de gerentes orfaos em runtime + escrita transacional (MAJOR v4.0.0) com refresh automatico de tabelas dinamicas antes do save (MINOR v4.1.0). Patches v4.0.1/v4.0.2 estabilizaram a escrita transacional e a reinjecao via Tabela; hotfix corrigiu o acento ausente na pasta publica (`PÚBLICA`).

**Pronto:**
- Captura de gerentes orfaos em runtime: PROCX → orfaos → formulario Electron (stdin) → reinjecao no PROCX → colagem sem `#N/D`
- Canal stdin bidirecional (`input_bridge`); DIRECIONA totalmente removido (codigo + config)
- Escrita transacional zero-copy em duas fases (`escrever_parcial` + `promover` apos fechar o Excel); base nunca sobrescrita in-place
- Refresh automatico de TODAS as PivotTables do workbook apos a colagem e antes do save (`atualizar_pivots`) — copia publica sai ja atualizada p/ consulta direta
- Confirmacao SIM/NAO ao fechar em runtime + sweep de `.partial`/`.bak` no supervisor (boot/fechamento)
- `UnRead=False` adiado para apos o save (janela de perda reduzida)
- Cascata da planilha base + auto-rerun via dialog Electron (exit 4) inalterada
- pytest passing (log_emitter, directory_bootstrap, cascata_base, gerentes_orfaos, rollback_transacional, input_bridge)
- Electron 33 + React 18 + Vite 5; tsc/Vite limpos

**Dividas tecnicas:**
- Invariante de ordem em `main()` sem teste de integracao — regressao passaria despercebida
- `input_bridge.peek_cancel()` e no-op em Windows (select nao funciona em pipe); cancelamento autoritativo depende do supervisor Electron
- Falha de `wb.save()` apos input do usuario perde o que foi digitado (e-mails seguem nao lidos → rerun re-solicita)
- Sem testes de UI (Vitest) nem E2E real com Outlook; smoke E2E nao executado
- `atualizar_planilha_excel()` cresceu (~7 responsabilidades), sem refatorar
- Installer NSIS bloqueado nesta maquina (symlink winCodeSign); distribuicao = pasta `./Aplicativo` (regeneravel)
