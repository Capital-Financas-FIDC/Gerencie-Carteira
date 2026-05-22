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

## Fase 2: Migracao Electron (v3.0.0 → v4.0.2) — atual

- MAJOR v3.0.0: CLI Rich substituida por protocolo JSON Lines (stdout) consumido por shell Electron + React
- Bootstrap idempotente do workspace `%USERPROFILE%\Documents\Gerencie_Carteira\`; defaults universais via `%USERPROFILE%`
- Cascata da planilha base + auto-rerun via dialog nativo (exit code 4)
- PATCH v3.0.1: base resolvida ANTES de marcar e-mails lidos (corrige perda de dados); recalculo forcado do Excel antes da checagem `#N/D`
- Inicializacao da estrutura de contexto (CONTEXT_SPEC, metaspec, index, timeline); `CLAUDE.md` passa a delegar para `AGENTS.md`
- Consolidacao em repo unico (fim da arvore `versions/`) + governanca SemVer: fonte unica `app/package.json`, entrypoint renomeado para nome estavel `gerencie_carteira.py`, versao propagada para UI (IPC) e Python (env)
- `git init` + enxerto sobre o historico legacy: v3.0.1 vira filho de v2.13.0 (linear, sem perda de commits/tags) no remote da org `Capital-Financas-FIDC/Gerencie-Carteira`
- Pipeline de build usavel: `build-app.ps1` entrega o app empacotado em `./Aplicativo/` (fora de `build/`) + launcher de um clique; `tsconfig` com `noEmit` (tsc so checa tipos, Vite bundla)
- MAJOR v4.0.0: fluxo `#N/D`/DIRECIONA substituido por captura de gerentes orfaos em runtime. Novo canal stdin bidirecional (`input_bridge`), formulario Electron/React, reinjecao no PROCX (colunas replicam linha anterior). DIRECIONA removido (codigo+config). Escrita transacional zero-copy (`xlsm_transacional` + sweep no supervisor). `UnRead=False` adiado para apos o save. Novas suites pytest (orfaos, rollback, input_bridge). MAJOR (e nao MINOR como rascunhado no prompt) por quebra operacional p/ usuario + `config.ini` incompativel
- PATCH v4.0.1: dois bugs do xlwings na escrita transacional. (1) marcador movido para ANTES da extensao (`.partial.xlsm`/`.bak.xlsm`) — o Excel deduz o formato pela extensao e quebrava com `.xlsm.partial`. (2) API dividida em duas fases: `escrever_parcial(wb,...)` com o Excel aberto e `promover(parcial,destino)` SO apos `wb.close()`/`app.quit()` (o Excel mantem lock no arquivo salvo -> `os.replace` dava `[WinError 32]`); `promover` faz retry no release preguicoso. Regex de sweep ajustada (Python + Electron)
- PATCH v4.0.2: reinjecao no PROCX via API da Tabela. A aba e um ListObject e a coluna E de 'E-Mail BD' usa referencia estruturada; escrever celulas abaixo da Tabela NAO a expandia -> PROCX nao cobria a nova linha -> `#N/D`. Agora `reinjetar_procx` usa `ListObject.ListRows.Add` (Tabela expande, colunas calculadas autopreenchem); config opcional `tabela_procx`; fallback de celulas se a aba nao tiver Tabela
- HOTFIX (post-v4.0.2): nome da pasta publica no `config.ini` estava sem acento (`GERENCIE CARTEIRA PUBLICA`); a pasta real no servidor e `GERENCIE CARTEIRA PÚBLICA`. `verify_public_path` retornava `accessible=False` em toda execucao -> etapa publica era pulada. Correcao em `config/config.ini` + docs.
- MINOR v4.1.0: `atualizar_pivots(wb)` itera todas as PivotTables de todas as abas e chama `RefreshTable()` apos a colagem em E-Mail BD e antes da escrita transacional. Garante que a copia publica ja sai com as tabelas dinamicas atualizadas (gestor abre e consulta direto). Falha em uma pivot vira warning; ausencia de pivots emite info, sem quebrar o pipeline
- PATCH v4.1.1: corrige mojibake no input de gerentes orfaos. No PyInstaller `--onefile`, `sys.stdin` em modo texto ignorava `PYTHONIOENCODING=utf-8` e caia em cp1252 -> "ADMINISTRAÇÃO" virava "ADMINISTRAÃ‡ÃƒO" na planilha. `input_bridge.request_input` agora le `sys.stdin.buffer` (bytes) e decodifica UTF-8 explicitamente; UnicodeDecodeError vira warning, nao derruba o stream
- MINOR v4.2.0: retencao automatica de backups. `limpar_backups_antigos()` mantem no maximo N arquivos (config `[Retencao] max_arquivos`, default 30 ~ 1 mes) nas pastas `planilhas` e `html`, removendo os mais antigos pela data `YYYY_MM_DD` capturada no nome (ignora `.partial`/`.bak`/templates). Chamada em `main()` apos cada save bem-sucedido. No mesmo release, `test_input_bridge` foi corrigido — o stub de `sys.stdin` nao expunha `.buffer`, ficando defasado desde o v4.1.1
- v4.2.0 (distribuicao): projeto migra para a pasta de rede. `config.ini` repontado para `...\GERENCIE CARTEIRA PÚBLICA\Software\data`; `ensure_workspace` passa a receber a raiz derivada do config. Novo `publicar.ps1` builda e espelha o app para `Software\Aplicativo` (robocopy `/MIR` remove o build antigo), com a versao no nome do `.exe`. Atalho `.cmd` com curinga na pasta-pai. Instalador NSIS abandonado (admin nao liberou Developer Mode/symlink)
- PATCH v4.2.1: isolamento dev/producao. `aplicar_pastas_dev` (core Python) e os ramos `isDev` do `main.ts` redirecionam as pastas de trabalho para `<repo>\data` quando o app roda do codigo-fonte — execucoes de teste deixam de gravar na pasta de producao da rede. O app empacotado e inalterado (usa o `config.ini`)

## Metricas Snapshot (2026-05-22)

| Metrica | Valor |
|---------|-------|
| Versao atual | v4.2.1 |
| Versionamento | SemVer; fonte unica `app/package.json`; repo unico |
| Releases historicas | ~32 (v1.0.0 → v4.2.1, agora linear no mesmo git) |
| Linguagens | Python (core), TypeScript/React (UI) |
| Testes backend | passing (~8 suites pytest) |
| Testes UI | inexistentes |
| Historico git | disponivel — remote `Capital-Financas-FIDC/Gerencie-Carteira`, branch `main` |
