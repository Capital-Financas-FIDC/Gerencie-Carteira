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

## Fase 2: Migracao Electron + correcao/performance (v3.0.0 → v4.2.10) — atual

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
- PATCH v4.2.2: revertida a versao no nome do `.exe` distribuido. Electron 33 crasha (`0x80000003` / ASAR integrity) quando o exe principal e renomeado — confirmado no Event Viewer apos o primeiro deploy do v4.2.1 (`Application Error 1000`, "Gerencie Carteira 4.2.1.exe"). `publicar.ps1` agora preserva o nome neutro `Gerencie Carteira.exe`; a versao do app continua visivel via `app:version` na UI
- PATCH v4.2.6: `limpar_publico_antigos` engolia `OSError` silenciosamente — quando algum gestor estava com a planilha publica antiga aberta no Excel, o `os.remove` falhava sem deixar rastro e a planilha do dia anterior ficava ao lado da nova (acumulo na pasta publica, que deve ter no maximo uma). Agora a funcao faz retry curto (4 tentativas com backoff) e retorna a lista de falhas; o caller emite `publico.cleanup.fail` (warning) por arquivo nao removido. Adicionado tambem um sweep antecipado antes de salvar o parcial publico (best-effort, falhas silenciosas) — segunda janela caso o usuario feche a planilha entre o inicio do pipeline e a promocao. Novo teste cobre o retorno de falhas
- PATCH v4.2.5: app rodando da rede crashava no boot com `0x80000003` (sintoma: "abre branco e fecha em ~1.5s"). Captura via `--enable-logging=stderr` mostrou `gpu_process_host: error_code=18` 10x seguidas e `FATAL: GPU process isn't usable. Goodbye` — o sandbox do GPU process do Chromium nao inicializa em network drive. Fix cirurgico em `app/electron/main.ts`: `app.commandLine.appendSwitch("disable-gpu-sandbox")` no topo (renderer/utility sandboxes permanecem ligados). Bonus: `build-app.ps1` ganhou limpeza do `win-unpacked` antes do `electron-builder` — o `try/catch` engolia falhas do builder e o `Test-Path` reusava silenciosamente o build velho, mascarando builds quebrados (descoberto quando uma tentativa de adicionar `electronFuses` ao `package.json` foi rejeitada pelo schema do electron-builder 25.1.8 mas o `publicar.ps1` reportou sucesso publicando o binario antigo). Bumps 4.2.3/4.2.4 foram iteracoes intermediarias de debugging sem mudanca efetiva no binario

- PATCH v4.2.7: pivot toda `#NOME?` com orfaos — apos `ListRows.Add` o `Calculate` nao reconstruia a arvore e o `RefreshTable` congelava o erro; fix forca `CalculateFullRebuild` antes das pivots
- PATCH v4.2.7: reinjecao no PROCX passa a preencher colunas de FORMULA (A/D/E) replicando a linha anterior via `Range.Copy` (`ListRows.Add` so autopreenche colunas calculadas registradas)
- MINOR v4.2.8: fetch do Outlook troca a varredura O(N) da Inbox por `Items.Restrict` (filtro server-side) — gargalo dos "15 min" em caixa grande/compartilhada; fallback p/ varredura antiga
- MINOR v4.2.8: instrumentacao de tempo por etapa (`timings_<run>.json`); launcher `.cmd` copia o app p/ `%LOCALAPPDATA%` e roda local, recopiando so quando `versao.txt` muda (evita stream pela rede a cada uso)
- PATCH v4.2.9: calculo MANUAL na automacao (open/edicoes mais rapidos) + `recalcular(full=...)` — full rebuild SO com orfaos, senao Calculate normal; restaura AUTOMATICO antes de salvar (copia publica recalcula p/ o gestor)
- PATCH v4.2.10: corrige regressao da v4.2.9 — o rebuild condicional deixava todo dia SEM orfao com a pivot inteira em `#NOME?`. A coluna Gerente usa XLOOKUP (`_xlfn.XLOOKUP`) copiada p/ linhas novas a cada run; `app.calculate()` normal nao religa o token -> `#NOME?` congelado no cache da pivot. `recalcular(app, full=True)` volta a ser incondicional. Diagnostico por forense cruzada das planilhas 06/06 (full rebuild, limpa) vs 06/09 (calculate, 30x `#NAME?`) + timings
- PATCH v4.2.11: `#NOME?` persistia nas maquinas do cadastro (Excel 2019) mesmo com o full rebuild — corrida de timing. `CalculateFullRebuild` devolve o controle ao COM ANTES de a propagacao multithread terminar em maquina lenta; o `RefreshTable` seguinte fotografa o `#NOME?` transitorio (`CalculateUntilAsyncQueriesDone` so espera queries externas). Fix: `recalcular` polla `Application.CalculationState` ate `xlDone` (timeout 30s) antes de retornar. Confirmado que o XLOOKUP calcula no Excel 2019 (nao e incompat. de versao)
- PATCH v4.2.12: botao "Abrir Planilha" morria em silencio. `log_emitter` usava `json.dumps(ensure_ascii=False)`; o exe PyInstaller ignora `PYTHONIOENCODING` e escreve stdout em cp1252 (mesmo sintoma do mojibake de stdin da v4.1.1), entao o `Ú` de `...GERENCIE CARTEIRA PÚBLICA...` chegava corrompido ao readline UTF-8 do `main.ts` -> `fs.existsSync`/whitelist falhavam e o React engolia o erro. Fix: `ensure_ascii=True` (wire ASCII puro, codepage do stdout irrelevante) + `openSpreadsheet` exibe o motivo da falha no log. Quebrou na v4.2.0, quando os paths migraram p/ a pasta de rede acentuada

## Metricas Snapshot (2026-06-11)

| Metrica | Valor |
|---------|-------|
| Versao atual | v4.2.12 |
| Versionamento | SemVer; fonte unica `app/package.json`; repo unico |
| Releases historicas | ~40 (v1.0.0 → v4.2.12, lineares no mesmo git); tags v4.x comecam em v4.2.11 |
| Linguagens | Python (core), TypeScript/React (UI) |
| Testes backend | passing (~10 suites pytest) |
| Testes UI | inexistentes |
| Historico git | remote `Capital-Financas-FIDC/Gerencie-Carteira`; trabalho recente na branch `fix/Consertando-Planilha` (sem merge em `main`; nao pushada) |
