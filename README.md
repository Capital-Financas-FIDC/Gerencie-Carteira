# Gerencie Carteira

> **Repo unico, versionamento SemVer.** A versao corrente vive em
> `app/package.json` (fonte unica) e e exibida na UI e nos logs em runtime.
> Politica de bump e checklist de release: ver `AGENTS.md > Versionamento`.
> O historico evolutivo fica em `context/timeline.md` e no git.

Aplicativo desktop Electron + React sobre um core Python (Outlook COM + xlwings):
automacao via interface grafica com botao de acao, indicador de status, tema
responsivo ao sistema (dark/light) e log estilizado em tempo real.

## Layout do release

```
.
├── app/          # Electron + React + Vite (UI)
│   ├── electron/ # main.ts (spawn Python + IPC) + preload.ts (contextBridge)
│   ├── src/      # React: components, hooks, styles, types
│   └── resources/# icon.ico + gerencie_carteira_core.exe (produto do PyInstaller)
├── backend/
│   ├── src/      # Pipeline Python v3 (JSON Lines)
│   │   ├── gerencie_carteira.py          # Entrypoint (nome estavel, sem versao)
│   │   ├── log_emitter.py               # stdout JSON helper
│   │   └── directory_bootstrap.py       # ensure_workspace idempotente
│   ├── tests/    # pytest (log_emitter, directory_bootstrap)
│   ├── pytest.ini
│   └── build_core.ps1                   # PyInstaller -> app/resources/
├── config/
│   └── config.ini                       # Defaults universais (%USERPROFILE%)
├── log/          # Logs locais (runtime)
└── build/
    ├── dist-python/      # Saida PyInstaller
    └── dist-electron/    # Saida electron-builder
```

## Estrutura de trabalho do usuario

Ao iniciar, o app cria idempotentemente:

```
%USERPROFILE%\Documents\Gerencie_Carteira\
├── planilhas\   # Excel diario (.xlsm)
├── html\        # Anexos brutos do Serasa
└── logs\        # Log de execucao
```

A pasta publica `A:\PUBLICA\GERENCIE CARTEIRA PUBLICA` e rota de rede pre-existente e NAO e criada pelo app — apenas verificada, com warning gracioso se offline.

## Dev

```bash
# Backend puro (emite JSON Lines no stdout)
cd backend
python -m pytest                                    # 24/24 tests
python src/gerencie_carteira.py                     # smoke real

# Electron dev (hot reload)
cd app
npm install
npm run dev                                         # abre janela Electron
```

## Build

```bash
# 1. Compilar o core Python (PyInstaller)
cd backend
./build_core.ps1                                    # -> app/resources/gerencie_carteira_core.exe

# 2. Compilar o app Electron
cd app
npm run build                                       # transpila TS + bundle Vite

# 3. Gerar distribucao
npm run dist            # installer NSIS (*)
npm run dist:dir        # pasta win-unpacked/ (portable manual)
npm run dist:portable   # zip portable autocontido
```

**(*) Nota sobre NSIS installer em Windows:** O electron-builder extrai cache `winCodeSign` que contem 2 symlinks de dylibs macOS irrelevantes para Windows. Se o usuario nao tiver privilegio para criar symlinks (comum em maquinas corporativas), a etapa final de gerar o .exe NSIS falha mas o `win-unpacked/` e gerado corretamente. Alternativas:

1. **Usar `dist:portable` ou `dist:dir`** — produz pasta autocontida que funciona sem instalacao
2. **Ativar Developer Mode no Windows** (`Configuracoes > Atualizacao e Seguranca > Para Desenvolvedores`) — permite symlinks sem admin, resolve o NSIS
3. **Rodar PowerShell como Admin** antes do `npm run dist`

## Testes

```bash
cd backend
python -m pytest    # log_emitter (6) + directory_bootstrap (8) + cascata_base (10)
```

## Rollback

Repo unico: cada release publicado e uma tag SemVer no git (`vMAJOR.MINOR.PATCH`).
Para voltar a uma versao anterior, faca checkout da tag correspondente e
recompile (`build_core.ps1` + `npm run dist*`). A pasta legada de workspace
`Documents\Gerencie Carteira` (com espaco) continua sendo detectada pela
cascata da planilha base — ver `AGENTS.md`.

## Protocolo JSON Lines (backend → UI)

Cada evento emitido e um objeto JSON em 1 linha (`\n` terminado, `flush=True`). Schema:

```typescript
{
  level: "info" | "success" | "warning" | "error" | "step",
  ts:    string,              // ISO-8601 com timezone local
  msg:   string,              // Texto humano
  step?: string,              // "outlook.fetch", "excel.save", ...
  progress?: number,          // 0..100 (opcional)
  data?: Record<string, any>  // Payload livre
}
```

Eventos terminais: `{step: "done", data: {result: {status, spreadsheet_path}}}`.
