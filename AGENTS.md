# AGENTS.md

This file provides guidance to any AI agent when working with code in this repository.

> All code comments, docs, and UI text are in **Brazilian Portuguese**. This
> is a **single repository** (no more per-version repos / `versions/` tree).
> The product is an Electron + React shell over a Python core. Versioning is
> **SemVer** with a single source of truth — see `## Versionamento (SemVer)`.
> Treat this file as authoritative; the workspace-level `..\AGENTS.md` is
> outdated.

## What this is

Windows-only desktop automation: extracts Serasa monitoring data from Outlook
emails → parses HTML attachments → appends rows to a macro-enabled Excel
workbook (`.xlsm`) → publishes a copy to a network share. v3.0.1 wraps the
Python pipeline in an Electron GUI; the Python core is unchanged from the CLI
lineage except that it now emits JSON Lines instead of Rich console output.

## Commands

```bash
# Backend (Python core) — emits JSON Lines on stdout
cd backend
python -m pytest                          # 24 tests (log_emitter, directory_bootstrap, cascata_base)
python -m pytest tests/test_cascata_base.py::test_name -v   # single test
python src/gerencie_carteira.py           # real smoke run (needs Outlook + Excel)

# Electron app (UI)
cd app
npm install
npm run dev          # Vite + Electron with hot reload
npm run build        # tsc -b && vite build (run before any electron-builder step)

# Build a distributable
cd backend && ./build_core.ps1            # PyInstaller -> app/resources/gerencie_carteira_core.exe
cd app && npm run dist                    # NSIS installer (see NSIS caveat in README)
cd app && npm run dist:portable           # self-contained portable zip (no admin needed)
```

There is **no `package.json`/scripts in `backend/`** and **no Vitest** — UI has
no automated tests. `build_core.ps1` must run before `npm run dist*` or the
packaged app will spawn a missing `.exe`.

## Architecture

```
Serasa email → Outlook (COM) → Python core (script or PyInstaller .exe)
                                   │  stdout: one JSON object per line
                          Electron main.ts (spawn + readline)
                                   │  webContents.send('script:log' | 'script:done')
                          React renderer (useScriptRunner reducer)
                                   │  xlwings (COM) drives Excel
                              Diário .xlsm  +  Cópia pública .xlsm
```

**Process boundary is the contract.** `backend/src/log_emitter.py` defines the
only sanctioned way Python talks to the UI: `emit()` writes exactly one JSON
line (flushed) to stdout; `emit_result()` is the terminal event
(`step: "done"`, `data.result.{status, spreadsheet_path}`). `main.ts` parses
each line; **unparseable lines become `level: warning` events, never crash the
UI**; stderr is surfaced as `level: error`. If you add Python output, route it
through `emit()` — a stray `print()` corrupts the stream.

**The base-spreadsheet cascade is the most fragile invariant.**
`encontrar_arquivo_base_excel()` searches in order: (1) local `planilhas`,
(2) legacy v2.14.1 folder `Documents\Gerencie Carteira\Diário` (note the space),
(3) network share `A:\PUBLICA\GERENCIE CARTEIRA PUBLICA`. Found-elsewhere bases
are copied local first. If all fail it emits `excel.base.needs_user` and exits
with code **4**. `main.ts` (via `useScriptRunner`) catches exit 4, opens a
native file dialog, copies the chosen file local (renaming to
`Gerencie Carteira_{yesterday}.xlsm` if it doesn't match the date pattern), and
**auto-reruns the whole pipeline**. The pattern
`Gerencie Carteira_YYYY_MM_DD.xlsm` is duplicated in `main.ts`
(`BASE_FILENAME_PATTERN`) and the Python core — keep them in sync.

**Ordering invariant (the reason the base-resolution fix exists):** in `main()`, the base must
be resolved (`encontrar_arquivo_base_excel`) *before*
`extrair_dados_dos_anexos()`, because the latter marks emails `UnRead = False`.
If the cascade exited *after* emails were consumed, the post-dialog auto-rerun
would find no unread mail and silently lose a day of data. Do not reorder these
calls. This ordering has **no test** (`test_cascata_base.py` tests the cascade
in isolation) — a regression here passes CI unnoticed.

**Known residual risk (not yet fixed):** emails are still marked read in
`extrair_dados_dos_anexos()` *before* `wb.save()`. If the Excel save fails, the
day's emails are already consumed and the data is lost. A full fix requires
deferring `UnRead = False` until after the save.

## Config & runtime layout

`config/config.ini` holds universal defaults using literal `%USERPROFILE%`
(read with `interpolation=None`; Python expands via `os.path.expandvars`).
`main.ts` also parses this file directly (regex) to build the `shell.openPath`
whitelist and to resolve the local `planilhas` dir for the dialog copy.

On every run the Python core idempotently bootstraps
`%USERPROFILE%\Documents\Gerencie_Carteira\{planilhas,html,logs}`
(`directory_bootstrap.ensure_workspace`). The public share
`A:\PUBLICA\GERENCIE CARTEIRA PUBLICA` is **only verified, never created** — if
offline, the publish step is skipped with a graceful warning.

Excel rules: workbooks **must** stay `.xlsm` (preserve VBA macros); the VLOOKUP
formula in the verification column (`config [Excel] coluna_verificacao`, default
`E`) is copied from the previous row into new rows to avoid `#REF`; `#N/D`
results mean an unregistered CNPJ and produce a "pendências" warning (and launch
`executavel_direciona` if present). The core forces `app.calculate()` before
reading that column to avoid false `#N/D` detection from async recalculation.

## Versionamento (SemVer)

Este repositorio e **unico** (a antiga arvore `versions/main/MAJOR.MINOR/PATCH/`
de repos legados foi descontinuada — historico fica no git e em
`context/timeline.md`). O entrypoint do core e
`backend/src/gerencie_carteira.py` (**nome estavel, sem versao no nome** — nunca
re-versionar o arquivo).

**Fonte unica da verdade:** o campo `version` de `app/package.json`. Nenhum
outro arquivo deve hardcodar a versao. Como ela se propaga:

- **UI:** `App.tsx` busca via IPC `app:version` (`app.getVersion()` no main, que
  le o `package.json`). Nunca escrever a versao literal em `App.tsx`/JSX.
- **Core Python:** recebe `APP_VERSION` no env (injetado pelo Electron em
  `main.ts`); fallback le `app/package.json`; ultimo recurso `"dev"`. Ver
  `_resolver_versao()`. Nunca hardcodar versao no Python nem no nome do arquivo.
- **Artefatos de build:** nomes neutros (sem versao embutida nos scripts).

**Regra de bump (toda alteracao funcional incrementa a versao):**

- **MAJOR** — mudanca incompativel: contrato JSON Lines, formato da planilha,
  protocolo IPC, ou quebra operacional para o usuario.
- **MINOR** — nova funcionalidade retrocompativel (novo passo no pipeline,
  nova acao na UI, novo handler IPC aditivo).
- **PATCH** — correcao de bug ou ajuste interno sem mudanca de contrato.

**Checklist de release (obrigatorio a cada alteracao):**

1. `npm version <major|minor|patch> --no-git-tag-version` em `app/` (edita so o
   `package.json` — a UI, os logs e o Python passam a refletir automaticamente).
2. Registrar a mudanca em `context/timeline.md` (fase atual) e, se mudou
   estado/divida/arquitetura, em `context/metaspec.md` (header: versao + data).
3. Commit; tag git `vMAJOR.MINOR.PATCH`. Rollback = checkout da tag + rebuild.

Nao reintroduzir versao em nomes de arquivo, comentarios de config, README,
JSX ou strings do Python — todos derivam da fonte unica.

## Context docs

`context/metaspec.md`, `context/index.md`, `context/timeline.md`, and
`context/analysis/` are maintained by the `context-*` skills and are the
canonical project memory — consult `metaspec.md` for current state and tech
debt before large changes; update via the `context-update` skill rather than
editing by hand.
