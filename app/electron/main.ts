import { app, BrowserWindow, dialog, ipcMain, shell } from "electron";
import { spawn, type ChildProcess } from "node:child_process";
import { createInterface } from "node:readline";
import path from "node:path";
import fs from "node:fs";

// Quando o app roda de uma unidade de rede (A:\PUBLICA\...), o sandbox do
// processo GPU do Chromium falha em iniciar (`gpu_process_host: error_code=18`
// -> apos 10 tentativas, FATAL "GPU process isn't usable. Goodbye" -> crash
// 0x80000003 no boot). Desligar o sandbox SO do GPU resolve, sem precisar
// abrir mao do sandbox dos demais processos.
app.commandLine.appendSwitch("disable-gpu-sandbox");

const BASE_FILENAME_PATTERN = /^Gerencie Carteira_\d{4}_\d{2}_\d{2}\.xlsm$/;

// --- Globals ---
let mainWindow: BrowserWindow | null = null;
let activeChild: ChildProcess | null = null;
let runStartMs = 0;
let lastSpreadsheetPath: string | null = null;
let allowedPrefixes: string[] = [];
let forceClose = false;

// --- Helpers ---
const isDev = !app.isPackaged;

function resolvePythonInvocation(): { cmd: string; args: string[]; cwd: string } {
  if (isDev) {
    const backendDir = path.resolve(__dirname, "../../backend/src");
    return {
      cmd: "python",
      args: [path.join(backendDir, "gerencie_carteira.py")],
      cwd: backendDir,
    };
  }
  const exePath = path.join(process.resourcesPath, "gerencie_carteira_core.exe");
  return { cmd: exePath, args: [], cwd: path.dirname(exePath) };
}

function resolveConfigPath(): string {
  if (isDev) {
    return path.resolve(__dirname, "../../config/config.ini");
  }
  return path.join(process.resourcesPath, "config", "config.ini");
}

// Em dev (nao empacotado) as pastas de trabalho ficam em <repo>/data — espelha
// o redirecionamento do core Python (aplicar_pastas_dev) para que execucoes de
// teste NAO toquem a pasta de producao na rede.
function resolveDevDataRoot(): string {
  return path.resolve(__dirname, "../../data");
}

function resolveLocalPlanilhasDir(): string {
  if (isDev) return path.join(resolveDevDataRoot(), "planilhas");
  // Le do config; fallback para %USERPROFILE%\Documents\Gerencie_Carteira\planilhas
  try {
    const cfg = fs.readFileSync(resolveConfigPath(), "utf-8");
    const m = cfg.match(/^pasta_diario_excel\s*=\s*(.+)$/m);
    if (m) {
      return m[1].replace(/%USERPROFILE%/gi, process.env.USERPROFILE || "").trim();
    }
  } catch {
    // ignore
  }
  return path.join(process.env.USERPROFILE || "", "Documents", "Gerencie_Carteira", "planilhas");
}

function resolvePublicDir(): string | null {
  if (isDev) return path.join(resolveDevDataRoot(), "publica");
  try {
    const cfg = fs.readFileSync(resolveConfigPath(), "utf-8");
    const m = cfg.match(/^pasta_copia_excel\s*=\s*(.+)$/m);
    if (m) {
      return m[1].replace(/%USERPROFILE%/gi, process.env.USERPROFILE || "").trim();
    }
  } catch {
    // ignore
  }
  return null;
}

// Recuperacao do supervisor: remove artefatos transacionais orfaos
// (*.partial.<ext> = save interrompido; *.bak.<ext> = backup de colisao) de
// execucoes abortadas. Idempotente — seguro rodar no boot e no fechamento.
// O marcador vem ANTES da extensao para preservar .xlsm para o xlwings.
const ORPHAN_RE = /\.(?:partial|bak)\.[^.]+$/i;
function sweepOrphans(): void {
  const dirs = [resolveLocalPlanilhasDir(), resolvePublicDir()].filter(
    (d): d is string => !!d && fs.existsSync(d)
  );
  for (const dir of dirs) {
    let entries: string[];
    try {
      entries = fs.readdirSync(dir);
    } catch {
      continue;
    }
    for (const name of entries) {
      if (ORPHAN_RE.test(name)) {
        try {
          fs.unlinkSync(path.join(dir, name));
        } catch {
          // best-effort
        }
      }
    }
  }
}

// Aguarda o core encerrar apos o cancel cooperativo; mata se estourar o prazo.
function waitChildExit(timeoutMs: number): Promise<void> {
  return new Promise((resolve) => {
    if (!activeChild) return resolve();
    const child = activeChild;
    const timer = setTimeout(() => {
      try {
        child.kill();
      } catch {
        // ignore
      }
      resolve();
    }, timeoutMs);
    child.once("close", () => {
      clearTimeout(timer);
      resolve();
    });
  });
}

function yesterdayDateStamp(): string {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}_${m}_${day}`;
}

function loadAllowedPathPrefixes(): void {
  // Em dev tudo fica sob <repo>/data — a whitelist so precisa dessa raiz.
  if (isDev) {
    allowedPrefixes = [resolveDevDataRoot()];
    return;
  }
  // Whitelist simples para shell.openPath — apenas paths sob workspace do usuario
  // ou sob os diretorios declarados no config.ini
  const userWorkspace = path.join(
    process.env.USERPROFILE || "",
    "Documents",
    "Gerencie_Carteira"
  );
  allowedPrefixes = [userWorkspace];

  try {
    const cfg = fs.readFileSync(resolveConfigPath(), "utf-8");
    for (const line of cfg.split(/\r?\n/)) {
      const m = line.match(/^pasta_(diario_excel|copia_excel|destino_html|logs)\s*=\s*(.+)$/);
      if (m) {
        const expanded = m[2].replace(/%USERPROFILE%/gi, process.env.USERPROFILE || "").trim();
        if (expanded) allowedPrefixes.push(expanded);
      }
    }
  } catch {
    // config ausente em dev — ok, whitelist ja tem o workspace default
  }
}

function isPathAllowed(target: string): boolean {
  const abs = path.resolve(target);
  return allowedPrefixes.some((prefix) => abs.toLowerCase().startsWith(path.resolve(prefix).toLowerCase()));
}

// --- Window ---
function createMainWindow(): void {
  mainWindow = new BrowserWindow({
    width: 900,
    height: 640,
    minWidth: 640,
    minHeight: 480,
    title: "Gerencie Carteira",
    backgroundColor: "#111827",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  mainWindow.setMenuBarVisibility(false);

  // Confirmacao de fechamento durante execucao + rollback transacional.
  mainWindow.on("close", (e) => {
    if (forceClose || !activeChild) return;
    e.preventDefault();
    void handleCloseDuringRun();
  });

  if (isDev && process.env.VITE_DEV_SERVER_URL) {
    mainWindow.loadURL(process.env.VITE_DEV_SERVER_URL);
    mainWindow.webContents.openDevTools({ mode: "detach" });
  } else {
    mainWindow.loadFile(path.join(__dirname, "../dist/index.html"));
  }
}

// --- IPC: run script ---
ipcMain.handle("script:run", async () => {
  if (activeChild) {
    return { ok: false, reason: "already_running" };
  }
  if (!mainWindow) return { ok: false, reason: "no_window" };

  const { cmd, args, cwd } = resolvePythonInvocation();
  runStartMs = Date.now();
  lastSpreadsheetPath = null;

  try {
    activeChild = spawn(cmd, args, {
      cwd,
      stdio: ["pipe", "pipe", "pipe"],
      windowsHide: true,
      env: {
        ...process.env,
        PYTHONIOENCODING: "utf-8",
        PYTHONUNBUFFERED: "1",
        // Fonte unica da versao: app/package.json (via app.getVersion()).
        APP_VERSION: app.getVersion(),
      },
    });
  } catch (err) {
    mainWindow.webContents.send("script:log", {
      level: "error",
      ts: new Date().toISOString(),
      msg: `Falha ao iniciar processo: ${(err as Error).message}`,
      step: "spawn.error",
    });
    return { ok: false, reason: "spawn_failed" };
  }

  const rl = createInterface({ input: activeChild.stdout! });
  rl.on("line", (line) => {
    if (!line.trim() || !mainWindow) return;
    let event: Record<string, unknown>;
    try {
      event = JSON.parse(line);
      if (typeof event !== "object" || event === null) throw new Error("not-object");
    } catch {
      event = {
        level: "warning",
        ts: new Date().toISOString(),
        msg: line,
        step: "stdout.unparsed",
      };
    }

    // Extrai path da planilha do evento terminal
    const data = (event as any).data;
    if (data && data.result && typeof data.result.spreadsheet_path === "string") {
      lastSpreadsheetPath = data.result.spreadsheet_path;
    }

    mainWindow.webContents.send("script:log", event);
  });

  activeChild.stderr!.on("data", (chunk: Buffer) => {
    if (!mainWindow) return;
    const text = chunk.toString("utf-8");
    mainWindow.webContents.send("script:log", {
      level: "error",
      ts: new Date().toISOString(),
      msg: text.trim(),
      step: "stderr",
    });
  });

  activeChild.on("close", (code) => {
    const durationMs = Date.now() - runStartMs;
    if (mainWindow) {
      mainWindow.webContents.send("script:done", {
        exitCode: code ?? -1,
        spreadsheetPath: lastSpreadsheetPath,
        durationMs,
      });
    }
    activeChild = null;
  });

  activeChild.on("error", (err) => {
    if (mainWindow) {
      mainWindow.webContents.send("script:log", {
        level: "error",
        ts: new Date().toISOString(),
        msg: `Erro no processo: ${err.message}`,
        step: "process.error",
      });
    }
  });

  return { ok: true };
});

// --- Escrita no stdin do core (canal de input UI -> Python) ---
function writeToChildStdin(obj: unknown): boolean {
  if (!activeChild || !activeChild.stdin || activeChild.stdin.destroyed) {
    return false;
  }
  try {
    activeChild.stdin.write(JSON.stringify(obj) + "\n");
    return true;
  } catch {
    return false;
  }
}

// Sentinela de cancelamento cooperativo (usado no fechamento em runtime)
function sendCancelToChild(): boolean {
  return writeToChildStdin({ cancel: true });
}

// --- IPC: fornecer input pedido pelo Python (ex: gerentes orfaos) ---
ipcMain.handle("script:provideInput", async (_evt, payload: unknown) => {
  if (!activeChild) return { ok: false, reason: "not_running" };
  const ok = writeToChildStdin(payload);
  return ok ? { ok: true } : { ok: false, reason: "stdin_unavailable" };
});

// --- IPC: versao do app (fonte unica: package.json) ---
ipcMain.handle("app:version", () => app.getVersion());

// --- IPC: cancel ---
ipcMain.handle("script:cancel", async () => {
  if (!activeChild) return { ok: false, reason: "not_running" };
  activeChild.kill();
  return { ok: true };
});

// --- IPC: abrir arquivo (com whitelist) ---
ipcMain.handle("file:open", async (_evt, target: string) => {
  if (typeof target !== "string" || !target) return { ok: false, reason: "invalid_path" };
  if (!isPathAllowed(target)) return { ok: false, reason: "path_not_allowed" };
  if (!fs.existsSync(target)) return { ok: false, reason: "file_not_found" };
  const err = await shell.openPath(target);
  return err ? { ok: false, reason: err } : { ok: true };
});

// --- IPC: dialog para selecionar planilha base manualmente ---
ipcMain.handle("dialog:choose-base", async () => {
  if (!mainWindow) return { ok: false, reason: "no_window" };

  const result = await dialog.showOpenDialog(mainWindow, {
    title: "Selecione a planilha base (.xlsm)",
    buttonLabel: "Usar como base",
    filters: [{ name: "Excel macro-habilitado", extensions: ["xlsm"] }],
    properties: ["openFile"],
  });

  if (result.canceled || result.filePaths.length === 0) {
    return { ok: false, reason: "canceled" };
  }

  const origem = result.filePaths[0];
  const pastaLocal = resolveLocalPlanilhasDir();

  try {
    fs.mkdirSync(pastaLocal, { recursive: true });
    const origName = path.basename(origem);
    // Normaliza nome: se nao bate o pattern, usa data de ontem para agir como base do dia de hoje
    const destName = BASE_FILENAME_PATTERN.test(origName)
      ? origName
      : `Gerencie Carteira_${yesterdayDateStamp()}.xlsm`;
    const destino = path.join(pastaLocal, destName);
    fs.copyFileSync(origem, destino);
    return { ok: true, destino, renamed: destName !== origName, origem };
  } catch (e) {
    return { ok: false, reason: `copy_failed: ${(e as Error).message}` };
  }
});

// Pergunta SIM/NAO ao fechar em runtime; se confirmar, cancela o core de
// forma cooperativa, aguarda/mata, varre artefatos .partial/.bak e fecha.
async function handleCloseDuringRun(): Promise<void> {
  if (!mainWindow) return;
  const { response } = await dialog.showMessageBox(mainWindow, {
    type: "question",
    buttons: ["Sim", "Não"],
    defaultId: 1,
    cancelId: 1,
    noLink: true,
    title: "Gerencie Carteira",
    message: "O programa está em execução. Deseja realmente fechar?",
    detail:
      "Se fechar agora, a execução será descartada e a planilha anterior " +
      "será mantida intacta (como se o programa não tivesse rodado).",
  });

  if (response !== 0) return; // "Não" — permanece em execucao

  sendCancelToChild();
  await waitChildExit(4000);
  activeChild = null;
  sweepOrphans();
  forceClose = true;
  mainWindow?.close();
}

// --- App lifecycle ---
app.whenReady().then(() => {
  loadAllowedPathPrefixes();
  sweepOrphans(); // recuperacao de execucao abortada anterior
  createMainWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createMainWindow();
  });
});

app.on("window-all-closed", () => {
  if (activeChild) activeChild.kill();
  sweepOrphans();
  if (process.platform !== "darwin") app.quit();
});
