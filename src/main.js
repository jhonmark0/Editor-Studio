const { app, BrowserWindow, dialog, ipcMain, shell } = require("electron");
const { autoUpdater } = require("electron-updater");
const path = require("node:path");
const fs = require("node:fs");
const { execFile } = require("node:child_process");
const { promisify } = require("node:util");
const { processJob } = require("./media");

let mainWindow;
let updateDialogOpen = false;
let updateDownloadRequested = false;
let availableUpdateVersion = "";
const appIcon = path.join(__dirname, "logo-256.png");
const execFileAsync = promisify(execFile);

function sendUpdateStatus(status, details = {}) {
  if (!mainWindow || mainWindow.isDestroyed()) return;
  mainWindow.webContents.send("update-status", { status, ...details });
}

app.setName("Editor Studio");
if (process.platform === "win32") app.setAppUserModelId("EditorStudio");

autoUpdater.autoDownload = false;
autoUpdater.autoInstallOnAppQuit = false;

if (process.platform === "win32") {
  autoUpdater.verifyUpdateCodeSignature = async (_publisherNames, installerPath) => {
    try {
      const escapedPath = installerPath.replace(/'/g, "''");
      const script = [
        "$signature = Get-AuthenticodeSignature -LiteralPath",
        `'${escapedPath}'`,
        "-ErrorAction Stop;",
        "if ($signature.SignerCertificate -and $signature.SignerCertificate.Thumbprint -eq",
        "'FBCEE81BEE1EAEA1CCE523B44E219EA52D7BCF74')",
        "{ Write-Output 'VALID' } else { Write-Output 'INVALID' }",
      ].join(" ");
      const { stdout } = await execFileAsync(
        "powershell.exe",
        ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
        { windowsHide: true, timeout: 20000 },
      );
      return stdout.trim() === "VALID"
        ? null
        : "Update installer is not signed by Editor Studio Local Testing.";
    } catch (error) {
      return `Update signature could not be verified: ${error.message}`;
    }
  };
}

function showAppMessageBox(options) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    return dialog.showMessageBox(mainWindow, options);
  }
  return dialog.showMessageBox(options);
}

autoUpdater.on("update-available", async ({ version }) => {
  if (updateDialogOpen) return;
  updateDialogOpen = true;
  availableUpdateVersion = version;

  try {
    const { response } = await showAppMessageBox({
      type: "info",
      title: "Update available",
      message: `Editor Studio ${version} is available.`,
      detail: "Download the update now, or choose Later to keep using this version.",
      buttons: ["Download", "Later"],
      defaultId: 0,
      cancelId: 1,
      noLink: true,
    });

    if (response === 0) {
      updateDownloadRequested = true;
      sendUpdateStatus("downloading", { version, percent: 0, transferred: 0, total: 0, bytesPerSecond: 0 });
      autoUpdater.downloadUpdate().catch((error) => {
        console.error("Failed to download the update.", error);
      });
    }
  } catch (error) {
    console.error("Failed to show the update prompt.", error);
  } finally {
    updateDialogOpen = false;
  }
});

autoUpdater.on("download-progress", (progress) => {
  sendUpdateStatus("downloading", { ...progress, version: availableUpdateVersion });
});

autoUpdater.on("update-downloaded", async ({ version }) => {
  updateDownloadRequested = false;
  availableUpdateVersion = version;
  sendUpdateStatus("downloaded", { version, percent: 100 });
  if (updateDialogOpen) return;
  updateDialogOpen = true;

  try {
    const { response } = await showAppMessageBox({
      type: "info",
      title: "Update ready",
      message: `Editor Studio ${version} has been downloaded.`,
      detail: "Restart the app to install the update, or choose Later.",
      buttons: ["Restart and install", "Later"],
      defaultId: 0,
      cancelId: 1,
      noLink: true,
    });

    if (response === 0) autoUpdater.quitAndInstall();
  } catch (error) {
    console.error("Failed to show the update-ready prompt.", error);
  } finally {
    updateDialogOpen = false;
  }
});

autoUpdater.on("error", (error) => {
  console.error("Application update error.", error);
  if (updateDownloadRequested) {
    sendUpdateStatus("error", {
      version: availableUpdateVersion,
      message: "The update could not be downloaded. Please try again later.",
    });
  }
  if (!updateDownloadRequested) return;

  updateDownloadRequested = false;
  dialog.showErrorBox(
    "Update download failed",
    "Editor Studio could not download or verify the update. Please try again later.",
  );
});

const fileFilters = {
  image: [
    { name: "Images", extensions: ["png", "jpg", "jpeg", "webp", "tif", "tiff", "bmp", "gif"] },
  ],
  pdf: [{ name: "PDF files", extensions: ["pdf"] }],
  word: [{ name: "Word documents", extensions: ["docx"] }],
  pdfAndImages: [{
    name: "PDFs and images",
    extensions: ["pdf", "png", "jpg", "jpeg", "webp", "tif", "tiff", "bmp", "gif"],
  }],
  video: [{ name: "Videos", extensions: ["mp4", "mov", "m4v", "mkv", "webm", "avi"] }],
};

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1240,
    height: 850,
    minWidth: 980,
    minHeight: 700,
    backgroundColor: "#f4f3ef",
    title: `Editor Studio v${app.getVersion()}`,
    icon: appIcon,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  mainWindow.removeMenu();
  mainWindow.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
  mainWindow.webContents.on("page-title-updated", (event) => event.preventDefault());
  mainWindow.webContents.on("will-navigate", (event) => event.preventDefault());
  mainWindow.loadFile(path.join(__dirname, "index.html"));
}

app.whenReady().then(() => {
  createWindow();
  if (app.isPackaged) {
    setTimeout(() => {
      autoUpdater.checkForUpdates().catch((error) => {
        console.error("Failed to check for application updates.", error);
      });
    }, 5000);
  }
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

ipcMain.handle("choose-files", async (_event, { type, multiple, action }) => {
  const filters = type === "pdf" && action === "combine" ? fileFilters.pdfAndImages : fileFilters[type];
  if (!filters) throw new Error("Unsupported file category.");
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: multiple ? ["openFile", "multiSelections"] : ["openFile"],
    filters,
  });
  return result.canceled ? [] : result.filePaths;
});

ipcMain.handle("choose-output-folder", async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ["openDirectory", "createDirectory"],
    title: "Choose a folder for completed files",
  });
  return result.canceled ? null : result.filePaths[0];
});

ipcMain.handle("get-app-version", () => app.getVersion());

ipcMain.handle("submit-job", async (event, job) => {
  const report = (message, agent) => {
    if (!event.sender.isDestroyed()) {
      event.sender.send("job-activity", { message, agent, time: new Date().toISOString() });
    }
  };
  return processJob(job, report);
});

ipcMain.handle("reveal-file", async (_event, filePath) => {
  if (typeof filePath !== "string" || !path.isAbsolute(filePath) || !fs.existsSync(filePath)) {
    throw new Error("The requested result file could not be found.");
  }
  shell.showItemInFolder(filePath);
});
