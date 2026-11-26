const { app, BrowserWindow, ipcMain } = require("electron");
const { randomUUID } = require("node:crypto");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const READY_CHANNEL = "word-pdf-converter-ready";
const START_CHANNEL = "word-pdf-converter-start";
const COMPLETE_CHANNEL = "word-pdf-converter-complete";
const FAILED_CHANNEL = "word-pdf-converter-failed";
const CONVERSION_TIMEOUT_MS = 120000;
const MAX_INPUT_BYTES = 100 * 1024 * 1024;
const MAX_OUTPUT_BYTES = 150 * 1024 * 1024;
const activeConversions = new Map();
let handlersRegistered = false;

class WordPdfConversionError extends Error {
  constructor(userMessage, cause) {
    super(userMessage);
    this.name = "WordPdfConversionError";
    this.userMessage = userMessage;
    if (cause) this.cause = cause;
  }
}

function registerHandlers() {
  if (handlersRegistered) return;
  handlersRegistered = true;

  ipcMain.on(READY_CHANNEL, (event, token) => {
    const pending = activeConversions.get(event.sender.id);
    if (!pending || !pending.matches(event, token) || pending.started || pending.settled) return;
    pending.started = true;
    event.sender.send(START_CHANNEL, pending.bytes);
  });

  ipcMain.handle(COMPLETE_CHANNEL, (event, value, token) => {
    const pending = activeConversions.get(event.sender.id);
    if (!pending || !pending.matches(event, token) || pending.settled) return false;
    const bytes = value instanceof Uint8Array
      ? value
      : value instanceof ArrayBuffer ? new Uint8Array(value) : null;
    if (!bytes || bytes.byteLength < 5 || bytes.byteLength > MAX_OUTPUT_BYTES ||
        String.fromCharCode(...bytes.subarray(0, 5)) !== "%PDF-") {
      pending.fail(new WordPdfConversionError(
        "This Word document could not be converted to PDF. Please try again.",
      ));
      return false;
    }
    pending.finish(null, new Uint8Array(bytes));
    return true;
  });

  ipcMain.handle(FAILED_CHANNEL, (event, token) => {
    const pending = activeConversions.get(event.sender.id);
    if (!pending || !pending.matches(event, token) || pending.settled) return false;
    pending.fail(new WordPdfConversionError(
      "This Word document could not be converted to PDF. Make sure it is a valid .docx and try again.",
    ));
    return true;
  });
}

function convertDocxToPdfInElectron(input) {
  if (!app || !BrowserWindow || !app.isReady()) {
    return Promise.reject(new WordPdfConversionError(
      "The Word-to-PDF converter is unavailable. Close and reopen Editor Studio, then try again.",
    ));
  }

  const bytes = input instanceof Uint8Array
    ? input
    : input instanceof ArrayBuffer ? new Uint8Array(input) : null;
  if (!bytes || !bytes.byteLength) {
    return Promise.reject(new WordPdfConversionError("The Word document is empty."));
  }
  if (bytes.byteLength > MAX_INPUT_BYTES) {
    return Promise.reject(new WordPdfConversionError(
      "This Word document is too large to convert to PDF. Try a smaller document.",
    ));
  }

  registerHandlers();
  const token = randomUUID();
  const converterPath = path.join(__dirname, "word-converter.html");
  const converterUrl = pathToFileURL(converterPath);
  converterUrl.searchParams.set("token", token);
  const partition = `word-pdf-${randomUUID()}`;

  return new Promise((resolve, reject) => {
    let converterWindow;
    let converterSession;
    let webContentsId;
    let deadline;
    let settled = false;

    const cleanup = (error, result) => {
      if (settled) return;
      settled = true;
      clearTimeout(deadline);
      if (webContentsId) activeConversions.delete(webContentsId);
      if (converterWindow && !converterWindow.isDestroyed()) {
        converterWindow.destroy();
      }
      if (converterSession) {
        converterSession.clearStorageData().catch(() => {});
      }
      if (error) reject(error);
      else resolve(result);
    };

    const pending = {
      bytes: new Uint8Array(bytes),
      url: converterUrl.href,
      started: false,
      settled: false,
      matches(event, receivedToken) {
        return !this.settled && receivedToken === token && event.senderFrame?.url === this.url;
      },
      finish(error, result) {
        if (this.settled) return;
        this.settled = true;
        cleanup(error, result);
      },
      fail(error) {
        this.finish(error instanceof WordPdfConversionError
          ? error
          : new WordPdfConversionError(
            "This Word document could not be converted to PDF. Please try again.",
            error,
          ));
      },
    };

    try {
      converterWindow = new BrowserWindow({
        show: false,
        width: 1024,
        height: 768,
        title: "Word to PDF converter",
        webPreferences: {
          preload: path.join(__dirname, "word-converter-preload.js"),
          contextIsolation: true,
          nodeIntegration: false,
          sandbox: true,
          webSecurity: true,
          allowRunningInsecureContent: false,
          backgroundThrottling: false,
          spellcheck: false,
          partition,
          additionalArguments: [`--word-converter-token=${token}`],
        },
      });
      converterSession = converterWindow.webContents.session;
      webContentsId = converterWindow.webContents.id;
      activeConversions.set(webContentsId, pending);
      converterSession.setPermissionRequestHandler((_contents, _permission, callback) => callback(false));
      converterSession.setPermissionCheckHandler(() => false);
      converterWindow.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
      converterWindow.webContents.on("console-message", (details) => {
        if (details.level === "error") console.error("Word-to-PDF converter:", details.message);
      });
      converterWindow.webContents.on("will-navigate", (event) => event.preventDefault());
      converterWindow.webContents.on("will-redirect", (event) => event.preventDefault());
      converterWindow.webContents.on("render-process-gone", () => {
        pending.fail(new WordPdfConversionError(
          "The Word-to-PDF converter stopped unexpectedly. Close other apps and try again.",
        ));
      });
      converterWindow.webContents.on("did-fail-load", (_event, _code, _description, _url, isMainFrame) => {
        if (isMainFrame) {
          pending.fail(new WordPdfConversionError(
            "The Word-to-PDF converter could not start. Close and reopen Editor Studio, then try again.",
          ));
        }
      });
      converterWindow.webContents.on("closed", () => {
        activeConversions.delete(webContentsId);
        pending.fail(new WordPdfConversionError(
          "The Word-to-PDF converter closed before finishing. Please try again.",
        ));
      });

      deadline = setTimeout(() => {
        pending.fail(new WordPdfConversionError(
          "The Word-to-PDF conversion timed out. Try a smaller document or close other apps, then try again.",
        ));
      }, CONVERSION_TIMEOUT_MS);

      converterWindow.loadFile(converterPath, { query: { token } }).catch((error) => {
        pending.fail(new WordPdfConversionError(
          "The Word-to-PDF converter could not start. Close and reopen Editor Studio, then try again.",
          error,
        ));
      });
    } catch (error) {
      pending.fail(new WordPdfConversionError(
        "The Word-to-PDF converter could not start. Close and reopen Editor Studio, then try again.",
        error,
      ));
    }
  });
}

module.exports = { convertDocxToPdfInElectron };
