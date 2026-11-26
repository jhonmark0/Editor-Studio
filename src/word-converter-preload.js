const { contextBridge, ipcRenderer } = require("electron");

const tokenArgument = process.argv.find((argument) => argument.startsWith("--word-converter-token="));
const token = tokenArgument ? tokenArgument.slice("--word-converter-token=".length) : "";

function asBytes(value) {
  return value instanceof Uint8Array ? value : new Uint8Array(value);
}

contextBridge.exposeInMainWorld("wordPdfConverter", {
  ready: () => ipcRenderer.send("word-pdf-converter-ready", token),
  onStart: (callback) => {
    const listener = (_event, value) => callback(asBytes(value));
    ipcRenderer.on("word-pdf-converter-start", listener);
    return () => ipcRenderer.removeListener("word-pdf-converter-start", listener);
  },
  complete: (value) => ipcRenderer.invoke("word-pdf-converter-complete", asBytes(value), token),
  fail: () => ipcRenderer.invoke("word-pdf-converter-failed", token),
});
