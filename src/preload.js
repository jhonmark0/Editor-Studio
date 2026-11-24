const { contextBridge, ipcRenderer, webUtils } = require("electron");

let droppedFilesCallback = null;

document.addEventListener("drop", (event) => {
  const droppedOnPicker = event.composedPath().some((target) => target && target.id === "choose-files");
  if (!droppedOnPicker) return;

  event.preventDefault();
  const paths = Array.from(event.dataTransfer.files, (file) => webUtils.getPathForFile(file));
  if (droppedFilesCallback) droppedFilesCallback(paths);
});

contextBridge.exposeInMainWorld("automationOffice", {
  chooseFiles: (options) => ipcRenderer.invoke("choose-files", options),
  getAppVersion: () => ipcRenderer.invoke("get-app-version"),
  onFilesDropped: (callback) => {
    droppedFilesCallback = callback;
    return () => {
      if (droppedFilesCallback === callback) droppedFilesCallback = null;
    };
  },
  chooseOutputFolder: () => ipcRenderer.invoke("choose-output-folder"),
  submitJob: (job) => ipcRenderer.invoke("submit-job", job),
  revealFile: (filePath) => ipcRenderer.invoke("reveal-file", filePath),
  onActivity: (callback) => {
    const listener = (_event, activity) => callback(activity);
    ipcRenderer.on("job-activity", listener);
    return () => ipcRenderer.removeListener("job-activity", listener);
  },
  onUpdateStatus: (callback) => {
    const listener = (_event, status) => callback(status);
    ipcRenderer.on("update-status", listener);
    return () => ipcRenderer.removeListener("update-status", listener);
  },
});
