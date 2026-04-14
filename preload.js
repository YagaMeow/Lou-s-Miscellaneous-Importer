const { contextBridge, ipcRenderer } = require("electron");
contextBridge.exposeInMainWorld("versions", {
  node: () => process.versions.node,
  chrome: () => process.versions.chrome,
  electron: () => process.versions.electron,
  // ping: () => ipcRenderer.invoke("ping"),
});
contextBridge.exposeInMainWorld("electronAPI", {
  openFileDialog: (type) => ipcRenderer.invoke("dialog:openFile", type),
  parseWordFile: (filePath) => ipcRenderer.invoke("word:parse", filePath),
  parseExcelFile: (filePath) => ipcRenderer.invoke("excel:parse", filePath),
  saveFile: () => ipcRenderer.invoke("dialog:saveFile"),
});
