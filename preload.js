const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("electronAPI", {
  openFile: () => ipcRenderer.invoke("dialog:openFile"),
  openFolder: () => ipcRenderer.invoke("dialog:openFolder"),
  generateMarklists: (data) => ipcRenderer.invoke("generate-marklists", data)
});
