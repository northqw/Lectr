const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('lectrDesktop', {
    openMarkdownFile: () => ipcRenderer.invoke('lectr:open-markdown-file'),
    saveMarkdownFile: (payload) => ipcRenderer.invoke('lectr:save-markdown-file', payload),
    setZoomFactor: (zoomFactor) => ipcRenderer.invoke('lectr:set-zoom-factor', zoomFactor),
    getAppVersion: () => ipcRenderer.invoke('lectr:get-app-version')
});
