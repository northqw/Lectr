const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('lectrDesktop', {
    openMarkdownFile: () => ipcRenderer.invoke('lectr:open-markdown-file'),
    openMarkdownFileByPath: (payload) => ipcRenderer.invoke('lectr:open-markdown-file-by-path', payload),
    pickLinkFile: (payload) => ipcRenderer.invoke('lectr:pick-link-file', payload),
    pickImageFile: (payload) => ipcRenderer.invoke('lectr:pick-image-file', payload),
    saveMarkdownFile: (payload) => ipcRenderer.invoke('lectr:save-markdown-file', payload),
    exportPreviewPdf: (payload) => ipcRenderer.invoke('lectr:export-preview-pdf', payload),
    setZoomFactor: (zoomFactor) => ipcRenderer.invoke('lectr:set-zoom-factor', zoomFactor),
    getAppVersion: () => ipcRenderer.invoke('lectr:get-app-version')
});
