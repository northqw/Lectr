const path = require('node:path');
const fs = require('node:fs/promises');
const { app, BrowserWindow, shell, ipcMain, dialog } = require('electron');

const devServerUrl = process.env.ELECTRON_START_URL;

function createWindow() {
    const mainWindow = new BrowserWindow({
        width: 1280,
        height: 840,
        minWidth: 980,
        minHeight: 640,
        autoHideMenuBar: true,
        icon: path.join(__dirname, '..', 'public', 'favicon.png'),
        webPreferences: {
            contextIsolation: true,
            nodeIntegration: false,
            sandbox: true,
            preload: path.join(__dirname, 'preload.cjs')
        }
    });

    mainWindow.webContents.session.setPermissionRequestHandler((_webContents, _permission, callback) => {
        callback(false);
    });

    if (devServerUrl) {
        mainWindow.loadURL(devServerUrl);
    } else {
        mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
    }

    mainWindow.webContents.setWindowOpenHandler(({ url }) => {
        shell.openExternal(url);
        return { action: 'deny' };
    });

    mainWindow.webContents.on('will-navigate', (event, url) => {
        const currentUrl = mainWindow.webContents.getURL();
        if (url !== currentUrl) {
            event.preventDefault();
            shell.openExternal(url);
        }
    });

    mainWindow.webContents.on('will-attach-webview', (event) => {
        event.preventDefault();
    });
}

ipcMain.handle('lectr:open-markdown-file', async () => {
    const result = await dialog.showOpenDialog({
        title: 'Open Markdown',
        properties: ['openFile'],
        filters: [
            { name: 'Markdown', extensions: ['md', 'markdown', 'txt'] },
            { name: 'All files', extensions: ['*'] }
        ]
    });

    if (result.canceled || !result.filePaths || result.filePaths.length === 0) {
        return { opened: false };
    }

    const filePath = result.filePaths[0];
    const content = await fs.readFile(filePath, 'utf8');
    return {
        opened: true,
        filePath,
        fileName: path.basename(filePath),
        content
    };
});

ipcMain.handle('lectr:save-markdown-file', async (_event, payload) => {
    const normalizedPayload = payload && typeof payload === 'object' ? payload : {};
    const content = typeof normalizedPayload.content === 'string' ? normalizedPayload.content : '';
    const currentPath = typeof normalizedPayload.filePath === 'string' ? normalizedPayload.filePath : null;
    const forceDialog = normalizedPayload.forceDialog === true;
    const suggestedName = typeof normalizedPayload.suggestedName === 'string' && normalizedPayload.suggestedName.trim()
        ? normalizedPayload.suggestedName.trim()
        : 'document.md';

    if (currentPath && !forceDialog) {
        await fs.writeFile(currentPath, content, 'utf8');
        return { saved: true, filePath: currentPath, fileName: path.basename(currentPath) };
    }

    const defaultPath = currentPath
        ? currentPath
        : path.join(app.getPath('documents'), suggestedName);
    const result = await dialog.showSaveDialog({
        title: 'Save Markdown',
        defaultPath,
        filters: [
            { name: 'Markdown', extensions: ['md', 'markdown'] },
            { name: 'Text', extensions: ['txt'] },
            { name: 'All files', extensions: ['*'] }
        ]
    });

    if (result.canceled || !result.filePath) {
        return { saved: false };
    }

    await fs.writeFile(result.filePath, content, 'utf8');
    return { saved: true, filePath: result.filePath, fileName: path.basename(result.filePath) };
});

ipcMain.handle('lectr:set-zoom-factor', (event, zoomFactor) => {
    const numeric = Number(zoomFactor);
    const safeFactor = Number.isFinite(numeric)
        ? Math.min(1.3, Math.max(0.8, numeric))
        : 1;

    if (!event.sender.isDestroyed()) {
        event.sender.setZoomFactor(safeFactor);
    }

    return { ok: true, zoomFactor: safeFactor };
});

ipcMain.handle('lectr:get-app-version', () => {
    return app.getVersion();
});

app.whenReady().then(() => {
    createWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});
