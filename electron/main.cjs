const path = require('node:path');
const fs = require('node:fs/promises');
const { app, BrowserWindow, shell, ipcMain, dialog } = require('electron');

const devServerUrl = process.env.ELECTRON_START_URL;
const supportedTextExtensions = ['.md', '.markdown', '.txt'];

const toDecodedPath = (input) => {
    try {
        return decodeURIComponent(input);
    } catch (_error) {
        return input;
    }
};

const stripQueryAndHash = (rawPath) => {
    const queryIndex = rawPath.indexOf('?');
    const hashIndex = rawPath.indexOf('#');
    const cutIndex = [queryIndex, hashIndex]
        .filter((index) => index >= 0)
        .reduce((min, index) => Math.min(min, index), rawPath.length);
    return rawPath.slice(0, cutIndex);
};

const resolveLinkedFilePath = (linkTarget, sourceFilePath) => {
    const normalizedTarget = typeof linkTarget === 'string' ? linkTarget.trim() : '';
    if (!normalizedTarget || normalizedTarget.startsWith('#')) {
        return null;
    }

    if (/^(?:https?:|mailto:|tel:|data:|blob:)/i.test(normalizedTarget)) {
        return null;
    }

    let target = stripQueryAndHash(normalizedTarget);
    if (!target) {
        return null;
    }

    if (target.startsWith('file://')) {
        try {
            const parsed = new URL(target);
            target = parsed.pathname || '';
        } catch (_error) {
            return null;
        }
    }

    target = toDecodedPath(target);
    if (!target) {
        return null;
    }

    if (path.isAbsolute(target)) {
        return path.normalize(target);
    }

    if (!sourceFilePath || typeof sourceFilePath !== 'string') {
        return null;
    }

    return path.resolve(path.dirname(sourceFilePath), target);
};

const findExistingLinkedFilePath = async (candidatePath) => {
    if (!candidatePath) {
        return null;
    }

    const extension = path.extname(candidatePath).toLowerCase();
    const attempts = extension
        ? [candidatePath]
        : [candidatePath, ...supportedTextExtensions.map((suffix) => `${candidatePath}${suffix}`)];

    for (const attemptPath of attempts) {
        try {
            const stat = await fs.stat(attemptPath);
            if (stat.isFile()) {
                return attemptPath;
            }
        } catch (_error) {
            // continue searching
        }
    }

    return null;
};

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

ipcMain.handle('lectr:open-markdown-file-by-path', async (_event, payload) => {
    const normalizedPayload = payload && typeof payload === 'object' ? payload : {};
    const linkTarget = typeof normalizedPayload.linkTarget === 'string' ? normalizedPayload.linkTarget : '';
    const sourceFilePath = typeof normalizedPayload.sourceFilePath === 'string' ? normalizedPayload.sourceFilePath : '';
    const resolvedPath = resolveLinkedFilePath(linkTarget, sourceFilePath);
    if (!resolvedPath) {
        return { opened: false };
    }

    const existingPath = await findExistingLinkedFilePath(resolvedPath);
    if (!existingPath) {
        return { opened: false };
    }

    const content = await fs.readFile(existingPath, 'utf8');
    return {
        opened: true,
        filePath: existingPath,
        fileName: path.basename(existingPath),
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

ipcMain.handle('lectr:export-preview-pdf', async (_event, payload) => {
    const normalizedPayload = payload && typeof payload === 'object' ? payload : {};
    const html = typeof normalizedPayload.html === 'string' ? normalizedPayload.html : '';
    const lightCss = typeof normalizedPayload.lightCss === 'string' ? normalizedPayload.lightCss : '';
    const suggestedNameRaw = typeof normalizedPayload.suggestedName === 'string' && normalizedPayload.suggestedName.trim()
        ? normalizedPayload.suggestedName.trim()
        : 'markdown-preview.pdf';
    const suggestedName = suggestedNameRaw.toLowerCase().endsWith('.pdf')
        ? suggestedNameRaw
        : `${suggestedNameRaw}.pdf`;

    if (!html.trim()) {
        return { saved: false };
    }

    const result = await dialog.showSaveDialog({
        title: 'Export PDF',
        defaultPath: path.join(app.getPath('documents'), suggestedName),
        filters: [
            { name: 'PDF', extensions: ['pdf'] }
        ]
    });

    if (result.canceled || !result.filePath) {
        return { saved: false };
    }

    let exportWindow = null;
    try {
        exportWindow = new BrowserWindow({
            show: false,
            webPreferences: {
                contextIsolation: true,
                nodeIntegration: false,
                sandbox: true
            }
        });

        const documentHtml = `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <style>
      ${lightCss}

      @page {
        size: A4;
        margin: 10mm;
      }

      html, body {
        margin: 0;
        padding: 0;
        background: #fff;
        color: #24292f;
      }

      #pdf-root {
        box-sizing: border-box;
      }

      .markdown-body {
        background: #fff !important;
        color: #24292f !important;
      }

      .markdown-body > :first-child {
        margin-top: 0 !important;
      }

      .markdown-body > :last-child {
        margin-bottom: 0 !important;
      }
    </style>
  </head>
  <body>
    <article id="pdf-root" class="markdown-body">${html}</article>
  </body>
</html>`;

        await exportWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(documentHtml)}`);

        await exportWindow.webContents.executeJavaScript(
            'document.fonts && document.fonts.ready ? document.fonts.ready.then(() => true) : true',
            true
        ).catch(() => { });

        const pdfBuffer = await exportWindow.webContents.printToPDF({
            pageSize: 'A4',
            printBackground: true,
            preferCSSPageSize: true
        });

        await fs.writeFile(result.filePath, pdfBuffer);

        return {
            saved: true,
            filePath: result.filePath,
            fileName: path.basename(result.filePath)
        };
    } catch (error) {
        // eslint-disable-next-line no-console
        console.error('Failed to export PDF via desktop print pipeline', error);
        return { saved: false };
    } finally {
        if (exportWindow && !exportWindow.isDestroyed()) {
            exportWindow.close();
        }
    }
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
