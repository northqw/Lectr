import * as monaco from 'https://cdn.jsdelivr.net/npm/monaco-editor@0.52.2/+esm';
import { marked } from 'marked';
import DOMPurify from 'dompurify';

const init = () => {
    let scrollBarSync = false;
    let closeOpenMenus = () => { };
    let tabs = [];
    let activeTabId = null;
    let nextTabId = 1;
    let suppressEditorChange = false;
    let currentLanguage = 'en';
    let syncFromEditor = false;
    let syncFromPreview = false;
    let persistTabsTimer = null;
    let pendingMarkdownForRender = '';
    let renderFrameId = null;
    let lastRenderedMarkdown = null;

    const localStorageNamespace = 'com.lectr';
    const localStorageKey = 'last_state';
    const localStorageScrollBarKey = 'scroll_bar_settings';
    const localStorageThemeKey = 'theme_settings';
    const localStorageLanguageKey = 'language_settings';
    const localStorageZoomKey = 'ui_zoom_settings';
    const localStorageTabsStateKey = 'tabs_state';
    const translations = {
        en: {
            file: 'File',
            settings: 'Settings',
            open: 'Open...',
            save: 'Save',
            saveAs: 'Save As...',
            copy: 'Copy',
            copied: 'Copied!',
            exportPdf: 'Export to PDF',
            reset: 'Reset',
            syncScroll: 'Sync scroll',
            darkMode: 'Dark mode',
            language: 'Language',
            scale: 'Scale',
            newTab: 'New tab',
            closeTab: 'Close',
            confirmReset: 'Are you sure you want to reset? Your changes will be lost.',
            confirmCloseUnsavedSave: 'File "{title}" has unsaved changes. Save and close it?',
            confirmCloseUnsavedDiscard: 'Close "{title}" without saving?',
            saveAsPrompt: 'Save as',
            untitledBase: 'Untitled',
            saved: 'Saved',
            format: {
                bold: 'Bold',
                italic: 'Italic',
                strikethrough: 'Strikethrough',
                code: 'Inline code',
                h1: 'Heading 1',
                h2: 'Heading 2',
                ul: 'Unordered list',
                ol: 'Ordered list',
                quote: 'Quote',
                link: 'Link'
            }
        },
        ru: {
            file: 'Файл',
            settings: 'Настройки',
            open: 'Открыть...',
            save: 'Сохранить',
            saveAs: 'Сохранить как...',
            copy: 'Копировать',
            copied: 'Скопировано!',
            exportPdf: 'Экспорт в PDF',
            reset: 'Сбросить',
            syncScroll: 'Синхр. скролл',
            darkMode: 'Тёмная тема',
            language: 'Язык',
            scale: 'Масштаб',
            newTab: 'Новая вкладка',
            closeTab: 'Закрыть',
            confirmReset: 'Сбросить содержимое? Изменения будут потеряны.',
            confirmCloseUnsavedSave: 'В файле "{title}" есть несохраненные изменения. Сохранить и закрыть?',
            confirmCloseUnsavedDiscard: 'Закрыть "{title}" без сохранения?',
            saveAsPrompt: 'Сохранить как',
            untitledBase: 'Без имени',
            saved: 'Сохранено',
            format: {
                bold: 'Жирный',
                italic: 'Курсив',
                strikethrough: 'Зачеркнутый',
                code: 'Код',
                h1: 'Заголовок 1',
                h2: 'Заголовок 2',
                ul: 'Маркированный список',
                ol: 'Нумерованный список',
                quote: 'Цитата',
                link: 'Ссылка'
            }
        }
    };
    // default template
    const defaultInput = `# Lectr Markdown guide

## Headers

# This is a Heading h1
## This is a Heading h2
###### This is a Heading h6

## Emphasis

*This text will be italic*  
_This will also be italic_

**This text will be bold**  
__This will also be bold__

_You **can** combine them_

## Lists

### Unordered

* Item 1
* Item 2
* Item 2a
* Item 2b
    * Item 3a
    * Item 3b

### Ordered

1. Item 1
2. Item 2
3. Item 3
    1. Item 3a
    2. Item 3b

## Images

![This is an alt text.](image/Markdown-mark.svg "This is a sample image.")

## Links

Build your docs with **Lectr** and Markdown.

## Blockquotes

> Markdown is a lightweight markup language with plain-text-formatting syntax, created in 2004 by John Gruber with Aaron Swartz.
>
>> Markdown is often used to format readme files, for writing messages in online discussion forums, and to create rich text using a plain text editor.

## Tables

| Left columns  | Right columns |
| ------------- |:-------------:|
| left foo      | right foo     |
| left bar      | right bar     |
| left baz      | right baz     |

## Blocks of code

${"`"}${"`"}${"`"}
let message = 'Hello world';
alert(message);
${"`"}${"`"}${"`"}

## Inline code

This web site is using ${"`"}markedjs/marked${"`"}.
`;

    const getFileNameFromPath = (filePath) => {
        if (!filePath) {
            return null;
        }
        const normalized = filePath.replace(/\\/g, '/');
        const parts = normalized.split('/');
        return parts[parts.length - 1] || null;
    };

    const t = (key, params = {}) => {
        const dictionary = translations[currentLanguage] || translations.en;
        const fallback = translations.en;
        const source = dictionary[key] !== undefined ? dictionary[key] : fallback[key];
        if (typeof source !== 'string') {
            return key;
        }
        return Object.entries(params).reduce((value, [paramKey, paramValue]) => {
            return value.replace(`{${paramKey}}`, String(paramValue));
        }, source);
    };

    const getUntitledTitle = (index) => {
        const base = t('untitledBase');
        if (index <= 1) {
            return `${base}.md`;
        }
        return `${base}-${index}.md`;
    };

    const applyLocalization = () => {
        const byId = (id) => document.querySelector(`#${id}`);

        const mappings = [
            ['file-menu-label', 'file'],
            ['settings-menu-label', 'settings'],
            ['open-file-button', 'open'],
            ['save-file-button', 'save'],
            ['save-as-file-button', 'saveAs'],
            ['copy-button', 'copy'],
            ['export-button', 'exportPdf'],
            ['reset-button', 'reset'],
            ['sync-scroll-label', 'syncScroll'],
            ['theme-label', 'darkMode'],
            ['language-label', 'language'],
            ['zoom-label', 'scale']
        ];

        mappings.forEach(([id, key]) => {
            const element = byId(id);
            if (element) {
                element.textContent = t(key);
            }
        });

        const newTabButton = byId('new-tab-button');
        if (newTabButton) {
            const label = t('newTab');
            newTabButton.setAttribute('aria-label', label);
            newTabButton.setAttribute('title', label);
        }

        const formatButtons = Array.from(document.querySelectorAll('.format-button'));
        formatButtons.forEach((button) => {
            const formatType = button.getAttribute('data-format');
            if (!formatType) {
                return;
            }
            const formatLabel = translations[currentLanguage]?.format?.[formatType]
                || translations.en.format[formatType]
                || formatType;
            button.setAttribute('title', formatLabel);
            button.setAttribute('aria-label', formatLabel);
        });

        document.documentElement.setAttribute('lang', currentLanguage === 'ru' ? 'ru' : 'en');
        renderTabs();
    };

    const createTab = ({ title, content, filePath = null, dirty = false, lastSavedContent = null }) => ({
        id: `tab-${nextTabId++}`,
        title: title || getUntitledTitle(1),
        content: content || '',
        filePath,
        dirty,
        lastSavedContent: lastSavedContent ?? (dirty ? '' : (content || ''))
    });

    const getActiveTab = () => tabs.find((tab) => tab.id === activeTabId) || null;

    self.MonacoEnvironment = {
        getWorker(_, label) {
            return new Proxy({}, { get: () => () => { } });
        }
    }

    let setupEditor = () => {
        let editor = monaco.editor.create(document.querySelector('#editor'), {
            fontSize: 14,
            language: 'markdown',
            minimap: { enabled: false },
            scrollBeyondLastLine: false,
            automaticLayout: true,
            scrollbar: {
                vertical: 'visible',
                horizontal: 'visible'
            },
            wordWrap: 'on',
            hover: { enabled: false },
            quickSuggestions: false,
            suggestOnTriggerCharacters: false,
            folding: false
        });

        editor.onDidChangeModelContent(() => {
            let value = editor.getValue();
            if (!suppressEditorChange) {
                const activeTab = getActiveTab();
                if (activeTab) {
                    activeTab.content = value;
                    activeTab.dirty = value !== activeTab.lastSavedContent;
                    renderTabs();
                }
                saveLastContent(value);
            }
            scheduleConvert(value);
        });

        editor.onDidScrollChange((e) => {
            if (!scrollBarSync) {
                return;
            }
            if (syncFromPreview) {
                return;
            }

            const scrollTop = e.scrollTop;
            const scrollHeight = e.scrollHeight;
            const height = editor.getLayoutInfo().height;

            const maxScrollTop = scrollHeight - height;
            const scrollRatio = maxScrollTop > 0 ? scrollTop / maxScrollTop : 0;

            let previewElement = document.querySelector('#preview');
            if (!previewElement) {
                return;
            }
            let targetY = (previewElement.scrollHeight - previewElement.clientHeight) * scrollRatio;
            syncFromEditor = true;
            previewElement.scrollTo(0, targetY);
            requestAnimationFrame(() => {
                syncFromEditor = false;
            });
        });

        return editor;
    };

    let setupPreviewScrollSync = (editor) => {
        const previewElement = document.querySelector('#preview');
        if (!previewElement) {
            return;
        }

        previewElement.addEventListener('scroll', () => {
            if (!scrollBarSync) {
                return;
            }
            if (syncFromEditor) {
                return;
            }

            const maxPreviewScroll = previewElement.scrollHeight - previewElement.clientHeight;
            const previewRatio = maxPreviewScroll > 0 ? previewElement.scrollTop / maxPreviewScroll : 0;

            const editorScrollHeight = editor.getScrollHeight();
            const editorHeight = editor.getLayoutInfo().height;
            const maxEditorScroll = editorScrollHeight - editorHeight;
            const nextEditorTop = maxEditorScroll > 0 ? maxEditorScroll * previewRatio : 0;

            syncFromPreview = true;
            editor.setScrollTop(nextEditorTop);
            requestAnimationFrame(() => {
                syncFromPreview = false;
            });
        });
    };

    // Render markdown text as html
    let renderMarkdown = (markdown) => {
        let options = {
            headerIds: false,
            mangle: false
        };
        let html = marked.parse(markdown, options);
        let sanitized = DOMPurify.sanitize(html);
        document.querySelector('#output').innerHTML = sanitized;
    };

    let scheduleConvert = (markdown) => {
        pendingMarkdownForRender = markdown;
        if (renderFrameId !== null) {
            return;
        }

        renderFrameId = window.requestAnimationFrame(() => {
            renderFrameId = null;
            if (pendingMarkdownForRender === lastRenderedMarkdown) {
                return;
            }
            renderMarkdown(pendingMarkdownForRender);
            lastRenderedMarkdown = pendingMarkdownForRender;
        });
    };

    // Reset input text
    let reset = () => {
        const activeTab = getActiveTab();
        if (activeTab && activeTab.dirty) {
            var confirmed = window.confirm(t('confirmReset'));
            if (!confirmed) {
                return;
            }
        }
        if (activeTab) {
            activeTab.content = defaultInput;
            activeTab.dirty = activeTab.content !== activeTab.lastSavedContent;
            renderTabs();
        }
        presetValue(defaultInput);
        document.querySelectorAll('.column').forEach((element) => {
            element.scrollTo({ top: 0 });
        });
    };

    let presetValue = (value) => {
        suppressEditorChange = true;
        editor.setValue(value);
        suppressEditorChange = false;
        editor.revealPosition({ lineNumber: 1, column: 1 });
        editor.focus();
        scheduleConvert(value);
    };

    // ----- sync scroll position -----

    let initScrollBarSync = (settings) => {
        let checkbox = document.querySelector('#sync-scroll-checkbox');
        checkbox.checked = settings;
        scrollBarSync = settings;

        checkbox.addEventListener('change', (event) => {
            let checked = event.currentTarget.checked;
            scrollBarSync = checked;
            saveScrollBarSettings(checked);
        });
    };

    // ----- preview CSS loader (switch github-markdown css) -----
    const PREVIEW_CSS_LIGHT = 'css/github-markdown-light.css?v=1.11.0';
    const PREVIEW_CSS_DARK = 'css/github-markdown-dark_dimmed.css?v=1.11.0';

    let setPreviewCss = (useDark) => {
        const link = document.getElementById('gh-markdown-link');
        if (!link) {
            // fallback: create link element
            const newLink = document.createElement('link');
            newLink.id = 'gh-markdown-link';
            newLink.rel = 'stylesheet';
            newLink.href = useDark ? PREVIEW_CSS_DARK : PREVIEW_CSS_LIGHT;
            document.head.appendChild(newLink);
            return;
        }

        // Only update if href differs to avoid unnecessary reload
        const desired = useDark ? PREVIEW_CSS_DARK : PREVIEW_CSS_LIGHT;
        if (link.getAttribute('href') !== desired) {
            link.setAttribute('href', desired);
        }
    };

    // ----- theme toggle (dark/light) -----
    let setTheme = (enabled) => {
        document.documentElement.setAttribute('data-theme', enabled ? 'dark' : 'light');
    };

    let initThemeToggle = (settings) => {
        let checkbox = document.querySelector('#theme-checkbox');
        if (!checkbox) return;
        checkbox.checked = settings;
        setTheme(settings);

        // set Monaco editor theme to match page theme
        if (monaco && monaco.editor && typeof monaco.editor.setTheme === 'function') {
            monaco.editor.setTheme(settings ? 'vs-dark' : 'vs');
        }
        // set preview css to match theme
        setPreviewCss(settings);

        checkbox.addEventListener('change', (event) => {
            let checked = event.currentTarget.checked;
            setTheme(checked);
            saveThemeSettings(checked);
            setPreviewCss(checked);
            if (monaco && monaco.editor && typeof monaco.editor.setTheme === 'function') {
                monaco.editor.setTheme(checked ? 'vs-dark' : 'vs');
            }
        });
    };

    let enableScrollBarSync = () => {
        scrollBarSync = true;
    };

    let disableScrollBarSync = () => {
        scrollBarSync = false;
    };

    // ----- clipboard utils -----

    let copyToClipboard = (text, successHandler, errorHandler) => {
        if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
            navigator.clipboard.writeText(text).then(
                () => {
                    successHandler();
                },
                () => {
                    errorHandler();
                }
            );
            return;
        }

        const textArea = document.createElement('textarea');
        textArea.value = text;
        textArea.setAttribute('readonly', '');
        textArea.style.position = 'absolute';
        textArea.style.left = '-9999px';
        document.body.appendChild(textArea);
        textArea.select();
        try {
            document.execCommand('copy');
            successHandler();
        } catch (error) {
            errorHandler();
        } finally {
            document.body.removeChild(textArea);
        }
    };

    let notifyCopied = () => {
        let copyButton = document.querySelector('#copy-button');
        if (!copyButton) {
            return;
        }
        const defaultLabel = t('copy');
        copyButton.textContent = t('copied');
        setTimeout(() => {
            copyButton.textContent = defaultLabel;
        }, 1000);
    };

    let notifySaved = () => {
        let saveButton = document.querySelector('#save-file-button');
        if (!saveButton) {
            return;
        }
        const defaultLabel = t('save');
        saveButton.textContent = t('saved');
        setTimeout(() => {
            saveButton.textContent = defaultLabel;
        }, 1100);
    };

    let syncActiveTabFromEditor = (editor) => {
        const activeTab = getActiveTab();
        if (!activeTab || !editor) {
            return null;
        }
        activeTab.content = editor.getValue();
        activeTab.dirty = activeTab.content !== activeTab.lastSavedContent;
        return activeTab;
    };

    let persistSavedTab = (tab, saveResult) => {
        if (!tab || !saveResult || !saveResult.saved) {
            return false;
        }
        tab.filePath = saveResult.filePath || tab.filePath;
        tab.title = saveResult.fileName || getFileNameFromPath(tab.filePath) || tab.title;
        tab.lastSavedContent = tab.content;
        tab.dirty = false;
        renderTabs();
        return true;
    };

    let saveTab = async (tab, { forceDialog = false, showToast = true } = {}) => {
        if (!tab) {
            return { saved: false };
        }

        const saveResult = await saveFile({
            content: tab.content,
            filePath: tab.filePath,
            forceDialog,
            suggestedName: tab.title || 'document.md'
        });

        const saved = persistSavedTab(tab, saveResult);
        if (saved && showToast) {
            notifySaved();
        }

        return saveResult;
    };

    // ----- export preview -----

    let exportLightCssPromise = null;

    let getLightMarkdownCss = () => {
        if (exportLightCssPromise) {
            return exportLightCssPromise;
        }

        exportLightCssPromise = fetch(PREVIEW_CSS_LIGHT)
            .then((response) => {
                if (!response.ok) {
                    throw new Error(`Failed to load export CSS: ${response.status}`);
                }
                return response.text();
            })
            .catch((error) => {
                // eslint-disable-next-line no-console
                console.error('Failed to load light markdown CSS', error);
                return '';
            });

        return exportLightCssPromise;
    };

    let exportPreviewToPdf = () => {
        const previewElement = document.querySelector('#preview-wrapper');
        if (!previewElement) {
            return;
        }

        if (typeof window.html2pdf !== 'function') {
            window.alert('PDF export is not available yet. Please try again in a moment.');
            return;
        }

        getLightMarkdownCss().then((lightCss) => {
            const options = {
                margin: 10,
                filename: 'markdown-preview.pdf',
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: {
                    scale: 2,
                    useCORS: true,
                    onclone: (clonedDoc) => {
                        clonedDoc.documentElement.setAttribute('data-theme', 'light');

                        const markdownLink = clonedDoc.getElementById('gh-markdown-link');
                        if (markdownLink) {
                            markdownLink.setAttribute('href', PREVIEW_CSS_LIGHT);
                        }

                        if (lightCss) {
                            const style = clonedDoc.createElement('style');
                            style.id = 'export-light-css';
                            style.textContent = `${lightCss}
#preview-wrapper, #output, body {
  background: #fff !important;
  color: #24292f !important;
}`;
                            clonedDoc.head.appendChild(style);
                        }

                        const clonedPreview = clonedDoc.getElementById('preview-wrapper');
                        if (clonedPreview) {
                            clonedPreview.style.background = '#fff';
                            clonedPreview.style.color = '#24292f';
                            clonedPreview.style.width = '190mm';
                            clonedPreview.style.maxWidth = '190mm';
                        }

                        const clonedOutput = clonedDoc.getElementById('output');
                        if (clonedOutput) {
                            clonedOutput.style.background = '#fff';
                            clonedOutput.style.color = '#24292f';
                            clonedOutput.style.width = '190mm';
                            clonedOutput.style.maxWidth = '190mm';
                        }
                    }
                },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
            };

            window.html2pdf()
                .set(options)
                .from(previewElement)
                .save()
                .catch((error) => {
                    // eslint-disable-next-line no-console
                    console.error('Failed to export PDF', error);
                });
        });
    };

    // ----- tabs -----

    let renderTabs = () => {
        const tabsList = document.querySelector('#tabs-list');
        if (!tabsList) {
            return;
        }

        tabsList.innerHTML = '';
        tabs.forEach((tab) => {
            const tabElement = document.createElement('div');
            tabElement.className = `tab-item${tab.id === activeTabId ? ' active' : ''}${tab.dirty ? ' dirty' : ''}`;
            tabElement.setAttribute('data-tab-id', tab.id);
            tabElement.setAttribute('title', tab.filePath || tab.title);
            tabElement.setAttribute('draggable', 'true');

            const title = document.createElement('span');
            title.className = 'tab-title';
            title.textContent = tab.title;
            tabElement.appendChild(title);

            const closeButton = document.createElement('button');
            closeButton.type = 'button';
            closeButton.className = 'tab-close';
            closeButton.setAttribute('data-tab-close', tab.id);
            closeButton.setAttribute('aria-label', `${t('closeTab')} ${tab.title}`);
            closeButton.textContent = 'x';
            closeButton.setAttribute('draggable', 'false');
            tabElement.appendChild(closeButton);

            tabsList.appendChild(tabElement);
        });

        schedulePersistTabsState();
    };

    let activateTab = (tabId) => {
        const tab = tabs.find((item) => item.id === tabId);
        if (!tab) {
            return;
        }
        activeTabId = tabId;
        presetValue(tab.content);
        saveLastContent(tab.content);
        renderTabs();
    };

    let createUntitledTab = (content = '') => {
        const untitledCount = tabs.filter((tab) => tab.filePath === null).length + 1;
        const title = getUntitledTitle(untitledCount);
        const tab = createTab({
            title,
            content,
            filePath: null,
            dirty: false,
            lastSavedContent: content
        });
        tabs.push(tab);
        renderTabs();
        activateTab(tab.id);
        return tab;
    };

    let closeTab = async (tabId, editor) => {
        const tab = tabs.find((item) => item.id === tabId);
        if (!tab) {
            return;
        }

        if (tab.id === activeTabId && editor) {
            tab.content = editor.getValue();
            tab.dirty = tab.content !== tab.lastSavedContent;
        }

        if (tab.dirty) {
            const shouldSaveAndClose = window.confirm(t('confirmCloseUnsavedSave', { title: tab.title }));
            if (shouldSaveAndClose) {
                const saveResult = await saveTab(tab, {
                    forceDialog: !tab.filePath,
                    showToast: true
                });
                if (!saveResult || !saveResult.saved) {
                    return;
                }
            } else {
                const discardConfirmed = window.confirm(t('confirmCloseUnsavedDiscard', { title: tab.title }));
                if (!discardConfirmed) {
                    return;
                }
            }
        }

        const closingIndex = tabs.findIndex((item) => item.id === tabId);
        tabs.splice(closingIndex, 1);

        if (tabs.length === 0) {
            createUntitledTab(defaultInput);
            return;
        }

        const nextIndex = Math.max(0, closingIndex - 1);
        const nextTab = tabs[nextIndex];
        if (activeTabId === tabId && nextTab) {
            activateTab(nextTab.id);
        } else {
            renderTabs();
        }
    };

    let setupTabsUi = (editor) => {
        const tabsList = document.querySelector('#tabs-list');
        let dragTabId = null;
        if (tabsList) {
            tabsList.addEventListener('click', async (event) => {
                const closeTarget = event.target.closest('[data-tab-close]');
                if (closeTarget) {
                    event.preventDefault();
                    event.stopPropagation();
                    const tabId = closeTarget.getAttribute('data-tab-close');
                    if (tabId) {
                        await closeTab(tabId, editor);
                    }
                    return;
                }

                const tabTarget = event.target.closest('[data-tab-id]');
                if (tabTarget) {
                    const tabId = tabTarget.getAttribute('data-tab-id');
                    if (tabId) {
                        activateTab(tabId);
                    }
                }
            });

            tabsList.addEventListener('dragstart', (event) => {
                const tabTarget = event.target.closest('[data-tab-id]');
                if (!tabTarget) {
                    return;
                }
                dragTabId = tabTarget.getAttribute('data-tab-id');
                tabTarget.classList.add('dragging');
                if (event.dataTransfer) {
                    event.dataTransfer.effectAllowed = 'move';
                    event.dataTransfer.setData('text/plain', dragTabId || '');
                }
            });

            tabsList.addEventListener('dragover', (event) => {
                const dropTarget = event.target.closest('[data-tab-id]');
                if (!dropTarget || !dragTabId) {
                    return;
                }
                event.preventDefault();
            });

            tabsList.addEventListener('drop', (event) => {
                const dropTarget = event.target.closest('[data-tab-id]');
                if (!dropTarget || !dragTabId) {
                    return;
                }
                event.preventDefault();
                const targetTabId = dropTarget.getAttribute('data-tab-id');
                if (!targetTabId || targetTabId === dragTabId) {
                    return;
                }

                const dragIndex = tabs.findIndex((tab) => tab.id === dragTabId);
                const targetIndex = tabs.findIndex((tab) => tab.id === targetTabId);
                if (dragIndex === -1 || targetIndex === -1) {
                    return;
                }

                const targetRect = dropTarget.getBoundingClientRect();
                const insertAfter = event.clientX > targetRect.left + targetRect.width / 2;
                const [dragged] = tabs.splice(dragIndex, 1);
                let insertIndex = tabs.findIndex((tab) => tab.id === targetTabId);
                if (insertAfter) {
                    insertIndex += 1;
                }
                tabs.splice(insertIndex, 0, dragged);
                renderTabs();
            });

            tabsList.addEventListener('dragend', () => {
                dragTabId = null;
                tabsList.querySelectorAll('.tab-item.dragging').forEach((element) => {
                    element.classList.remove('dragging');
                });
            });
        }

        const newTabButton = document.querySelector('#new-tab-button');
        if (newTabButton) {
            newTabButton.addEventListener('click', () => {
                createUntitledTab('');
            });
        }
    };

    // ----- setup -----

    let setupHeaderMenus = () => {
        const menus = Array.from(document.querySelectorAll('.menu'));
        if (menus.length === 0) {
            return;
        }

        closeOpenMenus = () => {
            menus.forEach((menuElement) => {
                menuElement.classList.remove('open');
                const trigger = menuElement.querySelector('.menu-trigger');
                if (trigger) {
                    trigger.setAttribute('aria-expanded', 'false');
                }
            });
        };

        const toggleMenu = (menuElement) => {
            const trigger = menuElement.querySelector('.menu-trigger');
            const shouldOpen = !menuElement.classList.contains('open');
            closeOpenMenus();
            if (shouldOpen) {
                menuElement.classList.add('open');
                if (trigger) {
                    trigger.setAttribute('aria-expanded', 'true');
                }
            }
        };

        menus.forEach((menuElement) => {
            const trigger = menuElement.querySelector('.menu-trigger');
            if (!trigger) {
                return;
            }

            trigger.addEventListener('click', (event) => {
                event.preventDefault();
                toggleMenu(menuElement);
            });
        });

        document.addEventListener('pointerdown', (event) => {
            const target = event.target;
            const isInsideMenu = target instanceof Element && !!target.closest('.menu');
            if (!isInsideMenu) {
                closeOpenMenus();
            }
        });

        document.addEventListener('keydown', (event) => {
            if (event.key === 'Escape') {
                closeOpenMenus();
            }
        });
    };

    let setupResetButton = () => {
        const resetButton = document.querySelector('#reset-button');
        if (!resetButton) {
            return;
        }
        resetButton.addEventListener('click', (event) => {
            event.preventDefault();
            reset();
            closeOpenMenus();
        });
    };

    let setupCopyButton = (editor) => {
        const copyButton = document.querySelector('#copy-button');
        if (!copyButton) {
            return;
        }
        copyButton.addEventListener('click', (event) => {
            event.preventDefault();
            let value = editor.getValue();
            copyToClipboard(value, () => {
                notifyCopied();
            },
                () => {
                    // nothing to do
                });
            closeOpenMenus();
        });
    };

    let setupExportButton = () => {
        const exportButton = document.querySelector('#export-button');
        if (!exportButton) {
            return;
        }
        exportButton.addEventListener('click', (event) => {
            event.preventDefault();
            exportPreviewToPdf();
            closeOpenMenus();
        });
    };

    let fallbackOpenFile = () => new Promise((resolve) => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.md,.markdown,.txt,text/markdown,text/plain';
        input.addEventListener('change', async () => {
            const file = input.files && input.files[0];
            if (!file) {
                resolve({ opened: false });
                return;
            }
            const content = await file.text();
            resolve({
                opened: true,
                filePath: null,
                fileName: file.name,
                content
            });
        }, { once: true });
        input.click();
    });

    let openFile = async () => {
        if (window.lectrDesktop && typeof window.lectrDesktop.openMarkdownFile === 'function') {
            try {
                const result = await window.lectrDesktop.openMarkdownFile();
                if (result && result.opened) {
                    return result;
                }
            } catch (error) {
                // eslint-disable-next-line no-console
                console.error('Failed to open file via desktop bridge', error);
            }
        }
        return fallbackOpenFile();
    };

    let fallbackSaveFile = (content, suggestedName = 'document.md') => {
        const fileName = window.prompt(t('saveAsPrompt'), suggestedName);
        if (!fileName) {
            return { saved: false };
        }

        const blob = new Blob([content], { type: 'text/markdown;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        return { saved: true, filePath: null, fileName };
    };

    let saveFile = async ({ content, filePath, forceDialog = false, suggestedName = 'document.md' }) => {
        if (window.lectrDesktop && typeof window.lectrDesktop.saveMarkdownFile === 'function') {
            try {
                const result = await window.lectrDesktop.saveMarkdownFile({
                    content,
                    filePath,
                    forceDialog,
                    suggestedName
                });
                if (result) {
                    return result;
                }
            } catch (error) {
                // eslint-disable-next-line no-console
                console.error('Failed to save via desktop bridge', error);
            }
        }

        return fallbackSaveFile(content, suggestedName);
    };

    let setupOpenButton = () => {
        const openButton = document.querySelector('#open-file-button');
        if (!openButton) {
            return;
        }

        openButton.addEventListener('click', async (event) => {
            event.preventDefault();
            const result = await openFile();
            if (!result || !result.opened) {
                closeOpenMenus();
                return;
            }

            const title = result.fileName || getFileNameFromPath(result.filePath) || 'Opened.md';
            const content = typeof result.content === 'string' ? result.content : '';
            const existingByPath = result.filePath
                ? tabs.find((tab) => tab.filePath === result.filePath)
                : null;

            if (existingByPath) {
                existingByPath.content = content;
                existingByPath.lastSavedContent = content;
                existingByPath.dirty = false;
                existingByPath.title = title;
                activateTab(existingByPath.id);
            } else {
                const tab = createTab({
                    title,
                    content,
                    filePath: result.filePath || null,
                    dirty: false,
                    lastSavedContent: content
                });
                tabs.push(tab);
                renderTabs();
                activateTab(tab.id);
            }

            closeOpenMenus();
        });
    };

    let setupSaveButton = (editor) => {
        const saveButton = document.querySelector('#save-file-button');
        if (!saveButton) {
            return;
        }

        saveButton.addEventListener('click', async (event) => {
            event.preventDefault();
            const activeTab = getActiveTab();
            if (!activeTab) {
                closeOpenMenus();
                return;
            }
            syncActiveTabFromEditor(editor);
            await saveTab(activeTab, { forceDialog: false, showToast: true });
            closeOpenMenus();
        });
    };

    let setupSaveAsButton = (editor) => {
        const saveAsButton = document.querySelector('#save-as-file-button');
        if (!saveAsButton) {
            return;
        }

        saveAsButton.addEventListener('click', async (event) => {
            event.preventDefault();
            const activeTab = getActiveTab();
            if (!activeTab) {
                closeOpenMenus();
                return;
            }
            syncActiveTabFromEditor(editor);
            await saveTab(activeTab, { forceDialog: true, showToast: true });
            closeOpenMenus();
        });
    };

    let hasExactDelimiterAround = (fullText, startOffset, endOffset, prefix, suffix) => {
        if (startOffset < prefix.length || endOffset + suffix.length > fullText.length) {
            return false;
        }

        if (fullText.slice(startOffset - prefix.length, startOffset) !== prefix) {
            return false;
        }

        if (fullText.slice(endOffset, endOffset + suffix.length) !== suffix) {
            return false;
        }

        const sameRepeatedDelimiter = prefix === suffix
            && prefix.length > 0
            && prefix.split('').every((char) => char === prefix[0]);

        if (!sameRepeatedDelimiter) {
            return true;
        }

        const delimiter = prefix[0];
        let leftRun = 0;
        for (let i = startOffset - 1; i >= 0 && fullText[i] === delimiter; i -= 1) {
            leftRun += 1;
        }

        let rightRun = 0;
        for (let i = endOffset; i < fullText.length && fullText[i] === delimiter; i += 1) {
            rightRun += 1;
        }

        if (leftRun < prefix.length || rightRun < suffix.length) {
            return false;
        }

        // Keep bold when italic is toggled on bold text: **text** -> ***text***
        if (prefix.length === 1 && leftRun === 2 && rightRun === 2) {
            return false;
        }

        return true;
    };

    let hasExactDelimiterInside = (selectedText, prefix, suffix) => {
        if (!selectedText.startsWith(prefix) || !selectedText.endsWith(suffix)) {
            return false;
        }

        if (selectedText.length < prefix.length + suffix.length) {
            return false;
        }

        const sameRepeatedDelimiter = prefix === suffix
            && prefix.length > 0
            && prefix.split('').every((char) => char === prefix[0]);

        if (!sameRepeatedDelimiter) {
            return true;
        }

        const delimiter = prefix[0];
        const innerStart = selectedText[prefix.length];
        const innerEnd = selectedText[selectedText.length - suffix.length - 1];
        if (innerStart === delimiter || innerEnd === delimiter) {
            return false;
        }

        return true;
    };

    let applyWrapFormat = (editor, prefix, suffix, placeholder) => {
        const model = editor.getModel();
        const selection = editor.getSelection();
        if (!model || !selection) {
            return;
        }

        const fullText = model.getValue();
        const selectedText = model.getValueInRange(selection);
        const hasSelection = selectedText.length > 0;
        const startOffset = model.getOffsetAt({
            lineNumber: selection.startLineNumber,
            column: selection.startColumn
        });
        const endOffset = model.getOffsetAt({
            lineNumber: selection.endLineNumber,
            column: selection.endColumn
        });

        const canUnwrapAround = hasSelection
            && hasExactDelimiterAround(fullText, startOffset, endOffset, prefix, suffix);

        const canUnwrapInside = hasSelection
            && hasExactDelimiterInside(selectedText, prefix, suffix);

        if (canUnwrapAround) {
            const unwrapRange = new monaco.Range(
                model.getPositionAt(startOffset - prefix.length).lineNumber,
                model.getPositionAt(startOffset - prefix.length).column,
                model.getPositionAt(endOffset + suffix.length).lineNumber,
                model.getPositionAt(endOffset + suffix.length).column
            );
            editor.executeEdits('format-toolbar', [{
                range: unwrapRange,
                text: selectedText,
                forceMoveMarkers: true
            }]);

            const selectionStart = model.getPositionAt(startOffset - prefix.length);
            const selectionEnd = model.getPositionAt(startOffset - prefix.length + selectedText.length);
            editor.setSelection(new monaco.Selection(
                selectionStart.lineNumber,
                selectionStart.column,
                selectionEnd.lineNumber,
                selectionEnd.column
            ));
            editor.focus();
            return;
        }

        if (canUnwrapInside) {
            const unwrapped = selectedText.slice(prefix.length, selectedText.length - suffix.length);
            editor.executeEdits('format-toolbar', [{
                range: selection,
                text: unwrapped,
                forceMoveMarkers: true
            }]);

            const selectionStart = model.getPositionAt(startOffset);
            const selectionEnd = model.getPositionAt(startOffset + unwrapped.length);
            editor.setSelection(new monaco.Selection(
                selectionStart.lineNumber,
                selectionStart.column,
                selectionEnd.lineNumber,
                selectionEnd.column
            ));
            editor.focus();
            return;
        }

        const textToInsert = hasSelection ? selectedText : placeholder;
        const insertedText = `${prefix}${textToInsert}${suffix}`;

        editor.executeEdits('format-toolbar', [{
            range: selection,
            text: insertedText,
            forceMoveMarkers: true
        }]);

        const selectionStart = model.getPositionAt(startOffset + prefix.length);
        const selectionEnd = model.getPositionAt(startOffset + prefix.length + textToInsert.length);
        editor.setSelection(new monaco.Selection(
            selectionStart.lineNumber,
            selectionStart.column,
            selectionEnd.lineNumber,
            selectionEnd.column
        ));
        editor.focus();
    };

    let applyLineTransform = (editor, transformer) => {
        const model = editor.getModel();
        const selection = editor.getSelection();
        if (!model || !selection) {
            return;
        }

        let startLine = selection.startLineNumber;
        let endLine = selection.endLineNumber;
        if (endLine > startLine && selection.endColumn === 1) {
            endLine -= 1;
        }

        const range = new monaco.Range(
            startLine,
            1,
            endLine,
            model.getLineMaxColumn(endLine)
        );

        const lines = [];
        for (let line = startLine; line <= endLine; line += 1) {
            lines.push(model.getLineContent(line));
        }

        const transformedLines = transformer(lines);
        const prefixed = transformedLines.join('\n');

        const startOffset = model.getOffsetAt({ lineNumber: startLine, column: 1 });
        editor.executeEdits('format-toolbar', [{
            range,
            text: prefixed,
            forceMoveMarkers: true
        }]);

        const rangeStart = model.getPositionAt(startOffset);
        const rangeEnd = model.getPositionAt(startOffset + prefixed.length);
        editor.setSelection(new monaco.Selection(
            rangeStart.lineNumber,
            rangeStart.column,
            rangeEnd.lineNumber,
            rangeEnd.column
        ));
        editor.focus();
    };

    let stripHeadingPrefix = (lineText) => lineText.replace(/^(\s{0,3})#{1,6}\s+/, '$1');

    let applyHeadingFormat = (editor, level) => {
        const prefix = `${'#'.repeat(level)} `;
        const targetRegex = new RegExp(`^(\\s{0,3})#{${level}}\\s+`);
        applyLineTransform(editor, (lines) => {
            const isAlreadyTarget = lines.every((lineText) => targetRegex.test(lineText));
            if (isAlreadyTarget) {
                return lines.map((lineText) => stripHeadingPrefix(lineText));
            }
            return lines.map((lineText) => `${prefix}${stripHeadingPrefix(lineText)}`);
        });
    };

    let stripListPrefix = (lineText) => lineText.replace(/^(\s*)(?:[-*+]|\d+\.)\s+/, '$1');

    let applyUnorderedListFormat = (editor) => {
        const unorderedRegex = /^\s*[-*+]\s+/;
        applyLineTransform(editor, (lines) => {
            const isAlreadyUnordered = lines.every((lineText) => unorderedRegex.test(lineText));
            if (isAlreadyUnordered) {
                return lines.map((lineText) => stripListPrefix(lineText));
            }
            return lines.map((lineText) => `- ${stripListPrefix(lineText)}`);
        });
    };

    let applyOrderedListFormat = (editor) => {
        const orderedRegex = /^\s*\d+\.\s+/;
        applyLineTransform(editor, (lines) => {
            const isAlreadyOrdered = lines.every((lineText) => orderedRegex.test(lineText));
            if (isAlreadyOrdered) {
                return lines.map((lineText) => stripListPrefix(lineText));
            }
            return lines.map((lineText, index) => `${index + 1}. ${stripListPrefix(lineText)}`);
        });
    };

    let applyQuoteFormat = (editor) => {
        const quoteRegex = /^(\s{0,3})>\s?/;
        applyLineTransform(editor, (lines) => {
            const isAlreadyQuote = lines.every((lineText) => quoteRegex.test(lineText));
            if (isAlreadyQuote) {
                return lines.map((lineText) => lineText.replace(quoteRegex, '$1'));
            }
            return lines.map((lineText) => `> ${lineText.replace(quoteRegex, '$1')}`);
        });
    };

    let applyLinkFormat = (editor) => {
        const model = editor.getModel();
        const selection = editor.getSelection();
        if (!model || !selection) {
            return;
        }

        const selectedText = model.getValueInRange(selection);
        const hasSelection = selectedText.length > 0;
        const label = hasSelection ? selectedText : 'link text';
        const url = 'https://example.com';
        const insertedText = `[${label}](${url})`;
        const startOffset = model.getOffsetAt({
            lineNumber: selection.startLineNumber,
            column: selection.startColumn
        });

        editor.executeEdits('format-toolbar', [{
            range: selection,
            text: insertedText,
            forceMoveMarkers: true
        }]);

        const urlStartOffset = startOffset + label.length + 3;
        const urlEndOffset = urlStartOffset + url.length;
        const selectionStart = model.getPositionAt(urlStartOffset);
        const selectionEnd = model.getPositionAt(urlEndOffset);
        editor.setSelection(new monaco.Selection(
            selectionStart.lineNumber,
            selectionStart.column,
            selectionEnd.lineNumber,
            selectionEnd.column
        ));
        editor.focus();
    };

    let setupFormatToolbar = (editor) => {
        const toolbar = document.querySelector('#format-toolbar');
        if (!toolbar) {
            return;
        }

        const handlers = {
            bold: () => applyWrapFormat(editor, '**', '**', 'bold text'),
            italic: () => applyWrapFormat(editor, '*', '*', 'italic text'),
            strikethrough: () => applyWrapFormat(editor, '~~', '~~', 'strikethrough'),
            code: () => applyWrapFormat(editor, '`', '`', 'inline code'),
            h1: () => applyHeadingFormat(editor, 1),
            h2: () => applyHeadingFormat(editor, 2),
            ul: () => applyUnorderedListFormat(editor),
            ol: () => applyOrderedListFormat(editor),
            quote: () => applyQuoteFormat(editor),
            link: () => applyLinkFormat(editor)
        };

        const withUndoStop = (handler) => {
            editor.pushUndoStop();
            handler();
            editor.pushUndoStop();
        };

        toolbar.addEventListener('click', (event) => {
            const button = event.target.closest('.format-button');
            if (!button) {
                return;
            }
            event.preventDefault();
            const formatType = button.getAttribute('data-format');
            if (!formatType || !handlers[formatType]) {
                return;
            }
            withUndoStop(handlers[formatType]);
        });

        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyB, () => withUndoStop(handlers.bold));
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyI, () => withUndoStop(handlers.italic));
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyK, () => withUndoStop(handlers.link));
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.KeyX, () => withUndoStop(handlers.strikethrough));
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.KeyC, () => withUndoStop(handlers.code));
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.Digit1, () => withUndoStop(handlers.h1));
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.Digit2, () => withUndoStop(handlers.h2));
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.Digit8, () => withUndoStop(handlers.ul));
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.Digit7, () => withUndoStop(handlers.ol));
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.Period, () => withUndoStop(handlers.quote));
    };

    // ----- local state -----

    let loadLastContent = () => {
        try {
            const raw = localStorage.getItem(`${localStorageNamespace}_${localStorageKey}`);
            return raw ? JSON.parse(raw) : null;
        } catch (error) {
            return null;
        }
    };

    let saveLastContent = (content) => {
        try {
            localStorage.setItem(`${localStorageNamespace}_${localStorageKey}`, JSON.stringify(content));
        } catch (error) {
            // ignore storage errors
        }
    };

    let loadScrollBarSettings = () => {
        try {
            const raw = localStorage.getItem(`${localStorageNamespace}_${localStorageScrollBarKey}`);
            return raw ? JSON.parse(raw) : null;
        } catch (error) {
            return null;
        }
    };

    let loadThemeSettings = () => {
        let last = null;
        try {
            const raw = localStorage.getItem(`${localStorageNamespace}_${localStorageThemeKey}`);
            last = raw ? JSON.parse(raw) : null;
        } catch (error) {
            last = null;
        }
        if (last === null || last === undefined) {
            try {
                // fallback to raw localStorage boot key used by inline script
                const raw = localStorage.getItem('com.lectr_theme');
                if (raw === 'dark') return true;
                if (raw === 'light') return false;
            } catch (e) {
                // ignore
            }
        }
        return last;
    };

    let saveScrollBarSettings = (settings) => {
        try {
            localStorage.setItem(`${localStorageNamespace}_${localStorageScrollBarKey}`, JSON.stringify(settings));
        } catch (error) {
            // ignore storage errors
        }
    };

    let saveThemeSettings = (settings) => {
        try {
            localStorage.setItem(`${localStorageNamespace}_${localStorageThemeKey}`, JSON.stringify(settings));
        } catch (error) {
            // ignore storage errors
        }
        try {
            localStorage.setItem('com.lectr_theme', settings ? 'dark' : 'light');
        } catch (e) {
            // ignore storage errors
        }
    };

    let loadLanguageSettings = () => {
        try {
            const raw = localStorage.getItem(`${localStorageNamespace}_${localStorageLanguageKey}`);
            return raw ? JSON.parse(raw) : null;
        } catch (error) {
            return null;
        }
    };

    let saveLanguageSettings = (language) => {
        try {
            localStorage.setItem(`${localStorageNamespace}_${localStorageLanguageKey}`, JSON.stringify(language));
        } catch (error) {
            // ignore storage errors
        }
    };

    let loadZoomSettings = () => {
        try {
            const raw = localStorage.getItem(`${localStorageNamespace}_${localStorageZoomKey}`);
            return raw ? JSON.parse(raw) : null;
        } catch (error) {
            return null;
        }
    };

    let saveZoomSettings = (zoomValue) => {
        try {
            localStorage.setItem(`${localStorageNamespace}_${localStorageZoomKey}`, JSON.stringify(zoomValue));
        } catch (error) {
            // ignore storage errors
        }
    };

    let loadTabsState = () => {
        try {
            const raw = localStorage.getItem(`${localStorageNamespace}_${localStorageTabsStateKey}`);
            if (!raw) {
                return null;
            }

            const parsed = JSON.parse(raw);
            if (!parsed || !Array.isArray(parsed.tabs) || parsed.tabs.length === 0) {
                return null;
            }

            const normalizedTabs = parsed.tabs.map((tab, index) => {
                const content = typeof tab.content === 'string' ? tab.content : '';
                const lastSavedContent = typeof tab.lastSavedContent === 'string' ? tab.lastSavedContent : content;
                const defaultTitle = getUntitledTitle(index + 1);
                const title = typeof tab.title === 'string' && tab.title.trim() ? tab.title.trim() : defaultTitle;
                const filePath = typeof tab.filePath === 'string' && tab.filePath.trim() ? tab.filePath : null;

                return {
                    id: typeof tab.id === 'string' && tab.id.trim() ? tab.id : null,
                    title,
                    filePath,
                    content,
                    lastSavedContent,
                    dirty: tab.dirty === true || content !== lastSavedContent
                };
            });

            const activeTabId = typeof parsed.activeTabId === 'string' ? parsed.activeTabId : null;

            return {
                tabs: normalizedTabs,
                activeTabId
            };
        } catch (error) {
            return null;
        }
    };

    let saveTabsState = () => {
        try {
            const snapshot = {
                version: 1,
                activeTabId,
                tabs: tabs.map((tab) => ({
                    id: tab.id,
                    title: tab.title,
                    filePath: tab.filePath,
                    content: tab.content,
                    lastSavedContent: tab.lastSavedContent,
                    dirty: tab.dirty
                }))
            };
            localStorage.setItem(`${localStorageNamespace}_${localStorageTabsStateKey}`, JSON.stringify(snapshot));
        } catch (error) {
            // ignore storage errors
        }
    };

    let schedulePersistTabsState = () => {
        if (persistTabsTimer !== null) {
            window.clearTimeout(persistTabsTimer);
        }
        persistTabsTimer = window.setTimeout(() => {
            persistTabsTimer = null;
            saveTabsState();
        }, 120);
    };

    let applyUiZoom = (zoomValue) => {
        const numeric = Number.parseInt(zoomValue, 10);
        const safeZoom = Number.isFinite(numeric) ? Math.min(130, Math.max(80, numeric)) : 100;
        const zoomFactor = safeZoom / 100;
        if (window.lectrDesktop && typeof window.lectrDesktop.setZoomFactor === 'function') {
            window.lectrDesktop.setZoomFactor(zoomFactor).catch((error) => {
                // eslint-disable-next-line no-console
                console.error('Failed to apply desktop zoom', error);
            });
            return;
        }
        document.documentElement.style.fontSize = `${safeZoom}%`;
    };

    let setupAppVersion = async () => {
        const versionElement = document.querySelector('#app-version');
        if (!versionElement) {
            return;
        }

        const fallback = versionElement.textContent ? versionElement.textContent.trim() : 'v1.0.0';
        versionElement.textContent = fallback.startsWith('v') ? fallback : `v${fallback}`;

        if (!window.lectrDesktop || typeof window.lectrDesktop.getAppVersion !== 'function') {
            return;
        }

        try {
            const version = await window.lectrDesktop.getAppVersion();
            if (typeof version === 'string' && version.trim()) {
                versionElement.textContent = `v${version.trim()}`;
            }
        } catch (error) {
            // eslint-disable-next-line no-console
            console.error('Failed to load app version', error);
        }
    };

    let initLanguageSetting = (language) => {
        const languageSelect = document.querySelector('#language-select');
        const safeLanguage = language === 'ru' ? 'ru' : 'en';
        currentLanguage = safeLanguage;
        if (languageSelect) {
            languageSelect.value = safeLanguage;
            languageSelect.addEventListener('change', (event) => {
                const selected = event.currentTarget.value === 'ru' ? 'ru' : 'en';
                currentLanguage = selected;
                saveLanguageSettings(selected);
                applyLocalization();
            });
        }
        applyLocalization();
    };

    let initZoomSetting = (zoomValue) => {
        const zoomSelect = document.querySelector('#ui-zoom-select');
        const safeZoom = ['80', '90', '100', '110', '120', '130'].includes(String(zoomValue))
            ? String(zoomValue)
            : '100';
        applyUiZoom(safeZoom);
        if (zoomSelect) {
            zoomSelect.value = safeZoom;
            zoomSelect.addEventListener('change', (event) => {
                const selected = event.currentTarget.value;
                applyUiZoom(selected);
                saveZoomSettings(selected);
            });
        }
    };

    let setupDivider = () => {
        let lastLeftRatio = 0.5;
        const divider = document.getElementById('split-divider');
        const leftPane = document.getElementById('edit');
        const rightPane = document.getElementById('preview');
        const container = document.getElementById('container');

        let isDragging = false;

        divider.addEventListener('mouseenter', () => {
            divider.classList.add('hover');
        });

        divider.addEventListener('mouseleave', () => {
            if (!isDragging) {
                divider.classList.remove('hover');
            }
        });

        divider.addEventListener('mousedown', () => {
            isDragging = true;
            divider.classList.add('active');
            document.body.style.cursor = 'col-resize';
        });

        divider.addEventListener('dblclick', () => {
            const containerRect = container.getBoundingClientRect();
            const totalWidth = containerRect.width;
            const dividerWidth = divider.offsetWidth;
            const halfWidth = (totalWidth - dividerWidth) / 2;

            leftPane.style.width = halfWidth + 'px';
            rightPane.style.width = halfWidth + 'px';
        });

        document.addEventListener('mousemove', (e) => {
            if (!isDragging) return;
            document.body.style.userSelect = 'none';
            const containerRect = container.getBoundingClientRect();
            const totalWidth = containerRect.width;
            const offsetX = e.clientX - containerRect.left;
            const dividerWidth = divider.offsetWidth;

            // Prevent overlap or out-of-bounds
            const minWidth = 100;
            const maxWidth = totalWidth - minWidth - dividerWidth;
            const leftWidth = Math.max(minWidth, Math.min(offsetX, maxWidth));
            leftPane.style.width = leftWidth + 'px';
            rightPane.style.width = (totalWidth - leftWidth - dividerWidth) + 'px';
            lastLeftRatio = leftWidth / (totalWidth - dividerWidth);
        });

        document.addEventListener('mouseup', () => {
            if (isDragging) {
                isDragging = false;
                divider.classList.remove('active');
                divider.classList.remove('hover');
                document.body.style.cursor = 'default';
                document.body.style.userSelect = '';
            }
        });

        window.addEventListener('resize', () => {
            const containerRect = container.getBoundingClientRect();
            const totalWidth = containerRect.width;
            const dividerWidth = divider.offsetWidth;
            const availableWidth = totalWidth - dividerWidth;

            const newLeft = availableWidth * lastLeftRatio;
            const newRight = availableWidth * (1 - lastLeftRatio);

            leftPane.style.width = newLeft + 'px';
            rightPane.style.width = newRight + 'px';
        });
    };

    // ----- entry point -----
    let lastContent = loadLastContent();
    let editor = setupEditor();
    setupPreviewScrollSync(editor);
    setupTabsUi(editor);

    const languageSettings = loadLanguageSettings();
    initLanguageSetting(languageSettings === 'ru' ? 'ru' : 'en');
    const zoomSettings = loadZoomSettings();
    initZoomSetting(zoomSettings);

    let savedTabsState = loadTabsState();
    if (savedTabsState && Array.isArray(savedTabsState.tabs) && savedTabsState.tabs.length > 0) {
        const restoredIdMap = new Map();
        savedTabsState.tabs.forEach((savedTab, index) => {
            const restoredTab = createTab({
                title: savedTab.title || getUntitledTitle(index + 1),
                content: savedTab.content || '',
                filePath: savedTab.filePath || null,
                dirty: savedTab.dirty === true,
                lastSavedContent: typeof savedTab.lastSavedContent === 'string'
                    ? savedTab.lastSavedContent
                    : (savedTab.content || '')
            });
            tabs.push(restoredTab);
            if (savedTab.id) {
                restoredIdMap.set(savedTab.id, restoredTab.id);
            }
        });

        activeTabId = restoredIdMap.get(savedTabsState.activeTabId)
            || (tabs[0] ? tabs[0].id : null);
        const activeTab = getActiveTab();
        const restoredContent = activeTab ? activeTab.content : defaultInput;
        renderTabs();
        presetValue(restoredContent);
        saveLastContent(restoredContent);
    } else {
        const initialContent = lastContent || defaultInput;
        const initialTab = createTab({
            title: getUntitledTitle(1),
            content: initialContent,
            filePath: null,
            dirty: false,
            lastSavedContent: initialContent
        });
        tabs.push(initialTab);
        activeTabId = initialTab.id;
        renderTabs();
        presetValue(initialContent);
    }

    setupHeaderMenus();
    setupAppVersion();
    setupOpenButton();
    setupSaveButton(editor);
    setupSaveAsButton(editor);
    setupResetButton();
    setupCopyButton(editor);
    setupExportButton();
    setupFormatToolbar(editor);

    let scrollBarSettings = loadScrollBarSettings() || false;
    initScrollBarSync(scrollBarSettings);

    // initialize theme (dark/light)
    let themeSettings = loadThemeSettings();
    // normalize to boolean (saved value may be string or boolean)
    if (themeSettings === 'true' || themeSettings === true) {
        themeSettings = true;
    } else {
        themeSettings = false;
    }
    initThemeToggle(themeSettings);

    setupDivider();

    window.addEventListener('beforeunload', () => {
        if (persistTabsTimer !== null) {
            window.clearTimeout(persistTabsTimer);
            persistTabsTimer = null;
        }
        saveTabsState();
    });
};

window.addEventListener("load", () => {
    init();
});
