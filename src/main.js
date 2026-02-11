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
    let previewEditMode = false;
    let previewEditSyncTimer = null;
    let syncingFromPreviewEdit = false;
    let previewSavedRange = null;
    let previousEditPaneWidth = '';
    let previousPreviewPaneWidth = '';

    const localStorageNamespace = 'com.lectr';
    const localStorageKey = 'last_state';
    const localStorageScrollBarKey = 'scroll_bar_settings';
    const localStorageThemeKey = 'theme_settings';
    const localStorageLanguageKey = 'language_settings';
    const localStorageZoomKey = 'ui_zoom_settings';
    const localStorageTabsStateKey = 'tabs_state';
    const localStoragePreviewEditModeKey = 'preview_edit_mode_settings';
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
            previewEdit: 'Edit in preview',
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
            linkedFileNotFound: 'Cannot open linked file from preview.',
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
            previewEdit: 'Правка в preview',
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
            linkedFileNotFound: 'Не удалось открыть связанный файл из preview.',
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
    const defaultInputEn = `# Lectr Markdown guide

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

    const defaultInputRu = `# Гайд по Markdown в Lectr

## Заголовки

# Это заголовок h1
## Это заголовок h2
###### Это заголовок h6

## Выделение текста

*Этот текст будет курсивом*  
_Этот текст тоже будет курсивом_

**Этот текст будет жирным**  
__Этот текст тоже будет жирным__

_Можно **комбинировать** стили_

## Списки

### Маркированный

* Пункт 1
* Пункт 2
* Пункт 2a
* Пункт 2b
    * Пункт 3a
    * Пункт 3b

### Нумерованный

1. Пункт 1
2. Пункт 2
3. Пункт 3
    1. Пункт 3a
    2. Пункт 3b

## Изображения

![Пример alt-текста](image/Markdown-mark.svg "Пример изображения")

## Ссылки

Пишите документацию в **Lectr** с помощью Markdown.

## Цитаты

> Markdown — легковесный язык разметки с простым текстовым синтаксисом.
>
>> Его часто используют для README, заметок и сообщений на форумах.

## Таблицы

| Левый столбец | Правый столбец |
| ------------- |:--------------:|
| левый foo     | правый foo     |
| левый bar     | правый bar     |
| левый baz     | правый baz     |

## Блок кода

${"`"}${"`"}${"`"}
let message = 'Привет, мир';
alert(message);
${"`"}${"`"}${"`"}

## Встроенный код

На этом сайте используется ${"`"}markedjs/marked${"`"}.
`;

    const getDefaultInput = () => (currentLanguage === 'ru' ? defaultInputRu : defaultInputEn);

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
            ['preview-edit-label', 'previewEdit'],
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
            if (syncingFromPreviewEdit) {
                return;
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

    let setPreviewEditableState = (enabled) => {
        const output = document.querySelector('#output');
        if (!output) {
            return;
        }

        if (enabled) {
            output.setAttribute('contenteditable', 'true');
            output.setAttribute('spellcheck', 'true');
            output.classList.add('preview-editable');
            return;
        }

        output.removeAttribute('contenteditable');
        output.removeAttribute('spellcheck');
        output.classList.remove('preview-editable');
    };

    let setPreviewEditLayout = (enabled) => {
        const body = document.body;
        const editPane = document.querySelector('#edit');
        const previewPane = document.querySelector('#preview');
        if (!body || !editPane || !previewPane) {
            return;
        }

        if (enabled) {
            previousEditPaneWidth = editPane.style.width || '';
            previousPreviewPaneWidth = previewPane.style.width || '';
            body.classList.add('preview-edit-mode');
            previewPane.style.width = '100%';
            return;
        }

        body.classList.remove('preview-edit-mode');
        editPane.style.width = previousEditPaneWidth;
        previewPane.style.width = previousPreviewPaneWidth;
    };

    let getPreviewOutputElement = () => document.querySelector('#output');

    let isRangeInsideElement = (range, element) => {
        if (!range || !element) {
            return false;
        }
        return element.contains(range.commonAncestorContainer);
    };

    let savePreviewSelectionRange = () => {
        const output = getPreviewOutputElement();
        const selection = window.getSelection();
        if (!output || !selection || selection.rangeCount === 0) {
            return;
        }

        const range = selection.getRangeAt(0);
        if (!isRangeInsideElement(range, output)) {
            return;
        }
        previewSavedRange = range.cloneRange();
    };

    let restorePreviewSelectionRange = () => {
        const output = getPreviewOutputElement();
        const selection = window.getSelection();
        if (!output || !selection) {
            return false;
        }

        if (previewSavedRange && isRangeInsideElement(previewSavedRange, output)) {
            selection.removeAllRanges();
            selection.addRange(previewSavedRange);
            return true;
        }

        const range = document.createRange();
        range.selectNodeContents(output);
        range.collapse(false);
        selection.removeAllRanges();
        selection.addRange(range);
        previewSavedRange = range.cloneRange();
        return true;
    };

    let executePreviewCommand = (command, value = null) => {
        const output = getPreviewOutputElement();
        if (!output || !previewEditMode) {
            return false;
        }

        output.focus();
        restorePreviewSelectionRange();
        const success = value === null
            ? document.execCommand(command)
            : document.execCommand(command, false, value);
        savePreviewSelectionRange();
        return success;
    };

    let insertHtmlIntoPreviewSelection = (html) => {
        const output = getPreviewOutputElement();
        if (!output || !previewEditMode) {
            return false;
        }

        output.focus();
        restorePreviewSelectionRange();
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0) {
            return false;
        }

        const range = selection.getRangeAt(0);
        if (!isRangeInsideElement(range, output)) {
            return false;
        }

        const fragment = range.createContextualFragment(html);
        const lastNode = fragment.lastChild;
        range.deleteContents();
        range.insertNode(fragment);

        if (lastNode) {
            const afterRange = document.createRange();
            afterRange.setStartAfter(lastNode);
            afterRange.collapse(true);
            selection.removeAllRanges();
            selection.addRange(afterRange);
            previewSavedRange = afterRange.cloneRange();
            return true;
        }

        savePreviewSelectionRange();
        return true;
    };

    let escapeHtml = (text) => String(text || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');

    let escapeMarkdownText = (text) => text
        .replace(/\\/g, '\\\\')
        .replace(/([`*_{}\[\]()#+\-.!|>])/g, '\\$1');

    let inlineNodeToMarkdown = (node) => {
        if (!node) {
            return '';
        }

        if (node.nodeType === Node.TEXT_NODE) {
            return node.textContent || '';
        }

        if (node.nodeType !== Node.ELEMENT_NODE) {
            return '';
        }

        const element = node;
        const tag = element.tagName.toLowerCase();
        const childrenText = Array.from(element.childNodes).map((child) => inlineNodeToMarkdown(child)).join('');

        if (tag === 'br') {
            return '\n';
        }
        if (tag === 'strong' || tag === 'b') {
            return `**${childrenText}**`;
        }
        if (tag === 'em' || tag === 'i') {
            return `*${childrenText}*`;
        }
        if (tag === 'code' && element.parentElement?.tagName.toLowerCase() !== 'pre') {
            return `\`${childrenText}\``;
        }
        if (tag === 'a') {
            const href = element.getAttribute('href') || '';
            const label = childrenText || href;
            return `[${label}](${href})`;
        }

        return childrenText;
    };

    let inlineElementToMarkdown = (element) => {
        const raw = Array.from(element.childNodes).map((node) => inlineNodeToMarkdown(node)).join('');
        return raw
            .replace(/\u00a0/g, ' ')
            .replace(/[ \t]+\n/g, '\n')
            .replace(/\n{3,}/g, '\n\n')
            .trim();
    };

    let listItemToMarkdown = (itemElement, index, ordered) => {
        const clone = itemElement.cloneNode(true);
        clone.querySelectorAll('ul,ol').forEach((nestedList) => nestedList.remove());
        const itemText = inlineElementToMarkdown(clone) || ' ';
        const prefix = ordered ? `${index + 1}. ` : '- ';
        return `${prefix}${itemText}`;
    };

    let tableToMarkdown = (tableElement) => {
        const rowElements = Array.from(tableElement.querySelectorAll('tr'));
        if (rowElements.length === 0) {
            return '';
        }

        const rows = rowElements.map((rowElement) => {
            const cells = Array.from(rowElement.querySelectorAll('th,td'))
                .map((cell) => inlineElementToMarkdown(cell).replace(/\|/g, '\\|').trim());
            return cells;
        }).filter((cells) => cells.length > 0);

        if (rows.length === 0) {
            return '';
        }

        const maxColumns = rows.reduce((max, row) => Math.max(max, row.length), 0);
        const normalizedRows = rows.map((row) => {
            const filled = [...row];
            while (filled.length < maxColumns) {
                filled.push('');
            }
            return filled;
        });

        const header = normalizedRows[0];
        const separator = new Array(maxColumns).fill('---');
        const body = normalizedRows.slice(1);

        const asRow = (cells) => `| ${cells.join(' | ')} |`;
        const lines = [asRow(header), asRow(separator)];
        body.forEach((row) => lines.push(asRow(row)));
        return lines.join('\n');
    };

    let blockToMarkdown = (node) => {
        if (!node) {
            return '';
        }

        if (node.nodeType === Node.TEXT_NODE) {
            const text = (node.textContent || '').trim();
            return text ? escapeMarkdownText(text) : '';
        }

        if (node.nodeType !== Node.ELEMENT_NODE) {
            return '';
        }

        const element = node;
        const tag = element.tagName.toLowerCase();

        if (tag === 'h1' || tag === 'h2' || tag === 'h3' || tag === 'h4' || tag === 'h5' || tag === 'h6') {
            const level = Number.parseInt(tag.slice(1), 10);
            return `${'#'.repeat(level)} ${inlineElementToMarkdown(element)}`.trimEnd();
        }

        if (tag === 'p') {
            return inlineElementToMarkdown(element);
        }

        if (tag === 'blockquote') {
            const quotedBlocks = Array.from(element.children).map((child) => blockToMarkdown(child)).filter(Boolean);
            const quoted = quotedBlocks.length > 0 ? quotedBlocks.join('\n\n') : inlineElementToMarkdown(element);
            return quoted
                .split('\n')
                .map((line) => `> ${line}`)
                .join('\n');
        }

        if (tag === 'ul') {
            const items = Array.from(element.children)
                .filter((child) => child.tagName && child.tagName.toLowerCase() === 'li')
                .map((item, index) => listItemToMarkdown(item, index, false));
            return items.join('\n');
        }

        if (tag === 'ol') {
            const items = Array.from(element.children)
                .filter((child) => child.tagName && child.tagName.toLowerCase() === 'li')
                .map((item, index) => listItemToMarkdown(item, index, true));
            return items.join('\n');
        }

        if (tag === 'pre') {
            const codeElement = element.querySelector('code');
            const code = (codeElement ? codeElement.textContent : element.textContent || '').replace(/\n$/, '');
            return `\`\`\`\n${code}\n\`\`\``;
        }

        if (tag === 'hr') {
            return '---';
        }

        if (tag === 'table') {
            return tableToMarkdown(element);
        }

        return inlineElementToMarkdown(element);
    };

    let previewOutputToMarkdown = () => {
        const output = document.querySelector('#output');
        if (!output) {
            return '';
        }

        const blocks = [];
        output.childNodes.forEach((node) => {
            const markdownBlock = blockToMarkdown(node).trim();
            if (markdownBlock) {
                blocks.push(markdownBlock);
            }
        });

        return blocks.join('\n\n').replace(/\n{3,}/g, '\n\n');
    };

    let syncEditorFromPreview = (editor) => {
        if (!previewEditMode) {
            return;
        }

        const markdown = previewOutputToMarkdown();
        const activeTab = getActiveTab();
        if (!activeTab) {
            return;
        }

        syncingFromPreviewEdit = true;
        suppressEditorChange = true;
        editor.setValue(markdown);
        suppressEditorChange = false;
        syncingFromPreviewEdit = false;

        activeTab.content = markdown;
        activeTab.dirty = markdown !== activeTab.lastSavedContent;
        renderTabs();
        saveLastContent(markdown);
        lastRenderedMarkdown = markdown;
    };

    let setupPreviewEditSync = (editor) => {
        const output = document.querySelector('#output');
        if (!output) {
            return;
        }

        output.addEventListener('input', () => {
            if (!previewEditMode) {
                return;
            }

            if (previewEditSyncTimer !== null) {
                window.clearTimeout(previewEditSyncTimer);
            }

            previewEditSyncTimer = window.setTimeout(() => {
                previewEditSyncTimer = null;
                syncEditorFromPreview(editor);
            }, 220);
        });
    };

    let setupPreviewSelectionTracking = () => {
        document.addEventListener('selectionchange', () => {
            if (!previewEditMode) {
                return;
            }
            savePreviewSelectionRange();
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
        const output = document.querySelector('#output');
        if (!output) {
            return;
        }
        output.innerHTML = sanitized;
        previewSavedRange = null;
        setPreviewEditableState(previewEditMode);
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
        const resetContent = getDefaultInput();
        if (activeTab) {
            activeTab.content = resetContent;
            activeTab.dirty = activeTab.content !== activeTab.lastSavedContent;
            renderTabs();
        }
        presetValue(resetContent);
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

    let getSuggestedPdfName = () => {
        const activeTab = getActiveTab();
        const fallback = 'markdown-preview.pdf';
        if (!activeTab || !activeTab.title) {
            return fallback;
        }

        const rawTitle = String(activeTab.title).trim();
        if (!rawTitle) {
            return fallback;
        }

        const baseName = rawTitle.replace(/\.[^./\\]+$/, '');
        return `${baseName || 'markdown-preview'}.pdf`;
    };

    let exportPreviewToPdf = () => {
        const outputElement = document.querySelector('#output');
        if (!outputElement) {
            return;
        }

        if (window.lectrDesktop && typeof window.lectrDesktop.exportPreviewPdf === 'function') {
            getLightMarkdownCss().then((lightCss) => {
                return window.lectrDesktop.exportPreviewPdf({
                    html: outputElement.innerHTML,
                    lightCss,
                    suggestedName: getSuggestedPdfName()
                });
            }).catch((error) => {
                // eslint-disable-next-line no-console
                console.error('Failed to export PDF via desktop bridge', error);
            });
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
#output, body {
  background: #fff !important;
  color: #24292f !important;
}`;
                            clonedDoc.head.appendChild(style);
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
                .from(outputElement)
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

    let moveTab = (fromIndex, toIndex) => {
        if (fromIndex < 0 || toIndex < 0 || fromIndex >= tabs.length || toIndex >= tabs.length || fromIndex === toIndex) {
            return;
        }
        const [tab] = tabs.splice(fromIndex, 1);
        tabs.splice(toIndex, 0, tab);
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
            createUntitledTab(getDefaultInput());
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
        const dragState = {
            pointerId: null,
            tabId: null,
            tabElement: null,
            pointerOffsetX: 0,
            translateX: 0,
            started: false
        };

        const clearReorderStyles = () => {
            if (!tabsList) {
                return;
            }
            tabsList.querySelectorAll('.tab-item').forEach((element) => {
                element.classList.remove('reorder-anim');
                if (!element.classList.contains('dragging-pointer')) {
                    element.style.transform = '';
                }
            });
        };

        const animateReorder = (beforeLeftMap) => {
            if (!tabsList) {
                return;
            }
            const afterElements = Array.from(tabsList.querySelectorAll('.tab-item'));
            afterElements.forEach((element) => {
                if (element === dragState.tabElement) {
                    return;
                }
                const tabId = element.getAttribute('data-tab-id');
                if (!tabId || !beforeLeftMap.has(tabId)) {
                    return;
                }
                const previousLeft = beforeLeftMap.get(tabId);
                const nextLeft = element.getBoundingClientRect().left;
                const deltaX = previousLeft - nextLeft;
                if (Math.abs(deltaX) < 0.5) {
                    return;
                }

                element.classList.add('reorder-anim');
                element.style.transform = `translateX(${deltaX}px)`;
                requestAnimationFrame(() => {
                    element.style.transform = 'translateX(0)';
                });
            });
        };

        const reorderByPointer = () => {
            if (!tabsList || !dragState.tabElement || !dragState.tabId) {
                return;
            }

            const items = Array.from(tabsList.querySelectorAll('.tab-item'));
            const dragged = dragState.tabElement;
            const currentIndex = tabs.findIndex((tab) => tab.id === dragState.tabId);
            if (currentIndex === -1) {
                return;
            }

            const draggedRect = dragged.getBoundingClientRect();
            const draggedCenter = draggedRect.left + draggedRect.width / 2;
            const withoutDragged = items.filter((element) => element !== dragged);

            let insertIndex = withoutDragged.length;
            for (let index = 0; index < withoutDragged.length; index += 1) {
                const element = withoutDragged[index];
                const rect = element.getBoundingClientRect();
                const center = rect.left + rect.width / 2;
                if (draggedCenter < center) {
                    insertIndex = index;
                    break;
                }
            }

            const targetIndex = Math.min(tabs.length - 1, insertIndex);
            if (targetIndex === currentIndex) {
                return;
            }

            const beforeLeftMap = new Map();
            items.forEach((element) => {
                const tabId = element.getAttribute('data-tab-id');
                if (tabId) {
                    beforeLeftMap.set(tabId, element.getBoundingClientRect().left);
                }
            });

            moveTab(currentIndex, targetIndex);

            const nextTab = tabs[targetIndex + 1];
            if (nextTab) {
                const nextElement = tabsList.querySelector(`[data-tab-id="${nextTab.id}"]`);
                if (nextElement) {
                    tabsList.insertBefore(dragged, nextElement);
                }
            } else {
                tabsList.appendChild(dragged);
            }

            animateReorder(beforeLeftMap);
            schedulePersistTabsState();
        };

        const beginDrag = (event, tabElement) => {
            if (!tabsList || dragState.tabElement) {
                return;
            }
            if (event.button !== 0 || !tabElement) {
                return;
            }

            const tabId = tabElement.getAttribute('data-tab-id');
            if (!tabId) {
                return;
            }

            const rect = tabElement.getBoundingClientRect();
            dragState.pointerId = event.pointerId;
            dragState.tabId = tabId;
            dragState.tabElement = tabElement;
            dragState.pointerOffsetX = event.clientX - rect.left;
            dragState.translateX = 0;
            dragState.started = true;

            tabElement.classList.add('dragging-pointer');
            tabElement.setPointerCapture(event.pointerId);
            document.body.classList.add('tabs-dragging');
        };

        const updateDrag = (event) => {
            if (!tabsList || !dragState.started || dragState.pointerId !== event.pointerId || !dragState.tabElement) {
                return;
            }

            const dragged = dragState.tabElement;
            const rect = dragged.getBoundingClientRect();
            const desiredLeft = event.clientX - dragState.pointerOffsetX;
            const deltaX = desiredLeft - rect.left;
            dragState.translateX += deltaX;
            dragged.style.transform = `translateX(${dragState.translateX}px)`;

            reorderByPointer();
        };

        const endDrag = (event) => {
            if (!dragState.started || dragState.pointerId !== event.pointerId) {
                return;
            }

            if (dragState.tabElement) {
                dragState.tabElement.classList.remove('dragging-pointer');
                dragState.tabElement.style.transform = '';
                if (dragState.tabElement.hasPointerCapture(event.pointerId)) {
                    dragState.tabElement.releasePointerCapture(event.pointerId);
                }
            }

            dragState.pointerId = null;
            dragState.tabId = null;
            dragState.tabElement = null;
            dragState.pointerOffsetX = 0;
            dragState.translateX = 0;
            dragState.started = false;
            document.body.classList.remove('tabs-dragging');
            clearReorderStyles();
        };

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

            tabsList.addEventListener('pointerdown', (event) => {
                const closeTarget = event.target.closest('[data-tab-close]');
                if (closeTarget) {
                    return;
                }
                const tabTarget = event.target.closest('[data-tab-id]');
                if (!tabTarget) {
                    return;
                }
                beginDrag(event, tabTarget);
            });

            tabsList.addEventListener('pointermove', (event) => {
                updateDrag(event);
            });

            tabsList.addEventListener('pointerup', (event) => {
                endDrag(event);
            });

            tabsList.addEventListener('pointercancel', (event) => {
                endDrag(event);
            });

            tabsList.addEventListener('dragstart', (event) => {
                event.preventDefault();
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

    let openFileByPath = async ({ linkTarget, sourceFilePath }) => {
        if (window.lectrDesktop && typeof window.lectrDesktop.openMarkdownFileByPath === 'function') {
            try {
                const result = await window.lectrDesktop.openMarkdownFileByPath({
                    linkTarget,
                    sourceFilePath
                });
                if (result && result.opened) {
                    return result;
                }
            } catch (error) {
                // eslint-disable-next-line no-console
                console.error('Failed to open linked file via desktop bridge', error);
            }
        }

        return { opened: false };
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

    let saveActiveTabWithMode = async (editor, forceDialog) => {
        const activeTab = getActiveTab();
        if (!activeTab) {
            return { saved: false };
        }
        syncActiveTabFromEditor(editor);
        return saveTab(activeTab, { forceDialog, showToast: true });
    };

    let applyOpenedFileResult = (result) => {
        if (!result || !result.opened) {
            return false;
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
            return true;
        }

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
        return true;
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
            applyOpenedFileResult(result);
            closeOpenMenus();
        });
    };

    let setupPreviewLinkNavigation = () => {
        const output = document.querySelector('#output');
        if (!output) {
            return;
        }

        output.addEventListener('click', async (event) => {
            if (previewEditMode) {
                return;
            }

            const target = event.target;
            if (!(target instanceof Element)) {
                return;
            }

            const anchor = target.closest('a[href]');
            if (!anchor) {
                return;
            }

            const href = (anchor.getAttribute('href') || '').trim();
            if (!href) {
                return;
            }

            if (href.startsWith('#')) {
                return;
            }

            if (/^(?:https?:|mailto:|tel:|data:|blob:)/i.test(href)) {
                return;
            }

            const activeTab = getActiveTab();
            const sourceFilePath = activeTab && activeTab.filePath ? activeTab.filePath : null;
            if (!sourceFilePath && !href.startsWith('/') && !href.startsWith('file://')) {
                return;
            }

            event.preventDefault();
            const linkedResult = await openFileByPath({
                linkTarget: href,
                sourceFilePath
            });

            if (!applyOpenedFileResult(linkedResult)) {
                window.alert(t('linkedFileNotFound'));
            }
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
            await saveActiveTabWithMode(editor, false);
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
            await saveActiveTabWithMode(editor, true);
            closeOpenMenus();
        });
    };

    let setupGlobalShortcuts = (editor) => {
        const isEditableField = (target) => {
            if (!(target instanceof HTMLElement)) {
                return false;
            }
            if (target.isContentEditable) {
                return true;
            }
            const tagName = target.tagName.toLowerCase();
            return tagName === 'input' || tagName === 'textarea' || tagName === 'select';
        };

        const isPreviewEditableTarget = (target) => {
            const output = getPreviewOutputElement();
            if (!output || !(target instanceof Node)) {
                return false;
            }
            return target === output || output.contains(target);
        };

        const isPreviewSelectionActive = () => {
            const output = getPreviewOutputElement();
            const selection = window.getSelection();
            if (!output || !selection || selection.rangeCount === 0) {
                return false;
            }
            return output.contains(selection.anchorNode);
        };

        document.addEventListener('keydown', async (event) => {
            const isMod = event.ctrlKey || event.metaKey;
            if (!isMod) {
                return;
            }

            const key = event.key.toLowerCase();
            const target = event.target;
            const editableField = isEditableField(target);
            const previewTarget = isPreviewEditableTarget(target) || isPreviewSelectionActive();

            if (previewEditMode && previewTarget) {
                let previewFormatType = null;

                if (!event.shiftKey && key === 'b') {
                    previewFormatType = 'bold';
                } else if (!event.shiftKey && key === 'i') {
                    previewFormatType = 'italic';
                } else if (!event.shiftKey && key === 'k') {
                    previewFormatType = 'link';
                } else if (event.shiftKey && event.code === 'KeyX') {
                    previewFormatType = 'strikethrough';
                } else if (event.shiftKey && event.code === 'KeyC') {
                    previewFormatType = 'code';
                } else if (event.shiftKey && event.code === 'Digit1') {
                    previewFormatType = 'h1';
                } else if (event.shiftKey && event.code === 'Digit2') {
                    previewFormatType = 'h2';
                } else if (event.shiftKey && event.code === 'Digit8') {
                    previewFormatType = 'ul';
                } else if (event.shiftKey && event.code === 'Digit7') {
                    previewFormatType = 'ol';
                } else if (event.shiftKey && event.code === 'Period') {
                    previewFormatType = 'quote';
                }

                if (previewFormatType) {
                    event.preventDefault();
                    applyPreviewFormat(previewFormatType, editor);
                    return;
                }
            }

            if (key === 's') {
                event.preventDefault();
                await saveActiveTabWithMode(editor, event.shiftKey);
                return;
            }

            if (key === 'w') {
                event.preventDefault();
                const activeTab = getActiveTab();
                if (activeTab) {
                    await closeTab(activeTab.id, editor);
                }
                return;
            }

            if (key === 't' && !event.shiftKey) {
                if (editableField) {
                    return;
                }
                event.preventDefault();
                createUntitledTab('');
            }
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

    let schedulePreviewSyncFromToolbar = (editor) => {
        if (previewEditSyncTimer !== null) {
            window.clearTimeout(previewEditSyncTimer);
        }
        previewEditSyncTimer = window.setTimeout(() => {
            previewEditSyncTimer = null;
            syncEditorFromPreview(editor);
        }, 80);
    };

    let applyPreviewFormat = (formatType, editor) => {
        if (!previewEditMode) {
            return false;
        }

        const output = getPreviewOutputElement();
        if (!output) {
            return false;
        }

        let handled = true;
        if (formatType === 'bold') {
            executePreviewCommand('bold');
        } else if (formatType === 'italic') {
            executePreviewCommand('italic');
        } else if (formatType === 'strikethrough') {
            executePreviewCommand('strikeThrough');
        } else if (formatType === 'h1') {
            executePreviewCommand('formatBlock', '<h1>');
        } else if (formatType === 'h2') {
            executePreviewCommand('formatBlock', '<h2>');
        } else if (formatType === 'ul') {
            executePreviewCommand('insertUnorderedList');
        } else if (formatType === 'ol') {
            executePreviewCommand('insertOrderedList');
        } else if (formatType === 'quote') {
            executePreviewCommand('formatBlock', '<blockquote>');
        } else if (formatType === 'link') {
            const url = window.prompt('URL', 'https://example.com');
            if (!url) {
                return true;
            }
            restorePreviewSelectionRange();
            const selectedText = (window.getSelection() ? window.getSelection().toString().trim() : '');
            if (selectedText) {
                executePreviewCommand('createLink', url);
            } else {
                insertHtmlIntoPreviewSelection(`<a href="${escapeHtml(url)}">${escapeHtml('link text')}</a>`);
            }
        } else if (formatType === 'code') {
            restorePreviewSelectionRange();
            const selectedText = window.getSelection() ? window.getSelection().toString() : '';
            insertHtmlIntoPreviewSelection(`<code>${escapeHtml(selectedText || 'inline code')}</code>`);
        } else {
            handled = false;
        }

        if (handled) {
            schedulePreviewSyncFromToolbar(editor);
        }
        return handled;
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

        toolbar.addEventListener('mousedown', (event) => {
            if (!previewEditMode) {
                return;
            }
            const button = event.target.closest('.format-button');
            if (!button) {
                return;
            }
            event.preventDefault();
        });

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

            if (previewEditMode && applyPreviewFormat(formatType, editor)) {
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
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyS, () => {
            void saveActiveTabWithMode(editor, false);
        });
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.KeyS, () => {
            void saveActiveTabWithMode(editor, true);
        });
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyW, () => {
            const activeTab = getActiveTab();
            if (activeTab) {
                void closeTab(activeTab.id, editor);
            }
        });
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyT, () => {
            createUntitledTab('');
        });
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

    let loadPreviewEditModeSetting = () => {
        try {
            const raw = localStorage.getItem(`${localStorageNamespace}_${localStoragePreviewEditModeKey}`);
            return raw ? JSON.parse(raw) : null;
        } catch (error) {
            return null;
        }
    };

    let savePreviewEditModeSetting = (enabled) => {
        try {
            localStorage.setItem(`${localStorageNamespace}_${localStoragePreviewEditModeKey}`, JSON.stringify(enabled));
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

        const fallback = versionElement.textContent ? versionElement.textContent.trim() : 'v1.0.1';
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

    let initPreviewEditMode = (enabled, editor) => {
        const checkbox = document.querySelector('#preview-edit-checkbox');
        previewEditMode = enabled === true;
        setPreviewEditableState(previewEditMode);
        setPreviewEditLayout(previewEditMode);
        if (previewEditMode) {
            const output = getPreviewOutputElement();
            if (output) {
                output.focus();
            }
        }

        if (!checkbox) {
            return;
        }

        checkbox.checked = previewEditMode;
        checkbox.addEventListener('change', (event) => {
            const checked = event.currentTarget.checked === true;
            previewEditMode = checked;
            savePreviewEditModeSetting(checked);
            setPreviewEditableState(checked);
            setPreviewEditLayout(checked);
            if (checked) {
                const output = getPreviewOutputElement();
                if (output) {
                    output.focus();
                }
            }

            if (previewEditSyncTimer !== null) {
                window.clearTimeout(previewEditSyncTimer);
                previewEditSyncTimer = null;
            }

            if (!checked) {
                const value = editor.getValue();
                scheduleConvert(value);
                return;
            }

            scheduleConvert(editor.getValue());
        });
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
    setupPreviewEditSync(editor);
    setupPreviewSelectionTracking();
    setupTabsUi(editor);

    const languageSettings = loadLanguageSettings();
    initLanguageSetting(languageSettings === 'ru' ? 'ru' : 'en');
    const zoomSettings = loadZoomSettings();
    initZoomSetting(zoomSettings);
    const previewEditSettings = loadPreviewEditModeSetting();
    initPreviewEditMode(previewEditSettings === true, editor);

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
        const restoredContent = activeTab ? activeTab.content : getDefaultInput();
        renderTabs();
        presetValue(restoredContent);
        saveLastContent(restoredContent);
    } else {
        const initialContent = lastContent || getDefaultInput();
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
    setupPreviewLinkNavigation();
    setupSaveButton(editor);
    setupSaveAsButton(editor);
    setupResetButton();
    setupCopyButton(editor);
    setupExportButton();
    setupFormatToolbar(editor);
    setupGlobalShortcuts(editor);

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
        if (previewEditSyncTimer !== null) {
            window.clearTimeout(previewEditSyncTimer);
            previewEditSyncTimer = null;
        }
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
