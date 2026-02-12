import { marked } from 'marked';
import DOMPurify from 'dompurify';

const init = async () => {
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
    let restoringTabScroll = false;
    let previewEditMode = false;
    let previewEditSyncTimer = null;
    let syncingFromPreviewEdit = false;
    let previewSavedRange = null;
    let previousEditPaneWidth = '';
    let previousPreviewPaneWidth = '';
    let editorIsScrolled = false;
    let previewIsScrolled = false;
    let refreshFormatToolbarState = () => { };
    let refreshOnboardingLocalization = () => { };
    let setLanguagePreference = () => { };
    let setPreviewEditModePreference = () => { };
    let applyThemePreference = () => { };
    let monaco = null;
    let monacoLoadPromise = null;
    let notesArchive = [];
    let activeNoteTooltipRef = null;
    let noteTooltipElement = null;
    const immersionScrollThresholdPx = 18;
    const editorTopInsetExtraPx = 0;

    const localStorageNamespace = 'com.lectr';
    const localStorageKey = 'last_state';
    const localStorageThemeKey = 'theme_settings';
    const localStorageLanguageKey = 'language_settings';
    const localStorageZoomKey = 'ui_zoom_settings';
    const localStorageTabsStateKey = 'tabs_state';
    const localStoragePreviewEditModeKey = 'preview_edit_mode_settings';
    const localStorageOnboardingKey = 'onboarding_v1';
    const localStorageNotesArchiveKey = 'notes_archive';
    const externalLinkPattern = /^(?:https?:|mailto:|tel:)/i;
    const blockedHrefPattern = /^(?:javascript:|vbscript:|data:|file:)/i;
    let isDomPurifyConfigured = false;
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
            onboardingTitle: 'Set up Lectr',
            onboardingStep: 'Step {step} of {total}',
            onboardingLanguageTitle: 'Choose language',
            onboardingLanguageEn: 'English',
            onboardingLanguageRu: 'Russian',
            onboardingModeTitle: 'Choose mode',
            onboardingModeAdvancedLabel: 'Advanced',
            onboardingModeAdvancedDescription: 'Raw markdown editor',
            onboardingModeSimpleLabel: 'Simple (Experimental)',
            onboardingModeSimpleDescription: 'Edit directly in preview',
            onboardingThemeTitle: 'Choose theme',
            onboardingThemeLight: 'Light',
            onboardingThemeDark: 'Dark',
            onboardingBack: 'Back',
            onboardingNext: 'Next',
            onboardingFinish: 'Finish',
            aboutTitle: 'About Lectr',
            aboutVersionLabel: 'Version',
            aboutDeveloperLabel: 'Developer',
            aboutRepositoryLabel: 'Repository',
            aboutLicenseLabel: 'License',
            aboutClose: 'Close',
            aboutRepoLink: 'GitHub',
            tablePopoverTitle: 'Insert table',
            tableColumns: 'Columns',
            tableRows: 'Rows',
            tableAlignment: 'Alignment',
            tableAlignmentLeft: 'Left',
            tableAlignmentCenter: 'Center',
            tableAlignmentRight: 'Right',
            tableIncludeHeader: 'Include header row',
            tableInsert: 'Insert table',
            tableAddColumn: 'Add column',
            tableRemoveColumn: 'Remove column',
            tableAddRow: 'Add row',
            tableRemoveRow: 'Remove row',
            tableDeleteTable: 'Delete table',
            tableNotFound: 'Place cursor inside a markdown table first.',
            tableMinColumns: 'Cannot remove the last table column.',
            tableNoDataRows: 'Cannot remove row: no table rows left.',
            tableColumnPrefix: 'Column',
            linkPopoverTitle: 'Insert link',
            linkText: 'Link text',
            linkTarget: 'File path or URL',
            linkBrowse: 'Browse...',
            linkAnchor: 'Anchor',
            linkAnchorNone: 'No anchor',
            linkInsert: 'Insert link',
            linkUpdate: 'Update link',
            linkRemove: 'Remove link',
            linkDefaultText: 'link text',
            linkDefaultTarget: 'https://example.com',
            linkTargetRequired: 'Enter file path or URL for the link.',
            previewCtrlClickHint: 'Ctrl/Cmd + click to open link',
            imagePopoverTitle: 'Insert image',
            imageAlt: 'Alt text',
            imageSource: 'File path or URL',
            imageBrowse: 'Browse...',
            imageTitle: 'Title (optional)',
            imageInsert: 'Insert image',
            imageDefaultAlt: 'image',
            imageDefaultSource: 'image/example.png',
            imageSourceRequired: 'Enter file path or URL for the image.',
            notePopoverTitle: 'Term archive',
            noteArchive: 'Archive entry',
            noteEmpty: 'No notes found.',
            noteSearch: 'Search',
            noteNew: 'Add note',
            noteEdit: 'Edit note',
            noteDelete: 'Delete entry',
            noteDeleteConfirm: 'Delete selected archive entry?',
            noteTitle: 'Term / title',
            noteBody: 'Definition / note',
            noteSave: 'Save to archive',
            noteEditorAddTitle: 'New archive note',
            noteEditorEditTitle: 'Edit archive note',
            noteEditorCancel: 'Cancel',
            noteBind: 'Bind to selection',
            noteUnbind: 'Remove binding',
            noteDefaultBody: 'Definition text',
            noteSelectionRequired: 'Select text first to bind note.',
            noteTitleRequired: 'Enter term title.',
            noteBodyRequired: 'Enter definition or note text.',
            noteEntryMissing: 'Select archive entry or create a new one.',
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
                link: 'Link',
                image: 'Image',
                table: 'Table',
                note: 'Term note'
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
            onboardingTitle: 'Настройка Lectr',
            onboardingStep: 'Шаг {step} из {total}',
            onboardingLanguageTitle: 'Выберите язык',
            onboardingLanguageEn: 'Английский',
            onboardingLanguageRu: 'Русский',
            onboardingModeTitle: 'Выберите режим',
            onboardingModeAdvancedLabel: 'Продвинутый',
            onboardingModeAdvancedDescription: 'Редактирование raw markdown',
            onboardingModeSimpleLabel: 'Простой (экспериментальный)',
            onboardingModeSimpleDescription: 'Редактирование прямо в preview',
            onboardingThemeTitle: 'Выберите тему',
            onboardingThemeLight: 'Светлая',
            onboardingThemeDark: 'Тёмная',
            onboardingBack: 'Назад',
            onboardingNext: 'Далее',
            onboardingFinish: 'Готово',
            aboutTitle: 'О приложении Lectr',
            aboutVersionLabel: 'Версия',
            aboutDeveloperLabel: 'Разработчик',
            aboutRepositoryLabel: 'Репозиторий',
            aboutLicenseLabel: 'Лицензия',
            aboutClose: 'Закрыть',
            aboutRepoLink: 'GitHub',
            tablePopoverTitle: 'Вставить таблицу',
            tableColumns: 'Столбцы',
            tableRows: 'Строки',
            tableAlignment: 'Выравнивание',
            tableAlignmentLeft: 'Слева',
            tableAlignmentCenter: 'По центру',
            tableAlignmentRight: 'Справа',
            tableIncludeHeader: 'Добавить строку заголовка',
            tableInsert: 'Вставить таблицу',
            tableAddColumn: 'Добавить столбец',
            tableRemoveColumn: 'Удалить столбец',
            tableAddRow: 'Добавить строку',
            tableRemoveRow: 'Удалить строку',
            tableDeleteTable: 'Удалить таблицу',
            tableNotFound: 'Сначала поставьте курсор внутрь markdown-таблицы.',
            tableMinColumns: 'Нельзя удалить последний столбец таблицы.',
            tableNoDataRows: 'Нельзя удалить строку: в таблице не осталось строк данных.',
            tableColumnPrefix: 'Колонка',
            linkPopoverTitle: 'Вставить ссылку',
            linkText: 'Текст ссылки',
            linkTarget: 'Путь к файлу или URL',
            linkBrowse: 'Выбрать...',
            linkAnchor: 'Якорь',
            linkAnchorNone: 'Без якоря',
            linkInsert: 'Вставить ссылку',
            linkUpdate: 'Обновить ссылку',
            linkRemove: 'Удалить ссылку',
            linkDefaultText: 'текст ссылки',
            linkDefaultTarget: 'https://example.com',
            linkTargetRequired: 'Введите путь к файлу или URL для ссылки.',
            previewCtrlClickHint: 'Ctrl/Cmd + клик, чтобы открыть ссылку',
            imagePopoverTitle: 'Вставить изображение',
            imageAlt: 'Alt-текст',
            imageSource: 'Путь к файлу или URL',
            imageBrowse: 'Выбрать...',
            imageTitle: 'Заголовок (необязательно)',
            imageInsert: 'Вставить изображение',
            imageDefaultAlt: 'изображение',
            imageDefaultSource: 'image/example.png',
            imageSourceRequired: 'Введите путь к файлу или URL для изображения.',
            notePopoverTitle: 'Архив терминов',
            noteArchive: 'Запись архива',
            noteEmpty: 'Заметки не найдены.',
            noteSearch: 'Поиск',
            noteNew: 'Добавить заметку',
            noteEdit: 'Редактировать заметку',
            noteDelete: 'Удалить запись',
            noteDeleteConfirm: 'Удалить выбранную запись архива?',
            noteTitle: 'Термин / заголовок',
            noteBody: 'Определение / заметка',
            noteSave: 'Сохранить в архив',
            noteEditorAddTitle: 'Новая запись архива',
            noteEditorEditTitle: 'Редактировать запись архива',
            noteEditorCancel: 'Отмена',
            noteBind: 'Привязать к выделению',
            noteUnbind: 'Убрать привязку',
            noteDefaultBody: 'Текст определения',
            noteSelectionRequired: 'Сначала выделите текст для привязки заметки.',
            noteTitleRequired: 'Введите название термина.',
            noteBodyRequired: 'Введите определение или текст заметки.',
            noteEntryMissing: 'Выберите запись архива или создайте новую.',
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
                link: 'Ссылка',
                image: 'Изображение',
                table: 'Таблица',
                note: 'Заметка-термин'
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
            ['theme-label', 'darkMode'],
            ['preview-edit-label', 'previewEdit'],
            ['language-label', 'language'],
            ['zoom-label', 'scale'],
            ['onboarding-title', 'onboardingTitle'],
            ['onboarding-language-title', 'onboardingLanguageTitle'],
            ['onboarding-language-en-label', 'onboardingLanguageEn'],
            ['onboarding-language-ru-label', 'onboardingLanguageRu'],
            ['onboarding-mode-title', 'onboardingModeTitle'],
            ['onboarding-mode-advanced-label', 'onboardingModeAdvancedLabel'],
            ['onboarding-mode-advanced-description', 'onboardingModeAdvancedDescription'],
            ['onboarding-mode-simple-label', 'onboardingModeSimpleLabel'],
            ['onboarding-mode-simple-description', 'onboardingModeSimpleDescription'],
            ['onboarding-theme-title', 'onboardingThemeTitle'],
            ['onboarding-theme-light-label', 'onboardingThemeLight'],
            ['onboarding-theme-dark-label', 'onboardingThemeDark'],
            ['onboarding-back-button', 'onboardingBack'],
            ['onboarding-next-button', 'onboardingNext'],
            ['about-title', 'aboutTitle'],
            ['about-version-label', 'aboutVersionLabel'],
            ['about-developer-label', 'aboutDeveloperLabel'],
            ['about-repo-label', 'aboutRepositoryLabel'],
            ['about-license-label', 'aboutLicenseLabel'],
            ['about-repo-link', 'aboutRepoLink'],
            ['table-popover-title', 'tablePopoverTitle'],
            ['table-columns-label', 'tableColumns'],
            ['table-rows-label', 'tableRows'],
            ['table-alignment-label', 'tableAlignment'],
            ['table-alignment-left-option', 'tableAlignmentLeft'],
            ['table-alignment-center-option', 'tableAlignmentCenter'],
            ['table-alignment-right-option', 'tableAlignmentRight'],
            ['table-header-label', 'tableIncludeHeader'],
            ['table-insert-button', 'tableInsert'],
            ['table-add-column-button', 'tableAddColumn'],
            ['table-remove-column-button', 'tableRemoveColumn'],
            ['table-add-row-button', 'tableAddRow'],
            ['table-remove-row-button', 'tableRemoveRow'],
            ['link-popover-title', 'linkPopoverTitle'],
            ['link-text-label', 'linkText'],
            ['link-target-label', 'linkTarget'],
            ['link-anchor-label', 'linkAnchor'],
            ['link-anchor-none-option', 'linkAnchorNone'],
            ['link-browse-button', 'linkBrowse'],
            ['link-insert-button', 'linkInsert'],
            ['link-remove-button', 'linkRemove'],
            ['note-popover-title', 'notePopoverTitle'],
            ['note-search-label', 'noteSearch'],
            ['note-new-entry-button', 'noteNew'],
            ['note-bind-button', 'noteBind'],
            ['note-unbind-button', 'noteUnbind'],
            ['note-editor-title-label', 'noteTitle'],
            ['note-editor-body-label', 'noteBody'],
            ['note-editor-save-button', 'noteSave'],
            ['note-editor-cancel-button', 'noteEditorCancel'],
            ['note-editor-delete-button', 'noteDelete'],
            ['image-popover-title', 'imagePopoverTitle'],
            ['image-alt-label', 'imageAlt'],
            ['image-src-label', 'imageSource'],
            ['image-browse-button', 'imageBrowse'],
            ['image-title-label', 'imageTitle'],
            ['image-insert-button', 'imageInsert']
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

        const brandButton = byId('brand-button');
        if (brandButton) {
            brandButton.setAttribute('aria-label', t('aboutTitle'));
            brandButton.setAttribute('title', t('aboutTitle'));
        }

        const aboutCloseButton = byId('about-close-button');
        if (aboutCloseButton) {
            aboutCloseButton.setAttribute('aria-label', t('aboutClose'));
            aboutCloseButton.setAttribute('title', t('aboutClose'));
        }

        const tablePopover = byId('table-popover');
        if (tablePopover) {
            tablePopover.setAttribute('aria-label', t('tablePopoverTitle'));
        }

        const linkPopover = byId('link-popover');
        if (linkPopover) {
            linkPopover.setAttribute('aria-label', t('linkPopoverTitle'));
        }

        const imagePopover = byId('image-popover');
        if (imagePopover) {
            imagePopover.setAttribute('aria-label', t('imagePopoverTitle'));
        }

        const notePopover = byId('note-popover');
        if (notePopover) {
            notePopover.setAttribute('aria-label', t('notePopoverTitle'));
        }

        const noteEditorOverlay = byId('note-editor-overlay');
        if (noteEditorOverlay) {
            noteEditorOverlay.setAttribute('aria-label', t('notePopoverTitle'));
        }

        const noteEditorCloseButton = byId('note-editor-close-button');
        if (noteEditorCloseButton) {
            noteEditorCloseButton.setAttribute('aria-label', t('noteEditorCancel'));
            noteEditorCloseButton.setAttribute('title', t('noteEditorCancel'));
        }

        const noteEditorModalTitle = byId('note-editor-modal-title');
        if (noteEditorModalTitle) {
            const currentMode = noteEditorModalTitle.getAttribute('data-mode') === 'edit' ? 'edit' : 'create';
            noteEditorModalTitle.textContent = currentMode === 'edit'
                ? t('noteEditorEditTitle')
                : t('noteEditorAddTitle');
        }
        const noteEntriesList = byId('note-entries-list');
        if (noteEntriesList) {
            const emptyState = noteEntriesList.querySelector('.note-empty-state');
            if (emptyState) {
                emptyState.textContent = t('noteEmpty');
            }
            const editButtons = Array.from(noteEntriesList.querySelectorAll('.note-entry-edit'));
            editButtons.forEach((button) => {
                button.setAttribute('aria-label', t('noteEdit'));
                button.setAttribute('title', t('noteEdit'));
            });
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

        const output = byId('output');
        if (output) {
            applyPreviewLinkHints(output);
        }

        document.documentElement.setAttribute('lang', currentLanguage === 'ru' ? 'ru' : 'en');
        renderTabs();
        refreshOnboardingLocalization();
    };

    const createTab = ({
        title,
        content,
        filePath = null,
        dirty = false,
        lastSavedContent = null,
        editorScrollTop = 0,
        previewScrollTop = 0
    }) => ({
        id: `tab-${nextTabId++}`,
        title: title || getUntitledTitle(1),
        content: content || '',
        filePath,
        dirty,
        lastSavedContent: lastSavedContent ?? (dirty ? '' : (content || '')),
        editorScrollTop: Number.isFinite(editorScrollTop) ? editorScrollTop : 0,
        previewScrollTop: Number.isFinite(previewScrollTop) ? previewScrollTop : 0
    });

    const getActiveTab = () => tabs.find((tab) => tab.id === activeTabId) || null;

    const updateTabContentState = (tab, content, { renderOnDirtyChange = true } = {}) => {
        if (!tab) {
            return false;
        }

        const nextContent = typeof content === 'string' ? content : '';
        const wasDirty = tab.dirty === true;
        tab.content = nextContent;
        tab.dirty = nextContent !== tab.lastSavedContent;

        const dirtyChanged = tab.dirty !== wasDirty;
        if (dirtyChanged && renderOnDirtyChange) {
            renderTabs();
        }

        return dirtyChanged;
    };

    const getWorkspaceTopOffsetPx = () => {
        const rootStyles = window.getComputedStyle(document.documentElement);
        const rawValue = rootStyles.getPropertyValue('--ui-workspace-top-offset').trim();
        const parsed = Number.parseFloat(rawValue);
        if (Number.isFinite(parsed) && parsed >= 0) {
            return parsed;
        }
        return 120;
    };

    const loadMonaco = async () => {
        if (monaco) {
            return monaco;
        }

        if (!monacoLoadPromise) {
            monacoLoadPromise = import('monaco-editor/esm/vs/editor/editor.api').then((module) => {
                monaco = module;
                return module;
            });
        }

        return monacoLoadPromise;
    };

    const getEditorPaddingOptions = () => ({
        top: Math.round(getWorkspaceTopOffsetPx() + editorTopInsetExtraPx),
        bottom: 16
    });

    const isPastImmersionThreshold = (scrollTop) => Number.isFinite(scrollTop) && scrollTop > immersionScrollThresholdPx;

    const refreshTopImmersionState = () => {
        const previewElement = document.querySelector('#preview');
        previewIsScrolled = !!previewElement && isPastImmersionThreshold(previewElement.scrollTop);
        if (editor && typeof editor.getScrollTop === 'function') {
            editorIsScrolled = isPastImmersionThreshold(editor.getScrollTop());
        }
        updateTopImmersionState();
    };

    const updateTopImmersionState = () => {
        if (!document.body) {
            return;
        }
        const shouldEnable = previewEditMode ? previewIsScrolled : (editorIsScrolled || previewIsScrolled);
        document.body.classList.toggle('content-immersed', shouldEnable);
    };

    const setTopImmersionFromTab = (tab) => {
        const nextEditorTop = Number.isFinite(tab?.editorScrollTop) ? tab.editorScrollTop : 0;
        const nextPreviewTop = Number.isFinite(tab?.previewScrollTop) ? tab.previewScrollTop : 0;
        editorIsScrolled = isPastImmersionThreshold(nextEditorTop);
        previewIsScrolled = isPastImmersionThreshold(nextPreviewTop);
        updateTopImmersionState();
    };

    const captureActiveTabScrollState = () => {
        const activeTab = getActiveTab();
        if (!activeTab) {
            return;
        }
        if (editor && typeof editor.getScrollTop === 'function') {
            activeTab.editorScrollTop = editor.getScrollTop();
        }
        const previewElement = document.querySelector('#preview');
        if (previewElement) {
            activeTab.previewScrollTop = previewElement.scrollTop;
        }
    };

    const applyTabScrollState = (tab) => {
        if (!tab) {
            return;
        }
        setTopImmersionFromTab(tab);
        const previewElement = document.querySelector('#preview');
        restoringTabScroll = true;
        if (previewElement) {
            previewElement.scrollTop = Number.isFinite(tab.previewScrollTop) ? tab.previewScrollTop : 0;
        }
        if (editor && typeof editor.setScrollTop === 'function') {
            editor.setScrollTop(Number.isFinite(tab.editorScrollTop) ? tab.editorScrollTop : 0);
        }
        window.requestAnimationFrame(() => {
            restoringTabScroll = false;
            refreshTopImmersionState();
        });
    };

    const applyEditorViewportTopInset = (editorInstance) => {
        if (!editorInstance || typeof editorInstance.updateOptions !== 'function') {
            return;
        }
        editorInstance.updateOptions({
            padding: getEditorPaddingOptions()
        });
    };

    self.MonacoEnvironment = {
        getWorker(_, label) {
            return new Proxy({}, { get: () => () => { } });
        }
    }

    let setupEditor = async () => {
        await loadMonaco();
        if (!monaco || !monaco.editor) {
            throw new Error('Monaco failed to initialize');
        }

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
            folding: false,
            padding: getEditorPaddingOptions()
        });

        editor.onDidChangeModelContent(() => {
            let value = editor.getValue();
            if (!suppressEditorChange) {
                const activeTab = getActiveTab();
                if (activeTab) {
                    updateTabContentState(activeTab, value);
                }
                saveLastContent(value);
            }
            if (syncingFromPreviewEdit) {
                return;
            }
            scheduleConvert(value);
        });

        editor.onDidScrollChange((e) => {
            editorIsScrolled = isPastImmersionThreshold(e.scrollTop);
            const activeTab = getActiveTab();
            if (activeTab) {
                activeTab.editorScrollTop = e.scrollTop;
            }
            updateTopImmersionState();

            if (restoringTabScroll) {
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
            previewIsScrolled = isPastImmersionThreshold(previewElement.scrollTop);
            const activeTab = getActiveTab();
            if (activeTab) {
                activeTab.previewScrollTop = previewElement.scrollTop;
            }
            updateTopImmersionState();

            if (restoringTabScroll) {
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
        }, { passive: true });
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
            applyPreviewLinkHints(output);
            applyNoteReferenceDecorations(output);
            return;
        }

        output.removeAttribute('contenteditable');
        output.removeAttribute('spellcheck');
        output.classList.remove('preview-editable');
        applyPreviewLinkHints(output);
        applyNoteReferenceDecorations(output);
        hideNoteTooltip();
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
            window.requestAnimationFrame(() => {
                ensurePreviewEditableCaret({ preserveSelection: true });
            });
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

    let ensurePreviewEditableCaret = ({ preserveSelection = true } = {}) => {
        if (!previewEditMode) {
            return;
        }

        const output = getPreviewOutputElement();
        const selection = window.getSelection();
        if (!output || !selection) {
            return;
        }

        if (output.childNodes.length === 0) {
            output.innerHTML = '<p><br></p>';
        }

        output.focus();

        if (preserveSelection && previewSavedRange && isRangeInsideElement(previewSavedRange, output)) {
            selection.removeAllRanges();
            selection.addRange(previewSavedRange);
            return;
        }

        const range = document.createRange();
        range.selectNodeContents(output);
        range.collapse(false);
        selection.removeAllRanges();
        selection.addRange(range);
        previewSavedRange = range.cloneRange();
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

    let unwrapPreviewBlockquoteAtSelection = () => {
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

        const startNode = range.startContainer;
        const contextElement = startNode.nodeType === Node.ELEMENT_NODE
            ? startNode
            : startNode.parentElement;
        if (!(contextElement instanceof Element)) {
            return false;
        }

        const blockquote = contextElement.closest('blockquote');
        if (!blockquote || !output.contains(blockquote)) {
            return false;
        }

        const marker = document.createElement('span');
        marker.setAttribute('data-lectr-caret-marker', '1');
        marker.textContent = '\u200b';

        const collapsedRange = range.cloneRange();
        collapsedRange.collapse(true);
        collapsedRange.insertNode(marker);

        const parent = blockquote.parentNode;
        if (!parent) {
            marker.remove();
            return false;
        }
        while (blockquote.firstChild) {
            parent.insertBefore(blockquote.firstChild, blockquote);
        }
        parent.removeChild(blockquote);

        const nextRange = document.createRange();
        nextRange.setStartAfter(marker);
        nextRange.collapse(true);
        selection.removeAllRanges();
        selection.addRange(nextRange);
        marker.remove();
        previewSavedRange = nextRange.cloneRange();
        return true;
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

    let escapeHtmlAttribute = (text) => escapeHtml(text).replace(/`/g, '&#96;');

    let getNoteEntryById = (noteId) => {
        const safeId = String(noteId || '').trim();
        if (!safeId) {
            return null;
        }
        return notesArchive.find((entry) => entry && entry.id === safeId) || null;
    };

    let ensureNoteTooltipElement = () => {
        if (noteTooltipElement && noteTooltipElement.isConnected) {
            return noteTooltipElement;
        }
        const tooltip = document.createElement('div');
        tooltip.id = 'note-hover-tooltip';
        tooltip.className = 'note-hover-tooltip';
        tooltip.hidden = true;
        document.body.appendChild(tooltip);
        noteTooltipElement = tooltip;
        return tooltip;
    };

    let hideNoteTooltip = () => {
        activeNoteTooltipRef = null;
        const tooltip = ensureNoteTooltipElement();
        tooltip.hidden = true;
        tooltip.innerHTML = '';
    };

    let showNoteTooltip = (targetElement) => {
        if (!(targetElement instanceof Element)) {
            hideNoteTooltip();
            return;
        }
        const noteId = String(targetElement.getAttribute('data-lectr-note-id') || '').trim();
        const entry = getNoteEntryById(noteId);
        if (!entry) {
            hideNoteTooltip();
            return;
        }

        const tooltip = ensureNoteTooltipElement();
        tooltip.innerHTML = `<strong>${escapeHtml(entry.title)}</strong><span>${escapeHtml(entry.body)}</span>`;
        tooltip.hidden = false;

        const rect = targetElement.getBoundingClientRect();
        const tooltipRect = tooltip.getBoundingClientRect();
        const top = Math.max(8, rect.bottom + 8);
        const left = Math.max(8, Math.min(rect.left, window.innerWidth - tooltipRect.width - 8));
        tooltip.style.top = `${top}px`;
        tooltip.style.left = `${left}px`;
        activeNoteTooltipRef = targetElement;
    };

    let normalizeMarkdownContent = (markdown) => {
        const raw = typeof markdown === 'string' ? markdown : '';
        if (!raw) {
            return '';
        }

        const removableTokens = new Set([
            '**',
            '****',
            '__',
            '____',
            '~~',
            '~~~~',
            '``',
            '````'
        ]);

        const normalized = raw
            .replace(/\r\n?/g, '\n')
            .split('\n')
            .map((line) => {
                const trimmed = line.trim();
                if (removableTokens.has(trimmed)) {
                    return '';
                }
                return line;
            })
            .join('\n');
        const normalizedLinks = normalized.replace(
            /\(([^()\n]+)\)\[([^\]\n]*[/.#:][^\]\n]*)\]/g,
            (_match, label, href) => `[${String(label).trim()}](${String(href).trim()})`
        );

        return normalizedLinks.replace(/\n{3,}/g, '\n\n');
    };

    let normalizeHeadingLabelText = (rawLabel) => {
        return String(rawLabel || '')
            .trim()
            .replace(/\[(.*?)\]\((.*?)\)/g, '$1')
            .replace(/[`*_~]/g, '')
            .trim();
    };

    let slugifyHeadingId = (value) => {
        const base = String(value || '')
            .trim()
            .toLowerCase()
            .normalize('NFKD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/[^\p{L}\p{N}\s-]/gu, '')
            .replace(/\s+/g, '-')
            .replace(/-+/g, '-')
            .replace(/^-|-$/g, '');
        return base || 'section';
    };

    let applyHeadingIdsToElement = (rootElement) => {
        if (!(rootElement instanceof Element)) {
            return;
        }

        const slugCounts = new Map();
        const headings = Array.from(rootElement.querySelectorAll('h1,h2,h3,h4,h5,h6'));
        headings.forEach((headingElement) => {
            const label = normalizeHeadingLabelText(headingElement.textContent || '');
            if (!label) {
                return;
            }

            const baseSlug = slugifyHeadingId(label);
            const duplicateCount = slugCounts.get(baseSlug) || 0;
            slugCounts.set(baseSlug, duplicateCount + 1);
            const headingId = duplicateCount === 0 ? baseSlug : `${baseSlug}-${duplicateCount}`;
            headingElement.setAttribute('id', headingId);
        });
    };

    let applyPreviewLinkHints = (rootElement) => {
        if (!(rootElement instanceof Element)) {
            return;
        }

        const anchors = Array.from(rootElement.querySelectorAll('a[href]'));
        anchors.forEach((anchorElement) => {
            if (!(anchorElement instanceof HTMLAnchorElement)) {
                return;
            }

            if (!previewEditMode) {
                if (anchorElement.dataset.lectrHintManaged === '1') {
                    anchorElement.removeAttribute('title');
                    delete anchorElement.dataset.lectrHintManaged;
                }
                return;
            }

            const existingTitle = String(anchorElement.getAttribute('title') || '').trim();
            if (!existingTitle) {
                anchorElement.setAttribute('title', t('previewCtrlClickHint'));
                anchorElement.dataset.lectrHintManaged = '1';
            }
        });
    };

    let applyNoteReferenceDecorations = (rootElement) => {
        if (!(rootElement instanceof Element)) {
            return;
        }

        const noteReferences = Array.from(rootElement.querySelectorAll('[data-lectr-note-id]'));
        noteReferences.forEach((element) => {
            const noteId = String(element.getAttribute('data-lectr-note-id') || '').trim();
            if (!noteId) {
                return;
            }
            element.classList.add('lectr-note-ref');
            const entry = getNoteEntryById(noteId);
            if (!entry) {
                return;
            }
            const hint = `${entry.title}: ${entry.body}`.trim();
            if (!hint) {
                return;
            }
            element.setAttribute('data-lectr-note-preview', hint);
        });
    };

    let decorateNoteReferencesForPdf = (rootElement) => {
        if (!(rootElement instanceof Element)) {
            return;
        }

        const noteReferences = Array.from(rootElement.querySelectorAll('[data-lectr-note-id]'));
        noteReferences.forEach((element) => {
            const noteId = String(element.getAttribute('data-lectr-note-id') || '').trim();
            const entry = getNoteEntryById(noteId);
            if (!entry) {
                return;
            }

            const noteTarget = element;
            noteTarget.style.background = '#fff8cc';
            noteTarget.style.borderBottom = '1px dashed #9a7f32';
            noteTarget.style.padding = '0 1px';

            const noteBox = document.createElement('div');
            noteBox.setAttribute('class', 'lectr-pdf-note-box');
            noteBox.style.display = 'block';
            noteBox.style.margin = '6px 0 10px';
            noteBox.style.padding = '7px 9px';
            noteBox.style.border = '1px solid #c4cad1';
            noteBox.style.borderLeft = '3px solid #7a8ca0';
            noteBox.style.borderRadius = '6px';
            noteBox.style.background = '#f7f9fb';
            noteBox.style.color = '#1f2b37';
            noteBox.style.fontSize = '12px';
            noteBox.style.lineHeight = '1.45';
            noteBox.innerHTML = `<strong>${escapeHtml(entry.title)}</strong><div>${escapeHtml(entry.body)}</div>`;
            noteTarget.insertAdjacentElement('afterend', noteBox);
        });
    };

    let getElementByIdWithin = (rootElement, idValue) => {
        if (!(rootElement instanceof Element) || !idValue) {
            return null;
        }
        if (typeof CSS !== 'undefined' && typeof CSS.escape === 'function') {
            return rootElement.querySelector(`#${CSS.escape(idValue)}`);
        }
        const escaped = String(idValue).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
        return rootElement.querySelector(`[id="${escaped}"]`);
    };

    let scrollPreviewToAnchor = (rawAnchor) => {
        const anchorId = String(rawAnchor || '').trim().replace(/^#/, '');
        if (!anchorId) {
            return;
        }

        let remainingAttempts = 10;
        const attemptScroll = () => {
            const output = getPreviewOutputElement();
            if (!output) {
                return;
            }
            const target = getElementByIdWithin(output, anchorId);
            if (target) {
                target.scrollIntoView({ block: 'start', behavior: 'smooth' });
                return;
            }
            remainingAttempts -= 1;
            if (remainingAttempts > 0) {
                window.requestAnimationFrame(attemptScroll);
            }
        };

        attemptScroll();
    };

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
            if (!childrenText.trim()) {
                return '';
            }
            return `**${childrenText}**`;
        }
        if (tag === 'em' || tag === 'i') {
            if (!childrenText.trim()) {
                return '';
            }
            return `*${childrenText}*`;
        }
        if (tag === 's' || tag === 'strike' || tag === 'del') {
            if (!childrenText.trim()) {
                return '';
            }
            return `~~${childrenText}~~`;
        }
        if (tag === 'code' && element.parentElement?.tagName.toLowerCase() !== 'pre') {
            if (!childrenText.trim()) {
                return '';
            }
            return `\`${childrenText}\``;
        }
        if (tag === 'a') {
            const href = element.getAttribute('href') || '';
            const label = childrenText || href;
            if (!href && !label.trim()) {
                return '';
            }
            return `[${label}](${href})`;
        }
        if (tag === 'span') {
            const noteId = String(element.getAttribute('data-lectr-note-id') || '').trim();
            if (noteId) {
                return `<span class="lectr-note-ref" data-lectr-note-id="${escapeHtmlAttribute(noteId)}">${childrenText}</span>`;
            }
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
            return text || '';
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

        const markdown = normalizeMarkdownContent(previewOutputToMarkdown());
        const activeTab = getActiveTab();
        if (!activeTab) {
            return;
        }

        syncingFromPreviewEdit = true;
        suppressEditorChange = true;
        editor.setValue(markdown);
        suppressEditorChange = false;
        syncingFromPreviewEdit = false;

        updateTabContentState(activeTab, markdown);
        saveLastContent(markdown);
        lastRenderedMarkdown = markdown;
    };

    let setupPreviewEditSync = (editor) => {
        const output = document.querySelector('#output');
        if (!output) {
            return;
        }

        output.addEventListener('keydown', (event) => {
            if (!previewEditMode) {
                return;
            }
            if (event.key !== 'Tab') {
                return;
            }
            event.preventDefault();
            output.focus();
            restorePreviewSelectionRange();
            const inserted = document.execCommand('insertText', false, '    ');
            if (!inserted) {
                insertHtmlIntoPreviewSelection('    ');
            } else {
                savePreviewSelectionRange();
            }
            refreshFormatToolbarState();
        });

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
                refreshFormatToolbarState();
            }, 220);
        });
    };

    let setupPreviewSelectionTracking = () => {
        document.addEventListener('selectionchange', () => {
            if (!previewEditMode) {
                return;
            }
            savePreviewSelectionRange();
            refreshFormatToolbarState();
        });
    };

    let ensureDomPurifyConfiguration = () => {
        if (isDomPurifyConfigured) {
            return;
        }

        DOMPurify.addHook('afterSanitizeAttributes', (node) => {
            if (!(node instanceof Element) || node.tagName.toLowerCase() !== 'a') {
                return;
            }

            const href = (node.getAttribute('href') || '').trim();
            if (!href || blockedHrefPattern.test(href)) {
                node.removeAttribute('href');
                node.removeAttribute('target');
                node.removeAttribute('rel');
                return;
            }

            if (externalLinkPattern.test(href)) {
                node.setAttribute('target', '_blank');
                node.setAttribute('rel', 'noopener noreferrer');
                return;
            }

            node.removeAttribute('target');
            node.removeAttribute('rel');
        });

        isDomPurifyConfigured = true;
    };

    const sanitizeRenderedMarkdown = (html) => {
        ensureDomPurifyConfiguration();
        return DOMPurify.sanitize(html, {
            USE_PROFILES: { html: true },
            FORBID_TAGS: ['script', 'style', 'iframe', 'object', 'embed', 'form', 'input', 'button', 'textarea', 'select']
        });
    };

    // Render markdown text as html
    let renderMarkdown = (markdown) => {
        let options = {
            headerIds: false,
            mangle: false
        };
        const normalizedMarkdown = normalizeMarkdownContent(markdown);
        let html = marked.parse(normalizedMarkdown, options);
        let sanitized = sanitizeRenderedMarkdown(html);
        const output = document.querySelector('#output');
        if (!output) {
            return;
        }
        output.innerHTML = sanitized;
        applyHeadingIdsToElement(output);
        applyPreviewLinkHints(output);
        applyNoteReferenceDecorations(output);
        previewSavedRange = null;
        setPreviewEditableState(previewEditMode);
        refreshTopImmersionState();
        refreshFormatToolbarState();
        if (previewEditMode) {
            window.requestAnimationFrame(() => {
                ensurePreviewEditableCaret({ preserveSelection: false });
            });
        }
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
            updateTabContentState(activeTab, resetContent);
        }
        presetValue(resetContent);
        document.querySelectorAll('.column').forEach((element) => {
            element.scrollTo({ top: 0 });
        });
        editorIsScrolled = false;
        previewIsScrolled = false;
        updateTopImmersionState();
    };

    let presetValue = (value) => {
        const normalizedValue = normalizeMarkdownContent(value);
        suppressEditorChange = true;
        editor.setValue(normalizedValue);
        suppressEditorChange = false;
        editor.revealPosition({ lineNumber: 1, column: 1 });
        // Force refresh after tab/file open to avoid stale render cache.
        lastRenderedMarkdown = null;
        scheduleConvert(normalizedValue);
        if (!previewEditMode) {
            editor.focus();
            refreshTopImmersionState();
            refreshFormatToolbarState();
            return;
        }
        window.requestAnimationFrame(() => {
            ensurePreviewEditableCaret({ preserveSelection: false });
            refreshTopImmersionState();
            refreshFormatToolbarState();
        });
    };

    // ----- preview CSS loader (switch github-markdown css) -----
    const PREVIEW_CSS_LIGHT = 'css/github-markdown-light.css';
    const PREVIEW_CSS_DARK = 'css/github-markdown-dark_dimmed.css';

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

    let applyThemeState = (enabled, { persist = true } = {}) => {
        const checked = enabled === true;
        const checkbox = document.querySelector('#theme-checkbox');
        if (checkbox) {
            checkbox.checked = checked;
        }

        setTheme(checked);
        setPreviewCss(checked);

        if (monaco && monaco.editor && typeof monaco.editor.setTheme === 'function') {
            monaco.editor.setTheme(checked ? 'vs-dark' : 'vs');
        }

        if (persist) {
            saveThemeSettings(checked);
        }
    };

    applyThemePreference = (enabled, options = {}) => {
        applyThemeState(enabled, options);
    };

    let initThemeToggle = (settings) => {
        let checkbox = document.querySelector('#theme-checkbox');
        if (!checkbox) return;
        applyThemeState(settings, { persist: false });

        if (checkbox.dataset.boundChange === '1') {
            return;
        }
        checkbox.dataset.boundChange = '1';

        checkbox.addEventListener('change', (event) => {
            let checked = event.currentTarget.checked;
            applyThemeState(checked, { persist: true });
        });
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
        updateTabContentState(activeTab, editor.getValue(), { renderOnDirtyChange: false });
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
    let html2PdfLoadPromise = null;

    const ensureHtml2PdfLoaded = () => {
        if (typeof window.html2pdf === 'function') {
            return Promise.resolve(window.html2pdf);
        }

        if (html2PdfLoadPromise) {
            return html2PdfLoadPromise;
        }

        html2PdfLoadPromise = new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js';
            script.crossOrigin = 'anonymous';
            script.referrerPolicy = 'no-referrer';
            script.integrity = 'sha512-GsLlZN/3F2ErC5ifS5QtgpiJtWd43JWSuIgh7mbzZ8zBps+dvLusV+eNQATqgA/HdeKFVgA5v3S/cIrLF7QnIg==';
            script.onload = () => {
                if (typeof window.html2pdf === 'function') {
                    resolve(window.html2pdf);
                    return;
                }
                reject(new Error('html2pdf failed to initialize'));
            };
            script.onerror = () => reject(new Error('Failed to load html2pdf'));
            document.head.appendChild(script);
        }).catch((error) => {
            html2PdfLoadPromise = null;
            throw error;
        });

        return html2PdfLoadPromise;
    };

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

    let rewriteMarkdownLinksForPdf = (rootElement) => {
        if (!(rootElement instanceof Element)) {
            return;
        }

        const anchors = Array.from(rootElement.querySelectorAll('a[href]'));
        anchors.forEach((anchorElement) => {
            const href = String(anchorElement.getAttribute('href') || '').trim();
            if (!href || href.startsWith('#')) {
                return;
            }

            if (/^(?:https?:|mailto:|tel:|data:|blob:|file:)/i.test(href)) {
                return;
            }

            const match = href.match(/^([^?#]+)(\?[^#]*)?(#.*)?$/);
            if (!match) {
                return;
            }

            const pathPart = match[1] || '';
            const queryPart = match[2] || '';
            const hashPart = match[3] || '';
            if (!/\.(?:md|markdown|txt)$/i.test(pathPart)) {
                return;
            }

            const pdfPath = pathPart.replace(/\.(?:md|markdown|txt)$/i, '.pdf');
            anchorElement.setAttribute('href', `${pdfPath}${queryPart}${hashPart}`);
        });
    };

    let exportPreviewToPdf = () => {
        const outputElement = document.querySelector('#output');
        if (!outputElement) {
            return;
        }
        const activeTab = getActiveTab();
        const sourceFilePath = activeTab && typeof activeTab.filePath === 'string' && activeTab.filePath.trim()
            ? activeTab.filePath
            : null;

        const exportHtmlContainer = document.createElement('div');
        exportHtmlContainer.innerHTML = outputElement.innerHTML;
        rewriteMarkdownLinksForPdf(exportHtmlContainer);
        decorateNoteReferencesForPdf(exportHtmlContainer);
        const exportHtml = exportHtmlContainer.innerHTML;

        if (window.lectrDesktop && typeof window.lectrDesktop.exportPreviewPdf === 'function') {
            getLightMarkdownCss().then((lightCss) => {
                return window.lectrDesktop.exportPreviewPdf({
                    html: exportHtml,
                    lightCss,
                    suggestedName: getSuggestedPdfName(),
                    sourceFilePath
                });
            }).catch((error) => {
                // eslint-disable-next-line no-console
                console.error('Failed to export PDF via desktop bridge', error);
            });
            return;
        }

        Promise.all([getLightMarkdownCss(), ensureHtml2PdfLoaded()]).then(([lightCss, html2pdf]) => {
            const options = {
                margin: 10,
                filename: 'markdown-preview.pdf',
                image: { type: 'jpeg', quality: 0.98 },
                enableLinks: true,
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
                            rewriteMarkdownLinksForPdf(clonedOutput);
                            decorateNoteReferencesForPdf(clonedOutput);
                        }
                    }
                },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
            };

            html2pdf()
                .set(options)
                .from(outputElement)
                .save()
                .catch((error) => {
                    // eslint-disable-next-line no-console
                    console.error('Failed to export PDF', error);
                });
        }).catch((error) => {
            // eslint-disable-next-line no-console
            console.error('Failed to prepare PDF export', error);
            window.alert('PDF export is not available yet. Please try again in a moment.');
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

            const dirtyIndicator = document.createElement('span');
            dirtyIndicator.className = 'tab-dirty-indicator';
            dirtyIndicator.textContent = '•';
            dirtyIndicator.setAttribute('aria-hidden', 'true');
            tabElement.appendChild(dirtyIndicator);

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
        captureActiveTabScrollState();
        activeTabId = tabId;
        setTopImmersionFromTab(tab);
        presetValue(tab.content);
        saveLastContent(tab.content);
        renderTabs();
        window.requestAnimationFrame(() => {
            applyTabScrollState(tab);
        });
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

    let animateTabClose = (tabElement) => {
        return new Promise((resolve) => {
            if (!(tabElement instanceof HTMLElement)) {
                resolve();
                return;
            }

            const rect = tabElement.getBoundingClientRect();
            tabElement.style.width = `${rect.width}px`;
            tabElement.style.minWidth = `${rect.width}px`;
            tabElement.style.maxWidth = `${rect.width}px`;
            tabElement.classList.add('tab-exit');

            let completed = false;
            const done = () => {
                if (completed) {
                    return;
                }
                completed = true;
                resolve();
            };

            window.requestAnimationFrame(() => {
                tabElement.classList.add('tab-exit-active');
            });

            tabElement.addEventListener('transitionend', done, { once: true });
            window.setTimeout(done, 240);
        });
    };

    let closeTab = async (tabId, editor, options = {}) => {
        const { tabElement = null, animate = false } = options;
        const tab = tabs.find((item) => item.id === tabId);
        if (!tab) {
            return false;
        }

        if (tab.id === activeTabId && editor) {
            updateTabContentState(tab, editor.getValue(), { renderOnDirtyChange: false });
        }

        if (tab.dirty) {
            const shouldSaveAndClose = window.confirm(t('confirmCloseUnsavedSave', { title: tab.title }));
            if (shouldSaveAndClose) {
                const saveResult = await saveTab(tab, {
                    forceDialog: !tab.filePath,
                    showToast: true
                });
                if (!saveResult || !saveResult.saved) {
                    return false;
                }
            } else {
                const discardConfirmed = window.confirm(t('confirmCloseUnsavedDiscard', { title: tab.title }));
                if (!discardConfirmed) {
                    return false;
                }
            }
        }

        if (animate && tabElement instanceof HTMLElement) {
            await animateTabClose(tabElement);
        }

        const closingIndex = tabs.findIndex((item) => item.id === tabId);
        tabs.splice(closingIndex, 1);

        if (tabs.length === 0) {
            createUntitledTab(getDefaultInput());
            return true;
        }

        const nextIndex = Math.max(0, closingIndex - 1);
        const nextTab = tabs[nextIndex];
        if (activeTabId === tabId && nextTab) {
            activateTab(nextTab.id);
        } else {
            renderTabs();
        }
        return true;
    };

    let setupTabsUi = (editor) => {
        const tabsList = document.querySelector('#tabs-list');
        const tabsBar = document.querySelector('#tabs-bar');
        const dragState = {
            pointerId: null,
            tabId: null,
            tabElement: null,
            originIndex: null,
            startClientX: 0,
            startClientY: 0,
            pointerOffsetX: 0,
            started: false,
            activated: false,
            suppressNextClick: false,
            pendingInsertIndex: null,
            slotSize: 0,
            originGap: null,
            floatingLeft: 0,
            floatingTop: 0
        };

        const clearDropSlotPreview = () => {
            if (!tabsList) {
                return;
            }
            tabsList.style.paddingRight = '';
            tabsList.querySelectorAll('.tab-item').forEach((element) => {
                element.classList.remove('drop-target');
                element.style.marginLeft = '';
            });
        };

        const clearFloatingStyles = (element) => {
            if (!element) {
                return;
            }
            element.style.position = '';
            element.style.left = '';
            element.style.top = '';
            element.style.width = '';
            element.style.height = '';
            element.style.margin = '';
            element.style.transform = '';
            element.style.zIndex = '';
            element.style.pointerEvents = '';
        };

        const captureTabPositions = () => {
            const positions = new Map();
            if (!tabsList) {
                return positions;
            }
            tabsList.querySelectorAll('.tab-item').forEach((element) => {
                const tabId = element.getAttribute('data-tab-id');
                if (!tabId) {
                    return;
                }
                const rect = element.getBoundingClientRect();
                positions.set(tabId, { left: rect.left });
            });
            return positions;
        };

        const animateTabPositions = (beforePositions) => {
            if (!tabsList || !(beforePositions instanceof Map) || beforePositions.size === 0) {
                return;
            }
            tabsList.querySelectorAll('.tab-item').forEach((element) => {
                const tabId = element.getAttribute('data-tab-id');
                if (!tabId || !beforePositions.has(tabId)) {
                    return;
                }

                const previous = beforePositions.get(tabId);
                const rect = element.getBoundingClientRect();
                const deltaX = previous.left - rect.left;
                if (Math.abs(deltaX) < 0.5) {
                    return;
                }

                element.classList.add('reorder-anim');
                element.style.transform = `translateX(${deltaX}px)`;
                void element.offsetWidth;
                element.style.transform = 'translateX(0)';
                element.addEventListener('transitionend', () => {
                    element.classList.remove('reorder-anim');
                    element.style.transform = '';
                }, { once: true });
            });
        };

        const getInsertIndexByCenter = (draggedCenter) => {
            if (!tabsList || !dragState.tabElement || !dragState.tabId) {
                return null;
            }

            const items = Array.from(tabsList.querySelectorAll('.tab-item'));
            const dragged = dragState.tabElement;
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

            return insertIndex;
        };

        const applyDropSlotPreview = (insertIndex) => {
            if (!tabsList || !dragState.tabElement || !dragState.tabId || insertIndex === null) {
                return;
            }
            const withoutDraggedCount = Math.max(0, tabs.length - 1);
            const boundedInsertIndex = Math.max(0, Math.min(withoutDraggedCount, insertIndex));
            if (dragState.pendingInsertIndex === boundedInsertIndex) {
                return;
            }
            const beforePositions = captureTabPositions();
            dragState.pendingInsertIndex = boundedInsertIndex;
            clearDropSlotPreview();

            const currentIndex = tabs.findIndex((tab) => tab.id === dragState.tabId);
            if (currentIndex === -1) {
                animateTabPositions(beforePositions);
                return;
            }

            const slotSize = Math.max(32, dragState.slotSize);
            const withoutDragged = Array.from(tabsList.querySelectorAll('.tab-item')).filter((element) => element !== dragState.tabElement);
            if (boundedInsertIndex >= withoutDragged.length) {
                tabsList.style.paddingRight = '0px';
                void tabsList.offsetWidth;
                tabsList.style.paddingRight = `${slotSize}px`;
                animateTabPositions(beforePositions);
                return;
            }
            const targetElement = withoutDragged[boundedInsertIndex];
            if (targetElement) {
                targetElement.classList.add('drop-target');
                targetElement.style.marginLeft = '0px';
                void targetElement.offsetWidth;
                targetElement.style.marginLeft = `${slotSize}px`;
            }
            animateTabPositions(beforePositions);
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
            dragState.startClientX = event.clientX;
            dragState.startClientY = event.clientY;
            dragState.pointerId = event.pointerId;
            dragState.tabId = tabId;
            dragState.tabElement = tabElement;
            dragState.originIndex = tabs.findIndex((tab) => tab.id === tabId);
            dragState.started = true;
            dragState.activated = false;
            dragState.pendingInsertIndex = dragState.originIndex;
            dragState.slotSize = Math.round(rect.width + 10);
            dragState.floatingLeft = rect.left;
            dragState.floatingTop = rect.top;

            tabElement.setPointerCapture(event.pointerId);
        };

        const updateDrag = (event) => {
            if (!tabsList || !dragState.started || dragState.pointerId !== event.pointerId || !dragState.tabElement) {
                return;
            }

            const dragged = dragState.tabElement;
            if (!dragState.activated) {
                const deltaX = Math.abs(event.clientX - dragState.startClientX);
                const deltaY = Math.abs(event.clientY - dragState.startClientY);
                if (deltaX < 4 && deltaY < 4) {
                    return;
                }

                const beforePositions = captureTabPositions();
                const rect = dragged.getBoundingClientRect();
                dragState.pointerOffsetX = event.clientX - rect.left;
                dragState.slotSize = Math.round(rect.width + 10);
                dragState.floatingTop = rect.top;
                dragState.floatingLeft = rect.left;
                dragState.activated = true;

                const originGap = document.createElement('div');
                originGap.className = 'tab-origin-gap';
                originGap.style.width = `${Math.round(rect.width)}px`;
                originGap.style.minWidth = `${Math.round(rect.width)}px`;
                originGap.style.maxWidth = `${Math.round(rect.width)}px`;
                if (tabsList) {
                    tabsList.insertBefore(originGap, dragged.nextSibling);
                    dragState.originGap = originGap;
                    window.requestAnimationFrame(() => {
                        if (!dragState.started || dragState.originGap !== originGap) {
                            return;
                        }
                        originGap.classList.add('visible');
                        void originGap.offsetWidth;
                        originGap.classList.add('collapse');
                    });
                    originGap.addEventListener('transitionend', () => {
                        if (originGap.parentElement) {
                            originGap.remove();
                        }
                        if (dragState.originGap === originGap) {
                            dragState.originGap = null;
                        }
                    }, { once: true });
                }

                dragged.classList.add('dragging-pointer');
                dragged.style.position = 'fixed';
                dragged.style.left = `${rect.left}px`;
                dragged.style.top = `${rect.top}px`;
                dragged.style.width = `${rect.width}px`;
                dragged.style.height = `${rect.height}px`;
                dragged.style.margin = '0';
                dragged.style.zIndex = '140';
                dragged.style.pointerEvents = 'none';
                document.body.appendChild(dragged);

                document.body.classList.add('tabs-dragging');
                if (tabsBar) {
                    tabsBar.classList.add('drag-slot-active');
                }
                animateTabPositions(beforePositions);
            }

            dragState.floatingLeft = event.clientX - dragState.pointerOffsetX;
            dragged.style.left = `${dragState.floatingLeft}px`;

            const draggedCenter = dragState.floatingLeft + dragState.slotSize / 2;
            const insertIndex = getInsertIndexByCenter(draggedCenter);
            applyDropSlotPreview(insertIndex);
        };

        const resetDragState = () => {
            if (dragState.originGap && dragState.originGap.parentElement) {
                dragState.originGap.remove();
            }
            dragState.originGap = null;
            dragState.pointerId = null;
            dragState.tabId = null;
            dragState.tabElement = null;
            dragState.originIndex = null;
            dragState.startClientX = 0;
            dragState.startClientY = 0;
            dragState.pointerOffsetX = 0;
            dragState.started = false;
            dragState.activated = false;
            dragState.pendingInsertIndex = null;
            dragState.slotSize = 0;
            dragState.floatingLeft = 0;
            dragState.floatingTop = 0;
            document.body.classList.remove('tabs-dragging');
        };

        const endDrag = (event = null, { commit = true } = {}) => {
            if (!dragState.started) {
                return;
            }
            if (event && dragState.pointerId !== event.pointerId) {
                return;
            }

            const wasActivated = dragState.activated;
            const draggedTabId = dragState.tabId;
            const proposedInsertIndex = dragState.pendingInsertIndex;
            const beforePositions = wasActivated ? captureTabPositions() : new Map();
            const floatingInBody = !!(dragState.tabElement && dragState.tabElement.parentElement === document.body);

            if (dragState.tabElement) {
                dragState.tabElement.classList.remove('dragging-pointer');
                if (Number.isInteger(dragState.pointerId) && dragState.tabElement.hasPointerCapture(dragState.pointerId)) {
                    dragState.tabElement.releasePointerCapture(dragState.pointerId);
                }
                if (floatingInBody) {
                    dragState.tabElement.remove();
                } else {
                    clearFloatingStyles(dragState.tabElement);
                }
            }

            if (tabsBar) {
                tabsBar.classList.remove('drag-slot-active');
            }

            if (commit && wasActivated && draggedTabId && Number.isInteger(proposedInsertIndex)) {
                const currentIndex = tabs.findIndex((tab) => tab.id === draggedTabId);
                const boundedTargetIndex = Math.max(0, Math.min(tabs.length - 1, proposedInsertIndex));
                if (currentIndex !== -1) {
                    moveTab(currentIndex, boundedTargetIndex);
                }
                renderTabs();
                clearDropSlotPreview();
                animateTabPositions(beforePositions);
                dragState.suppressNextClick = true;
            } else {
                clearDropSlotPreview();
            }

            resetDragState();
        };

        const cancelDrag = () => {
            endDrag(null, { commit: false });
        };

        if (tabsList) {
            tabsList.addEventListener('click', async (event) => {
                if (dragState.suppressNextClick) {
                    dragState.suppressNextClick = false;
                    event.preventDefault();
                    event.stopPropagation();
                    return;
                }

                const closeTarget = event.target.closest('[data-tab-close]');
                if (closeTarget) {
                    event.preventDefault();
                    event.stopPropagation();
                    const tabId = closeTarget.getAttribute('data-tab-close');
                    if (tabId) {
                        const parentTabElement = closeTarget.closest('[data-tab-id]');
                        await closeTab(tabId, editor, {
                            tabElement: parentTabElement,
                            animate: true
                        });
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

            tabsList.addEventListener('dragstart', (event) => {
                event.preventDefault();
            });
        }

        document.addEventListener('pointermove', (event) => {
            updateDrag(event);
        });

        document.addEventListener('pointerup', (event) => {
            endDrag(event);
        });

        document.addEventListener('pointercancel', (event) => {
            endDrag(event);
        });

        window.addEventListener('blur', () => {
            cancelDrag();
        });

        document.addEventListener('visibilitychange', () => {
            if (document.hidden) {
                cancelDrag();
            }
        });

        const newTabButton = document.querySelector('#new-tab-button');
        if (newTabButton) {
            newTabButton.addEventListener('click', () => {
                const beforePositions = captureTabPositions();
                const createdTab = createUntitledTab('');
                animateTabPositions(beforePositions);
                requestAnimationFrame(() => {
                    if (!createdTab || !createdTab.id) {
                        return;
                    }
                    const createdElement = tabsList ? tabsList.querySelector(`[data-tab-id="${createdTab.id}"]`) : null;
                    if (!createdElement) {
                        return;
                    }
                    createdElement.classList.add('tab-enter');
                    createdElement.addEventListener('animationend', () => {
                        createdElement.classList.remove('tab-enter');
                    }, { once: true });
                });
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

    let setupAboutDialog = () => {
        const overlay = document.querySelector('#about-overlay');
        const brandButton = document.querySelector('#brand-button');
        const closeButton = document.querySelector('#about-close-button');
        const versionValue = document.querySelector('#about-version-value');
        if (!overlay || !brandButton || !closeButton || !versionValue) {
            return;
        }

        const fallbackVersion = versionValue.textContent || 'v1.0.1';
        let appVersionLoaded = false;

        const isOpen = () => !overlay.hidden && overlay.classList.contains('open');

        const closeAbout = () => {
            overlay.classList.remove('open');
            overlay.setAttribute('aria-hidden', 'true');
            document.body.classList.remove('about-open');
            window.setTimeout(() => {
                overlay.hidden = true;
            }, 210);
        };

        const ensureAppVersion = async () => {
            if (appVersionLoaded) {
                return;
            }
            appVersionLoaded = true;
            if (!window.lectrDesktop || typeof window.lectrDesktop.getAppVersion !== 'function') {
                versionValue.textContent = fallbackVersion;
                return;
            }
            try {
                const version = await window.lectrDesktop.getAppVersion();
                if (typeof version === 'string' && version.trim()) {
                    versionValue.textContent = `v${version.trim()}`;
                    return;
                }
            } catch (error) {
                // eslint-disable-next-line no-console
                console.error('Failed to load app version', error);
            }
            versionValue.textContent = fallbackVersion;
        };

        const openAbout = () => {
            closeOpenMenus();
            overlay.hidden = false;
            overlay.setAttribute('aria-hidden', 'false');
            document.body.classList.add('about-open');
            window.requestAnimationFrame(() => {
                overlay.classList.add('open');
            });
            void ensureAppVersion();
        };

        brandButton.addEventListener('click', (event) => {
            event.preventDefault();
            if (isOpen()) {
                closeAbout();
                return;
            }
            openAbout();
        });

        closeButton.addEventListener('click', (event) => {
            event.preventDefault();
            closeAbout();
        });

        overlay.addEventListener('pointerdown', (event) => {
            if (event.target === overlay) {
                closeAbout();
            }
        });

        document.addEventListener('keydown', (event) => {
            if (event.key === 'Escape' && isOpen()) {
                closeAbout();
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
        if (previewEditMode) {
            syncEditorFromPreview(editor);
        }
        syncActiveTabFromEditor(editor);
        const result = await saveTab(activeTab, { forceDialog, showToast: true });

        if (previewEditMode) {
            window.requestAnimationFrame(() => {
                ensurePreviewEditableCaret({ preserveSelection: true });
                refreshFormatToolbarState();
            });
        }

        return result;
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
            const openByModifier = event.ctrlKey || event.metaKey;
            if (previewEditMode && !openByModifier) {
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
                event.preventDefault();
                scrollPreviewToAnchor(href.slice(1));
                return;
            }

            const hrefMatch = href.match(/^([^#]*)(#(.*))?$/);
            let linkAnchor = '';
            if (hrefMatch && typeof hrefMatch[3] === 'string') {
                const rawAnchor = hrefMatch[3].trim();
                if (rawAnchor) {
                    try {
                        linkAnchor = decodeURIComponent(rawAnchor);
                    } catch (_error) {
                        linkAnchor = rawAnchor;
                    }
                }
            }

            if (externalLinkPattern.test(href) || /^(?:data:|blob:)/i.test(href)) {
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
                return;
            }

            if (linkAnchor) {
                scrollPreviewToAnchor(linkAnchor);
            }
        });
    };

    let setupPreviewNoteInteractions = () => {
        const output = document.querySelector('#output');
        if (!output) {
            return;
        }

        const getNoteElement = (target) => {
            if (!(target instanceof Element)) {
                return null;
            }
            return target.closest('[data-lectr-note-id]');
        };

        output.addEventListener('mouseover', (event) => {
            const noteElement = getNoteElement(event.target);
            if (!noteElement) {
                return;
            }
            showNoteTooltip(noteElement);
        });

        output.addEventListener('mousemove', (event) => {
            const noteElement = getNoteElement(event.target);
            if (!noteElement) {
                if (activeNoteTooltipRef) {
                    hideNoteTooltip();
                }
                return;
            }
            if (activeNoteTooltipRef !== noteElement) {
                showNoteTooltip(noteElement);
            }
        });

        output.addEventListener('mouseleave', () => {
            hideNoteTooltip();
        });

        output.addEventListener('scroll', () => {
            hideNoteTooltip();
        });
    };

    let setupPreviewTableActions = (editor) => {
        const output = document.querySelector('#output');
        if (!output) {
            return;
        }

        const floatingButton = document.createElement('button');
        floatingButton.type = 'button';
        floatingButton.className = 'preview-table-action-button';
        floatingButton.setAttribute('aria-label', t('tablePopoverTitle'));
        floatingButton.setAttribute('title', t('tablePopoverTitle'));
        floatingButton.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 10a2 2 0 1 0 0 .01V10zm6 0a2 2 0 1 0 0 .01V10zm6 0a2 2 0 1 0 0 .01V10z"/></svg>';
        floatingButton.hidden = true;

        const actionMenu = document.createElement('div');
        actionMenu.className = 'preview-table-action-menu';
        actionMenu.hidden = true;

        const createMenuButton = (action, labelKey, danger = false) => {
            const button = document.createElement('button');
            button.type = 'button';
            button.className = danger ? 'danger' : '';
            button.setAttribute('data-table-action', action);
            button.textContent = t(labelKey);
            actionMenu.appendChild(button);
        };

        createMenuButton('add-column', 'tableAddColumn');
        createMenuButton('remove-column', 'tableRemoveColumn');
        createMenuButton('add-row', 'tableAddRow');
        createMenuButton('remove-row', 'tableRemoveRow');
        createMenuButton('delete-table', 'tableDeleteTable', true);

        document.body.appendChild(floatingButton);
        document.body.appendChild(actionMenu);

        let activeTable = null;
        let menuOpen = false;

        const getTableFromTarget = (target) => {
            if (!(target instanceof Element)) {
                return null;
            }
            const table = target.closest('table');
            if (!(table instanceof HTMLTableElement) || !output.contains(table)) {
                return null;
            }
            return table;
        };

        const getTableFromSelection = () => {
            const selection = window.getSelection();
            if (!selection || selection.rangeCount === 0) {
                return null;
            }
            const anchorNode = selection.anchorNode;
            if (!anchorNode) {
                return null;
            }
            const element = anchorNode.nodeType === Node.ELEMENT_NODE
                ? anchorNode
                : anchorNode.parentElement;
            if (!(element instanceof Element)) {
                return null;
            }
            return getTableFromTarget(element);
        };

        const updateMenuLabels = () => {
            floatingButton.setAttribute('aria-label', t('tablePopoverTitle'));
            floatingButton.setAttribute('title', t('tablePopoverTitle'));
            const buttons = Array.from(actionMenu.querySelectorAll('[data-table-action]'));
            buttons.forEach((button) => {
                const action = button.getAttribute('data-table-action');
                if (!action) {
                    return;
                }
                if (action === 'add-column') {
                    button.textContent = t('tableAddColumn');
                } else if (action === 'remove-column') {
                    button.textContent = t('tableRemoveColumn');
                } else if (action === 'add-row') {
                    button.textContent = t('tableAddRow');
                } else if (action === 'remove-row') {
                    button.textContent = t('tableRemoveRow');
                } else if (action === 'delete-table') {
                    button.textContent = t('tableDeleteTable');
                }
            });
        };

        const hideActionMenu = () => {
            menuOpen = false;
            actionMenu.hidden = true;
        };

        const clearActiveTable = () => {
            activeTable = null;
            floatingButton.hidden = true;
            hideActionMenu();
        };

        const positionActionButton = () => {
            if (!previewEditMode || !(activeTable instanceof HTMLTableElement) || !activeTable.isConnected || !output.contains(activeTable)) {
                clearActiveTable();
                return;
            }

            const rect = activeTable.getBoundingClientRect();
            if (rect.width <= 0 || rect.height <= 0 || rect.bottom < 0 || rect.top > window.innerHeight) {
                floatingButton.hidden = true;
                hideActionMenu();
                return;
            }

            const buttonSize = 30;
            const left = Math.round(Math.max(8, Math.min(window.innerWidth - buttonSize - 8, rect.right - buttonSize - 4)));
            const top = Math.round(Math.max(8, Math.min(window.innerHeight - buttonSize - 8, rect.bottom - buttonSize - 4)));
            floatingButton.style.left = `${left}px`;
            floatingButton.style.top = `${top}px`;
            floatingButton.hidden = false;
        };

        const positionActionMenu = () => {
            if (!menuOpen || floatingButton.hidden) {
                return;
            }
            const buttonRect = floatingButton.getBoundingClientRect();
            const menuRect = actionMenu.getBoundingClientRect();
            const desiredLeft = Math.round(buttonRect.right - menuRect.width);
            const left = Math.max(8, Math.min(window.innerWidth - menuRect.width - 8, desiredLeft));
            const aboveTop = Math.round(buttonRect.top - menuRect.height - 8);
            const top = aboveTop >= 8
                ? aboveTop
                : Math.round(Math.min(window.innerHeight - menuRect.height - 8, buttonRect.bottom + 8));
            actionMenu.style.left = `${left}px`;
            actionMenu.style.top = `${top}px`;
        };

        const showActionMenu = () => {
            if (floatingButton.hidden) {
                return;
            }
            updateMenuLabels();
            menuOpen = true;
            actionMenu.hidden = false;
            positionActionMenu();
        };

        const setActiveTable = (table) => {
            if (!(table instanceof HTMLTableElement)) {
                clearActiveTable();
                return;
            }
            if (activeTable !== table) {
                hideActionMenu();
            }
            activeTable = table;
            positionActionButton();
        };

        const ensureCaretInsideActiveTable = () => {
            if (!(activeTable instanceof HTMLTableElement) || !activeTable.isConnected) {
                return false;
            }
            const outputElement = getPreviewOutputElement();
            if (!outputElement) {
                return false;
            }
            outputElement.focus();
            const selectionTable = getTableFromSelection();
            if (selectionTable === activeTable) {
                savePreviewSelectionRange();
                return true;
            }
            const firstCell = activeTable.querySelector('th,td');
            if (!(firstCell instanceof Element)) {
                return false;
            }
            movePreviewCaretToCell(firstCell);
            return true;
        };

        const applyTableActionResult = (result) => {
            if (result && result.ok) {
                syncEditorFromPreview(editor);
                refreshFormatToolbarState();
                positionActionButton();
                if (menuOpen) {
                    positionActionMenu();
                }
                return;
            }
            if (result && result.reason === 'min_columns') {
                window.alert(t('tableMinColumns'));
                return;
            }
            if (result && result.reason === 'no_data_rows') {
                window.alert(t('tableNoDataRows'));
                return;
            }
            window.alert(t('tableNotFound'));
        };

        output.addEventListener('mousemove', (event) => {
            if (!previewEditMode) {
                clearActiveTable();
                return;
            }
            const table = getTableFromTarget(event.target);
            if (table) {
                setActiveTable(table);
                return;
            }
            const selectedTable = getTableFromSelection();
            if (selectedTable) {
                setActiveTable(selectedTable);
                return;
            }
            if (!menuOpen) {
                clearActiveTable();
            }
        });

        output.addEventListener('click', (event) => {
            if (!previewEditMode) {
                clearActiveTable();
                return;
            }
            const table = getTableFromTarget(event.target);
            if (table) {
                setActiveTable(table);
            }
        });

        output.addEventListener('scroll', () => {
            if (!previewEditMode) {
                clearActiveTable();
                return;
            }
            positionActionButton();
            positionActionMenu();
        });

        document.addEventListener('selectionchange', () => {
            if (!previewEditMode) {
                clearActiveTable();
                return;
            }
            const table = getTableFromSelection();
            if (table) {
                setActiveTable(table);
            } else if (!menuOpen) {
                clearActiveTable();
            }
        });

        floatingButton.addEventListener('click', (event) => {
            event.preventDefault();
            event.stopPropagation();
            if (menuOpen) {
                hideActionMenu();
                return;
            }
            showActionMenu();
        });

        actionMenu.addEventListener('click', (event) => {
            const target = event.target;
            if (!(target instanceof HTMLButtonElement)) {
                return;
            }
            const action = String(target.getAttribute('data-table-action') || '').trim();
            if (!action || !(activeTable instanceof HTMLTableElement) || !activeTable.isConnected) {
                return;
            }
            event.preventDefault();
            event.stopPropagation();

            if (action === 'delete-table') {
                activeTable.remove();
                hideActionMenu();
                clearActiveTable();
                syncEditorFromPreview(editor);
                refreshFormatToolbarState();
                return;
            }

            if (!ensureCaretInsideActiveTable()) {
                window.alert(t('tableNotFound'));
                return;
            }

            if (action === 'add-column') {
                applyTableActionResult(modifyPreviewTableColumn('add'));
                return;
            }
            if (action === 'remove-column') {
                applyTableActionResult(modifyPreviewTableColumn('remove'));
                return;
            }
            if (action === 'add-row') {
                applyTableActionResult(modifyPreviewTableRow('add'));
                return;
            }
            if (action === 'remove-row') {
                applyTableActionResult(modifyPreviewTableRow('remove'));
            }
        });

        document.addEventListener('pointerdown', (event) => {
            const target = event.target;
            if (!(target instanceof Element)) {
                return;
            }
            if (target.closest('.preview-table-action-button') || target.closest('.preview-table-action-menu')) {
                return;
            }
            if (activeTable && target.closest('table') === activeTable) {
                hideActionMenu();
                return;
            }
            clearActiveTable();
        });

        window.addEventListener('resize', () => {
            positionActionButton();
            positionActionMenu();
        });

        document.addEventListener('keydown', (event) => {
            if (event.key === 'Escape') {
                hideActionMenu();
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
                    if (previewFormatType === 'link') {
                        const linkToolButton = document.querySelector('#link-tool-button');
                        if (linkToolButton instanceof HTMLButtonElement) {
                            linkToolButton.click();
                            return;
                        }
                    }
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
                    const activeTabElement = document.querySelector(`#tabs-list [data-tab-id="${activeTab.id}"]`);
                    await closeTab(activeTab.id, editor, {
                        tabElement: activeTabElement,
                        animate: true
                    });
                }
                return;
            }

            if (key === 't' && !event.shiftKey) {
                if (editableField) {
                    return;
                }
                event.preventDefault();
                const newTabButton = document.querySelector('#new-tab-button');
                if (newTabButton instanceof HTMLButtonElement) {
                    newTabButton.click();
                } else {
                    createUntitledTab('');
                }
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

    let buildLinkHref = (target, anchor) => {
        const cleanTarget = String(target || '').trim();
        const cleanAnchor = String(anchor || '').trim().replace(/^#+/, '');
        if (cleanAnchor) {
            if (cleanTarget) {
                const withoutAnchor = cleanTarget.replace(/#.*$/, '');
                return `${withoutAnchor}#${cleanAnchor}`;
            }
            return `#${cleanAnchor}`;
        }
        return cleanTarget;
    };

    let splitLinkHref = (href) => {
        const value = String(href || '').trim();
        if (!value) {
            return { target: '', anchor: '' };
        }
        if (value.startsWith('#')) {
            return {
                target: '',
                anchor: value.slice(1)
            };
        }
        const hashIndex = value.lastIndexOf('#');
        if (hashIndex <= 0) {
            return { target: value, anchor: '' };
        }
        return {
            target: value.slice(0, hashIndex),
            anchor: value.slice(hashIndex + 1)
        };
    };

    let applyLinkFormatWithValues = (editor, { label, href, replaceRange = null }) => {
        const safeLabel = String(label || '').trim() || 'link text';
        const safeHref = String(href || '').trim();
        const markdownHref = safeHref.replace(/\s/g, '%20');
        const snippet = `[${safeLabel}](${markdownHref})`;
        const urlStart = safeLabel.length + 3;
        const urlEnd = urlStart + markdownHref.length;
        if (replaceRange) {
            replaceMarkdownRangeWithSnippet(editor, replaceRange, snippet, { start: urlStart, end: urlEnd });
            return;
        }
        insertMarkdownSnippet(editor, snippet, { start: urlStart, end: urlEnd });
    };

    let applyImageFormatWithValues = (editor, { alt, src, title }) => {
        const safeAlt = String(alt || '').trim() || 'image';
        const safeSrc = String(src || '').trim().replace(/\s/g, '%20');
        const safeTitle = String(title || '').trim().replace(/"/g, '\\"');
        const titlePart = safeTitle ? ` "${safeTitle}"` : '';
        const snippet = `![${safeAlt}](${safeSrc}${titlePart})`;
        const srcStart = safeAlt.length + 4;
        const srcEnd = srcStart + safeSrc.length;
        insertMarkdownSnippet(editor, snippet, { start: srcStart, end: srcEnd });
    };

    let parseIntegerInput = (value, min, max, fallback) => {
        const parsed = Number.parseInt(String(value), 10);
        if (!Number.isFinite(parsed)) {
            return fallback;
        }
        return Math.min(max, Math.max(min, parsed));
    };

    let getTableSeparatorCell = (alignment) => {
        if (alignment === 'center') {
            return ':---:';
        }
        if (alignment === 'right') {
            return '---:';
        }
        return ':---';
    };

    let buildMarkdownTable = ({ columns, rows, alignment, includeHeader, headerPrefix }) => {
        const safeColumns = parseIntegerInput(columns, 1, 12, 3);
        const safeRows = parseIntegerInput(rows, 1, 50, 3);
        const safePrefix = String(headerPrefix || 'Column').trim() || 'Column';
        const headerCells = includeHeader
            ? Array.from({ length: safeColumns }, (_, index) => `${safePrefix} ${index + 1}`)
            : new Array(safeColumns).fill('');
        const separatorCells = new Array(safeColumns).fill(getTableSeparatorCell(alignment));
        const bodyRows = Array.from({ length: safeRows }, () => new Array(safeColumns).fill(''));

        const asRow = (cells) => `| ${cells.join(' | ')} |`;
        const lines = [asRow(headerCells), asRow(separatorCells)];
        bodyRows.forEach((cells) => {
            lines.push(asRow(cells));
        });
        return lines.join('\n');
    };

    let insertMarkdownSnippet = (editor, snippet, selectionInSnippet = null) => {
        const model = editor.getModel();
        const selection = editor.getSelection();
        if (!model || !selection || typeof snippet !== 'string') {
            return;
        }

        const startOffset = model.getOffsetAt({
            lineNumber: selection.startLineNumber,
            column: selection.startColumn
        });
        editor.executeEdits('format-toolbar', [{
            range: selection,
            text: snippet,
            forceMoveMarkers: true
        }]);

        if (selectionInSnippet && Number.isFinite(selectionInSnippet.start) && Number.isFinite(selectionInSnippet.end)) {
            const safeStart = Math.max(0, Math.min(snippet.length, selectionInSnippet.start));
            const safeEnd = Math.max(safeStart, Math.min(snippet.length, selectionInSnippet.end));
            const selectionStart = model.getPositionAt(startOffset + safeStart);
            const selectionEnd = model.getPositionAt(startOffset + safeEnd);
            editor.setSelection(new monaco.Selection(
                selectionStart.lineNumber,
                selectionStart.column,
                selectionEnd.lineNumber,
                selectionEnd.column
            ));
            editor.focus();
            return;
        }

        const caretOffset = startOffset + snippet.length;
        const caretPosition = model.getPositionAt(caretOffset);
        editor.setSelection(new monaco.Selection(
            caretPosition.lineNumber,
            caretPosition.column,
            caretPosition.lineNumber,
            caretPosition.column
        ));
        editor.focus();
    };

    let replaceMarkdownRangeWithSnippet = (editor, range, snippet, selectionInSnippet = null) => {
        const model = editor.getModel();
        if (!model || !range || typeof snippet !== 'string') {
            return;
        }

        const startOffset = model.getOffsetAt({
            lineNumber: range.startLineNumber,
            column: range.startColumn
        });
        editor.executeEdits('format-toolbar', [{
            range,
            text: snippet,
            forceMoveMarkers: true
        }]);

        if (selectionInSnippet && Number.isFinite(selectionInSnippet.start) && Number.isFinite(selectionInSnippet.end)) {
            const safeStart = Math.max(0, Math.min(snippet.length, selectionInSnippet.start));
            const safeEnd = Math.max(safeStart, Math.min(snippet.length, selectionInSnippet.end));
            const selectionStart = model.getPositionAt(startOffset + safeStart);
            const selectionEnd = model.getPositionAt(startOffset + safeEnd);
            editor.setSelection(new monaco.Selection(
                selectionStart.lineNumber,
                selectionStart.column,
                selectionEnd.lineNumber,
                selectionEnd.column
            ));
            editor.focus();
            return;
        }

        const caretOffset = startOffset + snippet.length;
        const caretPosition = model.getPositionAt(caretOffset);
        editor.setSelection(new monaco.Selection(
            caretPosition.lineNumber,
            caretPosition.column,
            caretPosition.lineNumber,
            caretPosition.column
        ));
        editor.focus();
    };

    let splitMarkdownTableRow = (line) => {
        const raw = String(line || '').trim().replace(/^\|/, '').replace(/\|$/, '');
        const cells = [];
        let cell = '';
        let escaped = false;
        for (let index = 0; index < raw.length; index += 1) {
            const char = raw[index];
            if (escaped) {
                cell += char;
                escaped = false;
                continue;
            }
            if (char === '\\') {
                cell += char;
                escaped = true;
                continue;
            }
            if (char === '|') {
                cells.push(cell.trim());
                cell = '';
                continue;
            }
            cell += char;
        }
        cells.push(cell.trim());
        return cells;
    };

    let isMarkdownTableSeparatorRow = (cells) => {
        if (!Array.isArray(cells) || cells.length === 0) {
            return false;
        }
        return cells.every((cell) => /^:?-{3,}:?$/.test(String(cell || '').trim()));
    };

    let countUnescapedPipesBeforeOffset = (line, offsetExclusive) => {
        const safeLine = String(line || '');
        const safeOffset = Math.max(0, Math.min(safeLine.length, offsetExclusive));
        let escaped = false;
        let count = 0;
        for (let index = 0; index < safeOffset; index += 1) {
            const char = safeLine[index];
            if (escaped) {
                escaped = false;
                continue;
            }
            if (char === '\\') {
                escaped = true;
                continue;
            }
            if (char === '|') {
                count += 1;
            }
        }
        return count;
    };

    let hasUnescapedPipe = (line) => countUnescapedPipesBeforeOffset(line, String(line || '').length) > 0;

    let buildMarkdownTableFromRows = (rows) => {
        const safeRows = Array.isArray(rows) ? rows : [];
        const asRow = (cells) => `| ${cells.join(' | ')} |`;
        return safeRows.map((cells) => asRow(cells)).join('\n');
    };

    let getCellStartColumnInRenderedRow = (cells, targetCellIndex) => {
        const safeCells = Array.isArray(cells) ? cells : [];
        const safeIndex = Math.max(0, Math.min(safeCells.length - 1, targetCellIndex));
        let offset = 2;
        for (let index = 0; index < safeIndex; index += 1) {
            offset += String(safeCells[index] || '').length + 3;
        }
        return offset + 1;
    };

    let findMarkdownTableAtSelection = (editor) => {
        const model = editor.getModel();
        const selection = editor.getSelection();
        if (!model || !selection) {
            return null;
        }

        const currentLineNumber = selection.startLineNumber;
        const getLine = (lineNumber) => model.getLineContent(lineNumber);
        if (!hasUnescapedPipe(getLine(currentLineNumber))) {
            return null;
        }

        let startLine = currentLineNumber;
        while (startLine > 1 && hasUnescapedPipe(getLine(startLine - 1))) {
            startLine -= 1;
        }

        let endLine = currentLineNumber;
        const lineCount = model.getLineCount();
        while (endLine < lineCount && hasUnescapedPipe(getLine(endLine + 1))) {
            endLine += 1;
        }

        const lineEntries = [];
        for (let line = startLine; line <= endLine; line += 1) {
            const text = getLine(line);
            lineEntries.push({
                lineNumber: line,
                text,
                cells: splitMarkdownTableRow(text)
            });
        }

        if (lineEntries.length < 2) {
            return null;
        }

        const separatorRowIndex = lineEntries.findIndex((entry) => isMarkdownTableSeparatorRow(entry.cells));
        if (separatorRowIndex < 0) {
            return null;
        }

        const columnCount = lineEntries.reduce((max, entry) => Math.max(max, entry.cells.length), 0);
        if (columnCount <= 0) {
            return null;
        }

        const currentLine = getLine(currentLineNumber);
        const cursorOffsetInLine = Math.max(0, selection.startColumn - 1);
        const pipesBeforeCursor = countUnescapedPipesBeforeOffset(currentLine, cursorOffsetInLine);
        const activeColumnIndex = Math.max(0, Math.min(columnCount - 1, pipesBeforeCursor - 1));

        const rows = lineEntries.map((entry) => {
            const padded = [...entry.cells];
            while (padded.length < columnCount) {
                padded.push('');
            }
            return padded;
        });

        return {
            startLine,
            endLine,
            rows,
            separatorRowIndex,
            activeColumnIndex,
            activeRowIndex: currentLineNumber - startLine
        };
    };

    let modifyMarkdownTableColumn = (editor, action) => {
        const table = findMarkdownTableAtSelection(editor);
        if (!table) {
            return { ok: false, reason: 'not_found' };
        }

        const model = editor.getModel();
        if (!model) {
            return { ok: false, reason: 'not_found' };
        }

        const { rows, separatorRowIndex, activeColumnIndex, startLine, endLine } = table;
        const columnCount = rows[0] ? rows[0].length : 0;
        if (action === 'remove' && columnCount <= 1) {
            return { ok: false, reason: 'min_columns' };
        }

        const nextRows = rows.map((row, rowIndex) => {
            const next = [...row];
            if (action === 'add') {
                if (rowIndex === separatorRowIndex) {
                    const currentSeparator = String(next[activeColumnIndex] || '').trim();
                    const separatorCell = /^:?-{3,}:?$/.test(currentSeparator) ? currentSeparator : ':---';
                    next.splice(activeColumnIndex + 1, 0, separatorCell);
                } else {
                    next.splice(activeColumnIndex + 1, 0, '');
                }
                return next;
            }
            next.splice(activeColumnIndex, 1);
            return next;
        });

        const nextTableText = buildMarkdownTableFromRows(nextRows);
        const replaceRange = new monaco.Range(startLine, 1, endLine, model.getLineMaxColumn(endLine));
        editor.executeEdits('format-toolbar', [{
            range: replaceRange,
            text: nextTableText,
            forceMoveMarkers: true
        }]);

        const nextColumnIndex = action === 'add'
            ? Math.min(nextRows[0].length - 1, activeColumnIndex + 1)
            : Math.min(nextRows[0].length - 1, activeColumnIndex);
        const anchorLine = startLine;
        const preferredColumn = getCellStartColumnInRenderedRow(nextRows[0], nextColumnIndex);
        const lineMaxColumn = model.getLineMaxColumn(anchorLine);
        const safeColumn = Math.max(1, Math.min(lineMaxColumn, preferredColumn));
        editor.setSelection(new monaco.Selection(anchorLine, safeColumn, anchorLine, safeColumn));
        editor.focus();
        return { ok: true };
    };

    let getPreviewSelectionTableCell = () => {
        const output = getPreviewOutputElement();
        const selection = window.getSelection();
        if (!output || !selection || selection.rangeCount === 0) {
            return null;
        }
        const range = selection.getRangeAt(0);
        if (!isRangeInsideElement(range, output)) {
            return null;
        }
        const containerNode = range.startContainer;
        const element = containerNode.nodeType === Node.ELEMENT_NODE
            ? containerNode
            : containerNode.parentElement;
        if (!(element instanceof Element)) {
            return null;
        }
        const cell = element.closest('td,th');
        const table = element.closest('table');
        if (!cell || !table || !output.contains(table)) {
            return null;
        }
        return {
            cell,
            table
        };
    };

    let movePreviewCaretToCell = (cell) => {
        const selection = window.getSelection();
        if (!(cell instanceof Element) || !selection) {
            return;
        }
        const range = document.createRange();
        range.selectNodeContents(cell);
        range.collapse(true);
        selection.removeAllRanges();
        selection.addRange(range);
        previewSavedRange = range.cloneRange();
    };

    let modifyPreviewTableColumn = (action) => {
        const target = getPreviewSelectionTableCell();
        if (!target) {
            return { ok: false, reason: 'not_found' };
        }

        const { cell, table } = target;
        const rows = Array.from(table.querySelectorAll('tr')).filter((row) => row.cells.length > 0);
        if (rows.length === 0) {
            return { ok: false, reason: 'not_found' };
        }

        const currentIndex = Math.max(0, cell.cellIndex);
        const maxColumns = rows.reduce((max, row) => Math.max(max, row.cells.length), 0);
        if (action === 'remove' && maxColumns <= 1) {
            return { ok: false, reason: 'min_columns' };
        }

        if (action === 'add') {
            rows.forEach((row) => {
                const reference = row.cells[Math.min(currentIndex, row.cells.length - 1)] || null;
                const parentTag = row.parentElement ? row.parentElement.tagName.toLowerCase() : '';
                const newCell = document.createElement(parentTag === 'thead' ? 'th' : 'td');
                newCell.textContent = '';
                if (reference && reference.nextSibling) {
                    row.insertBefore(newCell, reference.nextSibling);
                    return;
                }
                row.appendChild(newCell);
            });
            const currentRow = cell.parentElement;
            const nextCell = currentRow && currentRow.cells[Math.min(currentIndex + 1, currentRow.cells.length - 1)];
            movePreviewCaretToCell(nextCell || cell);
            return { ok: true };
        }

        rows.forEach((row) => {
            if (row.cells[currentIndex]) {
                row.deleteCell(currentIndex);
                return;
            }
            if (row.cells.length > 0) {
                row.deleteCell(row.cells.length - 1);
            }
        });
        const currentRow = cell.parentElement;
        const fallbackIndex = Math.max(0, Math.min(currentIndex, (currentRow ? currentRow.cells.length : 1) - 1));
        const nextCell = currentRow && currentRow.cells[fallbackIndex];
        movePreviewCaretToCell(nextCell || table.querySelector('td,th'));
        return { ok: true };
    };

    let modifyTableColumn = (editor, action) => {
        if (previewEditMode) {
            const output = getPreviewOutputElement();
            if (!output) {
                return { ok: false, reason: 'not_found' };
            }
            output.focus();
            restorePreviewSelectionRange();
            return modifyPreviewTableColumn(action);
        }
        return modifyMarkdownTableColumn(editor, action);
    };

    let modifyMarkdownTableRow = (editor, action) => {
        const table = findMarkdownTableAtSelection(editor);
        if (!table) {
            return { ok: false, reason: 'not_found' };
        }

        const model = editor.getModel();
        if (!model) {
            return { ok: false, reason: 'not_found' };
        }

        const { rows, separatorRowIndex, activeColumnIndex, activeRowIndex, startLine, endLine } = table;
        const dataRowStartIndex = separatorRowIndex + 1;
        const hasDataRows = rows.length > dataRowStartIndex;
        const normalizedDataRowIndex = activeRowIndex > separatorRowIndex
            ? activeRowIndex
            : dataRowStartIndex;

        if (action === 'remove') {
            if (!hasDataRows) {
                return { ok: false, reason: 'no_data_rows' };
            }
            const removeIndex = Math.max(dataRowStartIndex, Math.min(rows.length - 1, normalizedDataRowIndex));
            rows.splice(removeIndex, 1);
        } else {
            const insertIndex = hasDataRows
                ? Math.max(dataRowStartIndex, Math.min(rows.length, normalizedDataRowIndex + 1))
                : rows.length;
            const columnCount = rows[0] ? rows[0].length : 1;
            rows.splice(insertIndex, 0, new Array(columnCount).fill(''));
        }

        const nextTableText = buildMarkdownTableFromRows(rows);
        const replaceRange = new monaco.Range(startLine, 1, endLine, model.getLineMaxColumn(endLine));
        editor.executeEdits('format-toolbar', [{
            range: replaceRange,
            text: nextTableText,
            forceMoveMarkers: true
        }]);

        const nextDataRowIndex = rows.length > dataRowStartIndex
            ? Math.max(dataRowStartIndex, Math.min(rows.length - 1, normalizedDataRowIndex))
            : separatorRowIndex;
        const anchorLine = startLine + nextDataRowIndex;
        const preferredColumn = getCellStartColumnInRenderedRow(rows[nextDataRowIndex] || rows[0], activeColumnIndex);
        const lineMaxColumn = model.getLineMaxColumn(anchorLine);
        const safeColumn = Math.max(1, Math.min(lineMaxColumn, preferredColumn));
        editor.setSelection(new monaco.Selection(anchorLine, safeColumn, anchorLine, safeColumn));
        editor.focus();
        return { ok: true };
    };

    let modifyPreviewTableRow = (action) => {
        const target = getPreviewSelectionTableCell();
        if (!target) {
            return { ok: false, reason: 'not_found' };
        }

        const { cell, table } = target;
        const allRows = Array.from(table.querySelectorAll('tr')).filter((row) => row.cells.length > 0);
        if (allRows.length === 0) {
            return { ok: false, reason: 'not_found' };
        }

        const row = cell.parentElement;
        const tbody = table.tBodies && table.tBodies.length > 0 ? table.tBodies[0] : null;
        const bodyRows = tbody
            ? Array.from(tbody.rows).filter((bodyRow) => bodyRow.cells.length > 0)
            : allRows.slice(1);
        const columnIndex = Math.max(0, cell.cellIndex);

        if (action === 'remove') {
            if (bodyRows.length === 0) {
                return { ok: false, reason: 'no_data_rows' };
            }
            const targetRow = row && bodyRows.includes(row) ? row : bodyRows[0];
            const fallbackRow = targetRow.nextElementSibling instanceof HTMLTableRowElement
                ? targetRow.nextElementSibling
                : (targetRow.previousElementSibling instanceof HTMLTableRowElement ? targetRow.previousElementSibling : null);
            targetRow.remove();
            const nextRow = fallbackRow && fallbackRow.cells.length > 0
                ? fallbackRow
                : (tbody && tbody.rows.length > 0 ? tbody.rows[0] : null);
            const nextCell = nextRow && nextRow.cells[Math.min(columnIndex, nextRow.cells.length - 1)];
            movePreviewCaretToCell(nextCell || table.querySelector('td,th'));
            return { ok: true };
        }

        const baseRow = row && row.cells.length > 0 ? row : (tbody && tbody.rows[0] ? tbody.rows[0] : null);
        const referenceRow = (baseRow && (!tbody || tbody.contains(baseRow))) ? baseRow : null;
        const columnCount = referenceRow
            ? referenceRow.cells.length
            : (bodyRows[0] ? bodyRows[0].cells.length : (allRows[0] ? allRows[0].cells.length : 1));
        const newRow = document.createElement('tr');
        for (let index = 0; index < Math.max(1, columnCount); index += 1) {
            const newCell = document.createElement('td');
            newCell.textContent = '';
            newRow.appendChild(newCell);
        }

        if (tbody) {
            if (referenceRow && referenceRow.nextSibling) {
                tbody.insertBefore(newRow, referenceRow.nextSibling);
            } else {
                tbody.appendChild(newRow);
            }
        } else if (referenceRow && referenceRow.nextSibling) {
            table.insertBefore(newRow, referenceRow.nextSibling);
        } else {
            table.appendChild(newRow);
        }

        const nextCell = newRow.cells[Math.min(columnIndex, newRow.cells.length - 1)];
        movePreviewCaretToCell(nextCell || newRow.cells[0]);
        return { ok: true };
    };

    let modifyTableRow = (editor, action) => {
        if (previewEditMode) {
            const output = getPreviewOutputElement();
            if (!output) {
                return { ok: false, reason: 'not_found' };
            }
            output.focus();
            restorePreviewSelectionRange();
            return modifyPreviewTableRow(action);
        }
        return modifyMarkdownTableRow(editor, action);
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
            const unwrapped = unwrapPreviewBlockquoteAtSelection();
            if (!unwrapped) {
                executePreviewCommand('formatBlock', '<blockquote>');
            }
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

        const formatButtons = Array.from(toolbar.querySelectorAll('.format-button'));
        const buttonByFormat = new Map();
        formatButtons.forEach((button) => {
            const formatType = button.getAttribute('data-format');
            if (!formatType) {
                return;
            }
            buttonByFormat.set(formatType, button);
            button.setAttribute('aria-pressed', 'false');
        });

        const setButtonActiveState = (formatType, active) => {
            const button = buttonByFormat.get(formatType);
            if (!button) {
                return;
            }
            button.classList.toggle('active', active === true);
            button.setAttribute('aria-pressed', active === true ? 'true' : 'false');
        };

        const setLinkPopoverState = (open) => {
            linkPopoverOpen = open === true;
            if (linkTool) {
                linkTool.classList.toggle('open', linkPopoverOpen);
            }
            if (linkPopover) {
                linkPopover.setAttribute('aria-hidden', linkPopoverOpen ? 'false' : 'true');
            }
            const linkButton = buttonByFormat.get('link');
            if (linkButton) {
                linkButton.setAttribute('aria-expanded', linkPopoverOpen ? 'true' : 'false');
            }
        };

        const setImagePopoverState = (open) => {
            imagePopoverOpen = open === true;
            if (imageTool) {
                imageTool.classList.toggle('open', imagePopoverOpen);
            }
            if (imagePopover) {
                imagePopover.setAttribute('aria-hidden', imagePopoverOpen ? 'false' : 'true');
            }
            const imageButton = buttonByFormat.get('image');
            if (imageButton) {
                imageButton.setAttribute('aria-expanded', imagePopoverOpen ? 'true' : 'false');
            }
        };

        const setNotePopoverState = (open) => {
            notePopoverOpen = open === true;
            if (noteTool) {
                noteTool.classList.toggle('open', notePopoverOpen);
            }
            if (notePopover) {
                notePopover.setAttribute('aria-hidden', notePopoverOpen ? 'false' : 'true');
            }
            const noteButton = buttonByFormat.get('note');
            if (noteButton) {
                noteButton.setAttribute('aria-expanded', notePopoverOpen ? 'true' : 'false');
            }
        };

        const linkTool = document.querySelector('#link-tool');
        const linkPopover = document.querySelector('#link-popover');
        const linkTextInput = document.querySelector('#link-text-input');
        const linkTargetInput = document.querySelector('#link-target-input');
        const linkBrowseButton = document.querySelector('#link-browse-button');
        const linkAnchorSelect = document.querySelector('#link-anchor-select');
        const linkInsertButton = document.querySelector('#link-insert-button');
        const linkRemoveButton = document.querySelector('#link-remove-button');
        const imageTool = document.querySelector('#image-tool');
        const imagePopover = document.querySelector('#image-popover');
        const imageAltInput = document.querySelector('#image-alt-input');
        const imageSrcInput = document.querySelector('#image-src-input');
        const imageBrowseButton = document.querySelector('#image-browse-button');
        const imageTitleInput = document.querySelector('#image-title-input');
        const imageInsertButton = document.querySelector('#image-insert-button');
        const noteTool = document.querySelector('#note-tool');
        const notePopover = document.querySelector('#note-popover');
        const noteSearchInput = document.querySelector('#note-search-input');
        const noteEntriesList = document.querySelector('#note-entries-list');
        const noteNewEntryButton = document.querySelector('#note-new-entry-button');
        const noteBindButton = document.querySelector('#note-bind-button');
        const noteUnbindButton = document.querySelector('#note-unbind-button');
        const noteEditorOverlay = document.querySelector('#note-editor-overlay');
        const noteEditorModalTitle = document.querySelector('#note-editor-modal-title');
        const noteEditorCloseButton = document.querySelector('#note-editor-close-button');
        const noteEditorTitleInput = document.querySelector('#note-editor-title-input');
        const noteEditorBodyInput = document.querySelector('#note-editor-body-input');
        const noteEditorSaveButton = document.querySelector('#note-editor-save-button');
        const noteEditorCancelButton = document.querySelector('#note-editor-cancel-button');
        const noteEditorDeleteButton = document.querySelector('#note-editor-delete-button');
        const tableTool = document.querySelector('#table-tool');
        const tablePopover = document.querySelector('#table-popover');
        const tableColumnsInput = document.querySelector('#table-columns-input');
        const tableRowsInput = document.querySelector('#table-rows-input');
        const tableAlignSelect = document.querySelector('#table-align-select');
        const tableHeaderCheckbox = document.querySelector('#table-header-checkbox');
        const tableInsertButton = document.querySelector('#table-insert-button');
        const tableAddColumnButton = document.querySelector('#table-add-column-button');
        const tableRemoveColumnButton = document.querySelector('#table-remove-column-button');
        const tableAddRowButton = document.querySelector('#table-add-row-button');
        const tableRemoveRowButton = document.querySelector('#table-remove-row-button');
        let linkPopoverOpen = false;
        let imagePopoverOpen = false;
        let notePopoverOpen = false;
        let tablePopoverOpen = false;
        let noteEditorOpen = false;
        let linkPopoverMode = 'insert';
        let linkPopoverContext = null;
        let notePopoverContext = null;
        let selectedNoteEntryId = '';
        let noteEditorMode = 'create';
        let noteEditorEntryId = '';

        const setTablePopoverState = (open) => {
            tablePopoverOpen = open === true;
            if (tableTool) {
                tableTool.classList.toggle('open', tablePopoverOpen);
            }
            if (tablePopover) {
                tablePopover.setAttribute('aria-hidden', tablePopoverOpen ? 'false' : 'true');
            }
            const tableButton = buttonByFormat.get('table');
            if (tableButton) {
                tableButton.setAttribute('aria-expanded', tablePopoverOpen ? 'true' : 'false');
            }
        };

        const closeLinkPopover = () => {
            linkPopoverMode = 'insert';
            linkPopoverContext = null;
            if (linkInsertButton) {
                linkInsertButton.textContent = t('linkInsert');
            }
            if (linkRemoveButton) {
                linkRemoveButton.hidden = true;
            }
            setLinkPopoverState(false);
        };

        const closeImagePopover = () => {
            setImagePopoverState(false);
        };

        const closeNotePopover = () => {
            notePopoverContext = null;
            if (noteUnbindButton) {
                noteUnbindButton.hidden = true;
            }
            setNotePopoverState(false);
        };

        const setNoteEditorState = (open) => {
            noteEditorOpen = open === true;
            if (noteEditorOverlay) {
                noteEditorOverlay.hidden = !noteEditorOpen;
                noteEditorOverlay.classList.toggle('open', noteEditorOpen);
                noteEditorOverlay.setAttribute('aria-hidden', noteEditorOpen ? 'false' : 'true');
            }
            document.body.classList.toggle('note-editor-open', noteEditorOpen);
        };

        const refreshNoteEditorHeader = () => {
            if (!noteEditorModalTitle) {
                return;
            }
            noteEditorModalTitle.textContent = noteEditorMode === 'edit'
                ? t('noteEditorEditTitle')
                : t('noteEditorAddTitle');
            noteEditorModalTitle.setAttribute('data-mode', noteEditorMode);
        };

        const closeNoteEditorModal = () => {
            noteEditorMode = 'create';
            noteEditorEntryId = '';
            refreshNoteEditorHeader();
            setNoteEditorState(false);
        };

        const closeTablePopover = () => {
            setTablePopoverState(false);
        };

        const closeInlinePopovers = () => {
            closeLinkPopover();
            closeImagePopover();
            closeNotePopover();
            closeTablePopover();
        };

        const getSelectedTextFromCurrentContext = () => {
            if (previewEditMode) {
                restorePreviewSelectionRange();
                return (window.getSelection() ? window.getSelection().toString().trim() : '');
            }
            const model = editor.getModel();
            const selection = editor.getSelection();
            if (!model || !selection) {
                return '';
            }
            return model.getValueInRange(selection).trim();
        };

        const getSelectionOffsets = (model, selection) => {
            if (!model || !selection) {
                return null;
            }
            const startOffset = model.getOffsetAt({
                lineNumber: selection.startLineNumber,
                column: selection.startColumn
            });
            const endOffset = model.getOffsetAt({
                lineNumber: selection.endLineNumber,
                column: selection.endColumn
            });
            return {
                startOffset,
                endOffset,
                hasSelection: startOffset !== endOffset
            };
        };

        const findMarkdownLinkInLine = (lineText, startInLine, endInLine) => {
            const pattern = /\[([^\]\n]*)\]\(([^)\n]*)\)/g;
            let best = null;
            let match = pattern.exec(lineText);
            while (match) {
                const matchStart = match.index;
                const matchEnd = matchStart + match[0].length;
                const intersects = !(endInLine <= matchStart || startInLine >= matchEnd);
                if (intersects) {
                    if (!best || match[0].length < best.raw.length) {
                        best = {
                            raw: match[0],
                            label: match[1] || '',
                            href: match[2] || '',
                            startInLine: matchStart,
                            endInLine: matchEnd
                        };
                    }
                }
                match = pattern.exec(lineText);
            }
            return best;
        };

        const renderNoteEntriesList = ({ selectedId = '', searchValue = '' } = {}) => {
            if (!noteEntriesList) {
                return;
            }
            const normalizedSearch = String(searchValue || '').trim().toLowerCase();
            const preferredSelected = String(selectedId || selectedNoteEntryId || '').trim();
            const sorted = notesArchive
                .slice()
                .sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
            const filtered = normalizedSearch
                ? sorted.filter((entry) => {
                    const title = String(entry?.title || '').toLowerCase();
                    const body = String(entry?.body || '').toLowerCase();
                    return title.includes(normalizedSearch) || body.includes(normalizedSearch);
                })
                : sorted;

            noteEntriesList.innerHTML = '';
            if (filtered.length === 0) {
                const emptyState = document.createElement('div');
                emptyState.className = 'note-empty-state';
                emptyState.textContent = t('noteEmpty');
                noteEntriesList.appendChild(emptyState);
            } else {
                filtered.forEach((entry) => {
                    if (!entry || !entry.id) {
                        return;
                    }

                    const row = document.createElement('div');
                    row.className = 'note-entry-row';
                    row.setAttribute('data-note-id', entry.id);

                    const selectButton = document.createElement('button');
                    selectButton.type = 'button';
                    selectButton.className = 'note-entry-item note-entry-select';
                    selectButton.setAttribute('data-note-id', entry.id);
                    selectButton.setAttribute('role', 'option');
                    const title = document.createElement('span');
                    title.className = 'note-entry-title';
                    title.textContent = entry.title;
                    const preview = document.createElement('span');
                    preview.className = 'note-entry-preview';
                    preview.textContent = entry.body;
                    selectButton.appendChild(title);
                    selectButton.appendChild(preview);

                    const editButton = document.createElement('button');
                    editButton.type = 'button';
                    editButton.className = 'note-entry-edit';
                    editButton.setAttribute('data-note-id', entry.id);
                    editButton.setAttribute('aria-label', t('noteEdit'));
                    editButton.setAttribute('title', t('noteEdit'));
                    editButton.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 17.3V21h3.7l10.8-10.8-3.7-3.7L3 17.3zm17.7-10.2a1 1 0 0 0 0-1.4L18.1 3.1a1 1 0 0 0-1.4 0l-1.6 1.6 3.7 3.7 1.9-1.3z"/></svg>';

                    if (entry.id === preferredSelected) {
                        selectButton.classList.add('active');
                        selectButton.setAttribute('aria-selected', 'true');
                    }

                    row.appendChild(selectButton);
                    row.appendChild(editButton);
                    noteEntriesList.appendChild(row);
                });
            }

            if (preferredSelected && !filtered.some((entry) => entry && entry.id === preferredSelected)) {
                selectedNoteEntryId = '';
            } else if (preferredSelected) {
                selectedNoteEntryId = preferredSelected;
            }
        };

        const getSelectedNoteArchiveId = () => {
            return String(selectedNoteEntryId || '').trim();
        };

        const getEditorNoteContext = () => {
            const model = editor.getModel();
            const selection = editor.getSelection();
            if (!model || !selection) {
                return null;
            }

            const lineNumber = selection.startLineNumber;
            if (lineNumber !== selection.endLineNumber) {
                return null;
            }

            const lineText = model.getLineContent(lineNumber);
            const lineStartOffset = model.getOffsetAt({ lineNumber, column: 1 });
            const offsets = getSelectionOffsets(model, selection);
            if (!offsets) {
                return null;
            }
            const startInLine = offsets.startOffset - lineStartOffset;
            const endInLine = offsets.endOffset - lineStartOffset;
            const pattern = /<span\s+class="lectr-note-ref"\s+data-lectr-note-id="([^"]+)"\s*>([\s\S]*?)<\/span>/g;
            let match = pattern.exec(lineText);
            while (match) {
                const matchStart = match.index;
                const matchEnd = matchStart + match[0].length;
                const intersects = !(endInLine <= matchStart || startInLine >= matchEnd);
                if (intersects) {
                    const noteId = String(match[1] || '').trim();
                    const label = String(match[2] || '');
                    const range = createMonacoRangeFromOffsets(
                        model,
                        lineStartOffset + matchStart,
                        lineStartOffset + matchEnd
                    );
                    if (!range || !noteId) {
                        return null;
                    }
                    return {
                        type: 'editor',
                        noteId,
                        label,
                        range
                    };
                }
                match = pattern.exec(lineText);
            }
            return null;
        };

        const getPreviewNoteContext = () => {
            const contextElement = getPreviewSelectionContextElement();
            if (!(contextElement instanceof Element)) {
                return null;
            }
            const noteElement = contextElement.closest('[data-lectr-note-id]');
            if (!noteElement) {
                return null;
            }
            const noteId = String(noteElement.getAttribute('data-lectr-note-id') || '').trim();
            if (!noteId) {
                return null;
            }
            return {
                type: 'preview',
                noteId,
                noteElement,
                label: (noteElement.textContent || '').trim()
            };
        };

        const resolveNotePopoverContext = () => {
            if (previewEditMode) {
                return getPreviewNoteContext();
            }
            return getEditorNoteContext();
        };

        const upsertNoteArchiveEntry = ({ entryId = '', title = '', body = '' }) => {
            const safeId = String(entryId || '').trim();
            const safeTitle = String(title || '').trim();
            const safeBody = String(body || '').trim();
            if (!safeTitle) {
                window.alert(t('noteTitleRequired'));
                return null;
            }
            if (!safeBody) {
                window.alert(t('noteBodyRequired'));
                return null;
            }

            const now = Date.now();
            if (safeId) {
                const index = notesArchive.findIndex((entry) => entry && entry.id === safeId);
                if (index !== -1) {
                    notesArchive[index] = {
                        ...notesArchive[index],
                        title: safeTitle,
                        body: safeBody,
                        updatedAt: now
                    };
                    saveNotesArchive();
                    return notesArchive[index];
                }
            }

            const newEntry = {
                id: `note-${now.toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
                title: safeTitle,
                body: safeBody,
                updatedAt: now
            };
            notesArchive.push(newEntry);
            saveNotesArchive();
            return newEntry;
        };

        const applyNoteBindingToEditor = (noteId) => {
            const model = editor.getModel();
            const selection = editor.getSelection();
            if (!model || !selection) {
                return false;
            }
            const selectedText = model.getValueInRange(selection);
            if (!selectedText || !selectedText.trim()) {
                window.alert(t('noteSelectionRequired'));
                return false;
            }
            if (selectedText.includes('\n')) {
                window.alert(t('noteSelectionRequired'));
                return false;
            }

            const escapedText = escapeHtml(selectedText);
            const snippet = `<span class="lectr-note-ref" data-lectr-note-id="${escapeHtmlAttribute(noteId)}">${escapedText}</span>`;
            editor.pushUndoStop();
            editor.executeEdits('note-bind', [{
                range: selection,
                text: snippet,
                forceMoveMarkers: true
            }]);
            editor.pushUndoStop();
            return true;
        };

        const applyNoteBindingToPreview = (noteId) => {
            const output = getPreviewOutputElement();
            const selection = window.getSelection();
            if (!output || !selection) {
                return false;
            }
            restorePreviewSelectionRange();
            if (selection.rangeCount === 0) {
                window.alert(t('noteSelectionRequired'));
                return false;
            }
            const range = selection.getRangeAt(0);
            if (range.collapsed || !isRangeInsideElement(range, output)) {
                window.alert(t('noteSelectionRequired'));
                return false;
            }

            const selectedText = selection.toString();
            if (!selectedText || !selectedText.trim()) {
                window.alert(t('noteSelectionRequired'));
                return false;
            }

            const noteElement = document.createElement('span');
            noteElement.className = 'lectr-note-ref';
            noteElement.setAttribute('data-lectr-note-id', noteId);
            noteElement.textContent = selectedText;
            range.deleteContents();
            range.insertNode(noteElement);
            updatePreviewSelectionAfterNode(noteElement);
            return true;
        };

        const refreshLinkPopoverActions = () => {
            if (linkInsertButton) {
                linkInsertButton.textContent = linkPopoverMode === 'edit' ? t('linkUpdate') : t('linkInsert');
            }
            if (linkRemoveButton) {
                linkRemoveButton.hidden = linkPopoverMode !== 'edit';
            }
        };

        const createMonacoRangeFromOffsets = (model, startOffset, endOffset) => {
            if (!model || !Number.isFinite(startOffset) || !Number.isFinite(endOffset)) {
                return null;
            }
            const safeStart = Math.max(0, Math.min(startOffset, endOffset));
            const safeEnd = Math.max(safeStart, Math.max(startOffset, endOffset));
            const startPosition = model.getPositionAt(safeStart);
            const endPosition = model.getPositionAt(safeEnd);
            return new monaco.Range(
                startPosition.lineNumber,
                startPosition.column,
                endPosition.lineNumber,
                endPosition.column
            );
        };

        const getEditorLinkContext = () => {
            const model = editor.getModel();
            const selection = editor.getSelection();
            if (!model || !selection) {
                return null;
            }
            const lineNumber = selection.startLineNumber;
            if (lineNumber !== selection.endLineNumber) {
                return null;
            }
            const lineText = model.getLineContent(lineNumber);
            const lineStartOffset = model.getOffsetAt({ lineNumber, column: 1 });
            const offsets = getSelectionOffsets(model, selection);
            if (!offsets) {
                return null;
            }
            const startInLine = offsets.startOffset - lineStartOffset;
            const endInLine = offsets.endOffset - lineStartOffset;
            const linkMatch = findMarkdownLinkInLine(lineText, startInLine, endInLine);
            if (!linkMatch) {
                return null;
            }
            const range = createMonacoRangeFromOffsets(
                model,
                lineStartOffset + linkMatch.startInLine,
                lineStartOffset + linkMatch.endInLine
            );
            if (!range) {
                return null;
            }

            return {
                type: 'editor',
                label: linkMatch.label,
                href: linkMatch.href,
                range
            };
        };

        const getPreviewLinkContext = () => {
            const contextElement = getPreviewSelectionContextElement();
            if (!(contextElement instanceof Element)) {
                return null;
            }

            const anchorElement = contextElement.closest('a[href]');
            if (!anchorElement) {
                return null;
            }

            return {
                type: 'preview',
                anchorElement,
                label: (anchorElement.textContent || '').trim(),
                href: String(anchorElement.getAttribute('href') || '').trim()
            };
        };

        const resolveLinkPopoverContext = () => {
            if (previewEditMode) {
                return getPreviewLinkContext();
            }
            return getEditorLinkContext();
        };

        const parseAnchorsFromMarkdown = (content) => {
            const raw = typeof content === 'string' ? content : '';
            if (!raw) {
                return [];
            }

            const lines = raw.split(/\r?\n/);
            const anchors = [];
            const slugCounts = new Map();
            let inFence = false;

            for (let index = 0; index < lines.length; index += 1) {
                const line = lines[index];
                if (/^\s*(```|~~~)/.test(line)) {
                    inFence = !inFence;
                    continue;
                }
                if (inFence) {
                    continue;
                }
                const match = line.match(/^\s{0,3}(#{1,6})\s+(.+?)\s*#*\s*$/);
                if (!match) {
                    continue;
                }
                const level = match[1].length;
                const label = normalizeHeadingLabelText(match[2] || '');
                if (!label) {
                    continue;
                }
                const baseSlug = slugifyHeadingId(label);
                const duplicateCount = slugCounts.get(baseSlug) || 0;
                slugCounts.set(baseSlug, duplicateCount + 1);
                const anchor = duplicateCount === 0 ? baseSlug : `${baseSlug}-${duplicateCount}`;
                anchors.push({ label, anchor, level, line: index + 1 });
            }

            return anchors;
        };

        const setLinkAnchorOptions = (anchors, selectedAnchor = '') => {
            if (!linkAnchorSelect) {
                return;
            }

            const safeAnchors = Array.isArray(anchors)
                ? anchors.filter((item) => item && typeof item.anchor === 'string' && item.anchor.trim())
                : [];
            const normalizedSelectedAnchor = String(selectedAnchor || '').trim().replace(/^#+/, '');
            linkAnchorSelect.innerHTML = '';

            const noneOption = document.createElement('option');
            noneOption.value = '';
            noneOption.textContent = t('linkAnchorNone');
            linkAnchorSelect.appendChild(noneOption);

            safeAnchors.forEach((item) => {
                const option = document.createElement('option');
                option.value = item.anchor.trim();
                const level = Number.isFinite(item.level) ? Math.max(1, Math.min(6, item.level)) : 1;
                const prefix = level > 1 ? `${'  '.repeat(level - 1)}- ` : '';
                option.textContent = `${prefix}${String(item.label || item.anchor).trim()}`;
                linkAnchorSelect.appendChild(option);
            });

            const hasSelected = normalizedSelectedAnchor
                && safeAnchors.some((item) => item.anchor.trim() === normalizedSelectedAnchor);
            if (normalizedSelectedAnchor && !hasSelected) {
                const customOption = document.createElement('option');
                customOption.value = normalizedSelectedAnchor;
                customOption.textContent = `#${normalizedSelectedAnchor}`;
                linkAnchorSelect.appendChild(customOption);
            }

            linkAnchorSelect.value = normalizedSelectedAnchor || '';
            linkAnchorSelect.disabled = false;
        };

        const pickLinkFileFallback = async () => new Promise((resolve) => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.md,.markdown,.txt,text/markdown,text/plain';
            input.addEventListener('change', async () => {
                const file = input.files && input.files[0];
                if (!file) {
                    resolve({ picked: false });
                    return;
                }
                const content = await file.text();
                resolve({
                    picked: true,
                    linkTarget: file.name,
                    fileName: file.name,
                    anchors: parseAnchorsFromMarkdown(content)
                });
            }, { once: true });
            input.click();
        });

        const pickImageFileFallback = async () => new Promise((resolve) => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = 'image/*';
            input.addEventListener('change', async () => {
                const file = input.files && input.files[0];
                if (!file) {
                    resolve({ picked: false });
                    return;
                }
                resolve({
                    picked: true,
                    linkTarget: file.name,
                    fileName: file.name
                });
            }, { once: true });
            input.click();
        });

        const updatePreviewSelectionAfterNode = (node) => {
            const selection = window.getSelection();
            if (!selection || !node || !node.parentNode) {
                return;
            }

            const nextRange = document.createRange();
            nextRange.setStartAfter(node);
            nextRange.collapse(true);
            selection.removeAllRanges();
            selection.addRange(nextRange);
            previewSavedRange = nextRange.cloneRange();
        };

        const insertLinkFromPopover = () => {
            const selectedText = linkPopoverContext && typeof linkPopoverContext.label === 'string' && linkPopoverContext.label.trim()
                ? linkPopoverContext.label.trim()
                : getSelectedTextFromCurrentContext();
            const label = String(linkTextInput ? linkTextInput.value : '').trim() || selectedText || t('linkDefaultText');
            const target = String(linkTargetInput ? linkTargetInput.value : '').trim();
            const anchor = String(linkAnchorSelect ? linkAnchorSelect.value : '').trim();
            const href = buildLinkHref(target, anchor);
            if (!href) {
                window.alert(t('linkTargetRequired'));
                return false;
            }

            if (previewEditMode) {
                if (linkPopoverMode === 'edit'
                    && linkPopoverContext
                    && linkPopoverContext.type === 'preview'
                    && linkPopoverContext.anchorElement instanceof HTMLAnchorElement
                    && linkPopoverContext.anchorElement.isConnected) {
                    const anchorElement = linkPopoverContext.anchorElement;
                    anchorElement.setAttribute('href', href);
                    anchorElement.textContent = label;
                    updatePreviewSelectionAfterNode(anchorElement);
                } else {
                    restorePreviewSelectionRange();
                    insertHtmlIntoPreviewSelection(`<a href="${escapeHtml(href)}">${escapeHtml(label)}</a>`);
                }
                schedulePreviewSyncFromToolbar(editor);
            } else {
                editor.pushUndoStop();
                if (linkPopoverMode === 'edit'
                    && linkPopoverContext
                    && linkPopoverContext.type === 'editor'
                    && linkPopoverContext.range) {
                    applyLinkFormatWithValues(editor, { label, href, replaceRange: linkPopoverContext.range });
                } else {
                    applyLinkFormatWithValues(editor, { label, href });
                }
                editor.pushUndoStop();
            }
            closeLinkPopover();
            return true;
        };

        const removeLinkFromPopover = () => {
            if (linkPopoverMode !== 'edit' || !linkPopoverContext) {
                closeLinkPopover();
                return false;
            }

            const replacementText = String(linkPopoverContext.label || linkPopoverContext.href || '');
            if (previewEditMode) {
                if (linkPopoverContext.type !== 'preview'
                    || !(linkPopoverContext.anchorElement instanceof HTMLAnchorElement)
                    || !linkPopoverContext.anchorElement.isConnected) {
                    closeLinkPopover();
                    return false;
                }
                const anchorElement = linkPopoverContext.anchorElement;
                const textNode = document.createTextNode(replacementText);
                anchorElement.replaceWith(textNode);
                updatePreviewSelectionAfterNode(textNode);
                schedulePreviewSyncFromToolbar(editor);
                closeLinkPopover();
                return true;
            }

            if (linkPopoverContext.type !== 'editor' || !linkPopoverContext.range) {
                closeLinkPopover();
                return false;
            }

            editor.pushUndoStop();
            replaceMarkdownRangeWithSnippet(editor, linkPopoverContext.range, replacementText);
            editor.pushUndoStop();
            closeLinkPopover();
            return true;
        };

        const insertImageFromPopover = () => {
            const alt = String(imageAltInput ? imageAltInput.value : '').trim() || t('imageDefaultAlt');
            const src = String(imageSrcInput ? imageSrcInput.value : '').trim();
            const title = String(imageTitleInput ? imageTitleInput.value : '').trim();
            if (!src) {
                window.alert(t('imageSourceRequired'));
                return false;
            }

            if (previewEditMode) {
                const titleAttribute = title ? ` title="${escapeHtml(title)}"` : '';
                restorePreviewSelectionRange();
                insertHtmlIntoPreviewSelection(`<img src="${escapeHtml(src)}" alt="${escapeHtml(alt)}"${titleAttribute}>`);
                schedulePreviewSyncFromToolbar(editor);
            } else {
                applyImageFormatWithValues(editor, { alt, src, title });
            }
            closeImagePopover();
            return true;
        };

        const pickLinkFile = async () => {
            const sourceFilePath = getActiveTab() && getActiveTab().filePath
                ? getActiveTab().filePath
                : null;

            if (window.lectrDesktop && typeof window.lectrDesktop.pickLinkFile === 'function') {
                try {
                    const result = await window.lectrDesktop.pickLinkFile({ sourceFilePath });
                    return result && typeof result === 'object' ? result : { picked: false };
                } catch (error) {
                    // eslint-disable-next-line no-console
                    console.error('Failed to pick link file via desktop bridge', error);
                    return { picked: false };
                }
            }

            return pickLinkFileFallback();
        };

        const pickImageFile = async () => {
            const sourceFilePath = getActiveTab() && getActiveTab().filePath
                ? getActiveTab().filePath
                : null;

            if (window.lectrDesktop && typeof window.lectrDesktop.pickImageFile === 'function') {
                try {
                    const result = await window.lectrDesktop.pickImageFile({ sourceFilePath });
                    return result && typeof result === 'object' ? result : { picked: false };
                } catch (error) {
                    // eslint-disable-next-line no-console
                    console.error('Failed to pick image via desktop bridge', error);
                    return { picked: false };
                }
            }

            return pickImageFileFallback();
        };

        const browseLinkFile = async () => {
            const result = await pickLinkFile();
            if (!result || result.picked !== true) {
                return;
            }

            if (linkTargetInput) {
                const target = typeof result.linkTarget === 'string' && result.linkTarget.trim()
                    ? result.linkTarget.trim()
                    : (result.fileName || '');
                linkTargetInput.value = target;
            }
            setLinkAnchorOptions(result.anchors);

            if (linkAnchorSelect && Array.isArray(result.anchors) && result.anchors.length > 0) {
                linkAnchorSelect.focus();
            }
        };

        const browseImageFile = async () => {
            const result = await pickImageFile();
            if (!result || result.picked !== true) {
                return;
            }

            if (imageSrcInput) {
                const target = typeof result.linkTarget === 'string' && result.linkTarget.trim()
                    ? result.linkTarget.trim()
                    : (result.fileName || '');
                imageSrcInput.value = target;
                imageSrcInput.focus();
            }
        };

        const openNoteEditorModal = ({ mode = 'create', entryId = '' } = {}) => {
            noteEditorMode = mode === 'edit' ? 'edit' : 'create';
            noteEditorEntryId = '';

            if (noteEditorMode === 'edit') {
                const entry = getNoteEntryById(entryId);
                if (!entry) {
                    return false;
                }
                noteEditorEntryId = entry.id;
                if (noteEditorTitleInput) {
                    noteEditorTitleInput.value = entry.title;
                }
                if (noteEditorBodyInput) {
                    noteEditorBodyInput.value = entry.body;
                }
            } else {
                const fallback = notePopoverContext && typeof notePopoverContext.label === 'string'
                    ? notePopoverContext.label.trim()
                    : getSelectedTextFromCurrentContext();
                if (noteEditorTitleInput) {
                    noteEditorTitleInput.value = fallback || '';
                }
                if (noteEditorBodyInput) {
                    noteEditorBodyInput.value = '';
                }
            }

            if (noteEditorDeleteButton) {
                noteEditorDeleteButton.hidden = noteEditorMode !== 'edit';
            }
            refreshNoteEditorHeader();
            setNoteEditorState(true);
            window.requestAnimationFrame(() => {
                if (noteEditorBodyInput && noteEditorMode === 'edit') {
                    noteEditorBodyInput.focus();
                    noteEditorBodyInput.select();
                    return;
                }
                if (noteEditorTitleInput) {
                    noteEditorTitleInput.focus();
                    noteEditorTitleInput.select();
                }
            });
            return true;
        };

        const saveNoteEntryFromModal = () => {
            const title = noteEditorTitleInput ? noteEditorTitleInput.value : '';
            const body = noteEditorBodyInput ? noteEditorBodyInput.value : '';
            const savedEntry = upsertNoteArchiveEntry({
                entryId: noteEditorMode === 'edit' ? noteEditorEntryId : '',
                title,
                body
            });
            if (!savedEntry) {
                return null;
            }

            selectedNoteEntryId = savedEntry.id;
            renderNoteEntriesList({
                selectedId: savedEntry.id,
                searchValue: noteSearchInput ? noteSearchInput.value : ''
            });
            const output = getPreviewOutputElement();
            if (output) {
                applyNoteReferenceDecorations(output);
            }
            hideNoteTooltip();
            closeNoteEditorModal();
            return savedEntry;
        };

        const deleteNoteArchiveEntry = (entryId) => {
            const safeId = String(entryId || '').trim();
            if (!safeId) {
                return false;
            }
            if (!window.confirm(t('noteDeleteConfirm'))) {
                return false;
            }
            notesArchive = notesArchive.filter((entry) => entry && entry.id !== safeId);
            saveNotesArchive();
            if (selectedNoteEntryId === safeId) {
                selectedNoteEntryId = '';
            }
            renderNoteEntriesList({
                selectedId: selectedNoteEntryId,
                searchValue: noteSearchInput ? noteSearchInput.value : ''
            });
            const output = getPreviewOutputElement();
            if (output) {
                applyNoteReferenceDecorations(output);
            }
            hideNoteTooltip();
            return true;
        };

        const bindNoteFromPopover = () => {
            const entryId = getSelectedNoteArchiveId();
            if (!entryId) {
                window.alert(t('noteEntryMissing'));
                return false;
            }

            let bound = false;
            if (previewEditMode) {
                bound = applyNoteBindingToPreview(entryId);
                if (bound) {
                    schedulePreviewSyncFromToolbar(editor);
                }
            } else {
                bound = applyNoteBindingToEditor(entryId);
            }
            if (!bound) {
                return false;
            }

            closeNotePopover();
            return true;
        };

        const removeNoteBindingFromPopover = () => {
            if (!notePopoverContext) {
                closeNotePopover();
                return false;
            }

            if (previewEditMode) {
                if (notePopoverContext.type !== 'preview'
                    || !(notePopoverContext.noteElement instanceof Element)
                    || !notePopoverContext.noteElement.isConnected) {
                    closeNotePopover();
                    return false;
                }
                const noteElement = notePopoverContext.noteElement;
                const plainText = document.createTextNode(noteElement.textContent || '');
                noteElement.replaceWith(plainText);
                updatePreviewSelectionAfterNode(plainText);
                schedulePreviewSyncFromToolbar(editor);
                closeNotePopover();
                return true;
            }

            if (notePopoverContext.type !== 'editor' || !notePopoverContext.range) {
                closeNotePopover();
                return false;
            }
            const replacementText = String(notePopoverContext.label || '');
            editor.pushUndoStop();
            replaceMarkdownRangeWithSnippet(editor, notePopoverContext.range, replacementText);
            editor.pushUndoStop();
            closeNotePopover();
            return true;
        };

        const openLinkPopover = () => {
            closeOpenMenus();
            closeImagePopover();
            closeNotePopover();
            closeTablePopover();
            if (noteEditorOpen) {
                closeNoteEditorModal();
            }
            linkPopoverContext = resolveLinkPopoverContext();
            linkPopoverMode = linkPopoverContext ? 'edit' : 'insert';
            refreshLinkPopoverActions();

            const selectedText = linkPopoverContext && typeof linkPopoverContext.label === 'string'
                ? linkPopoverContext.label.trim()
                : getSelectedTextFromCurrentContext();
            const href = linkPopoverContext && typeof linkPopoverContext.href === 'string'
                ? linkPopoverContext.href.trim()
                : t('linkDefaultTarget');
            const hrefParts = splitLinkHref(href);

            if (linkTextInput) {
                linkTextInput.value = selectedText || t('linkDefaultText');
            }
            if (linkTargetInput) {
                linkTargetInput.value = hrefParts.target || (hrefParts.anchor ? '' : href);
            }
            setLinkAnchorOptions([], hrefParts.anchor);
            setLinkPopoverState(true);
            window.requestAnimationFrame(() => {
                if (linkTargetInput) {
                    linkTargetInput.focus();
                    linkTargetInput.select();
                }
            });
        };

        const openImagePopover = () => {
            closeOpenMenus();
            closeLinkPopover();
            closeNotePopover();
            closeTablePopover();
            if (noteEditorOpen) {
                closeNoteEditorModal();
            }
            if (imageAltInput) {
                imageAltInput.value = t('imageDefaultAlt');
            }
            if (imageSrcInput) {
                imageSrcInput.value = t('imageDefaultSource');
            }
            if (imageTitleInput) {
                imageTitleInput.value = '';
            }
            setImagePopoverState(true);
            window.requestAnimationFrame(() => {
                if (imageSrcInput) {
                    imageSrcInput.focus();
                    imageSrcInput.select();
                }
            });
        };

        const openTablePopover = () => {
            closeOpenMenus();
            closeLinkPopover();
            closeImagePopover();
            closeNotePopover();
            if (noteEditorOpen) {
                closeNoteEditorModal();
            }
            setTablePopoverState(true);
            window.requestAnimationFrame(() => {
                if (tableColumnsInput) {
                    tableColumnsInput.focus();
                    tableColumnsInput.select();
                }
            });
        };

        const toggleTablePopover = () => {
            if (tablePopoverOpen) {
                closeTablePopover();
                return;
            }
            openTablePopover();
        };

        const toggleLinkPopover = () => {
            if (linkPopoverOpen) {
                closeLinkPopover();
                return;
            }
            openLinkPopover();
        };

        const toggleImagePopover = () => {
            if (imagePopoverOpen) {
                closeImagePopover();
                return;
            }
            openImagePopover();
        };

        const openNotePopover = () => {
            closeOpenMenus();
            closeLinkPopover();
            closeImagePopover();
            closeTablePopover();
            if (noteEditorOpen) {
                closeNoteEditorModal();
            }

            notePopoverContext = resolveNotePopoverContext();
            const contextEntry = notePopoverContext ? getNoteEntryById(notePopoverContext.noteId) : null;

            selectedNoteEntryId = contextEntry ? contextEntry.id : '';
            if (noteSearchInput) {
                noteSearchInput.value = '';
            }
            renderNoteEntriesList({ selectedId: selectedNoteEntryId, searchValue: '' });
            if (noteUnbindButton) {
                noteUnbindButton.hidden = !notePopoverContext;
            }

            setNotePopoverState(true);
            window.requestAnimationFrame(() => {
                if (noteSearchInput) {
                    noteSearchInput.focus();
                    noteSearchInput.select();
                }
            });
        };

        const toggleNotePopover = () => {
            if (notePopoverOpen) {
                closeNotePopover();
                return;
            }
            openNotePopover();
        };

        const insertGeneratedTable = () => {
            const safeColumns = parseIntegerInput(tableColumnsInput ? tableColumnsInput.value : 3, 1, 12, 3);
            const safeRows = parseIntegerInput(tableRowsInput ? tableRowsInput.value : 3, 1, 50, 3);
            const alignment = tableAlignSelect && tableAlignSelect.value
                ? tableAlignSelect.value
                : 'left';
            const includeHeader = tableHeaderCheckbox ? tableHeaderCheckbox.checked === true : true;
            if (tableColumnsInput) {
                tableColumnsInput.value = String(safeColumns);
            }
            if (tableRowsInput) {
                tableRowsInput.value = String(safeRows);
            }

            const headerPrefix = t('tableColumnPrefix');
            const snippet = buildMarkdownTable({
                columns: safeColumns,
                rows: safeRows,
                alignment,
                includeHeader,
                headerPrefix
            });
            const initialHeaderText = `${headerPrefix} 1`;
            const selectionInSnippet = includeHeader
                ? { start: 2, end: 2 + initialHeaderText.length }
                : null;

            const tryInsertTableIntoPreview = () => {
                if (!previewEditMode) {
                    return false;
                }
                const output = getPreviewOutputElement();
                if (!output) {
                    return false;
                }

                const renderedHtml = marked.parse(snippet, {
                    headerIds: false,
                    mangle: false
                });
                const sanitizedTableHtml = sanitizeRenderedMarkdown(renderedHtml);
                const tempContainer = document.createElement('div');
                tempContainer.innerHTML = sanitizedTableHtml;
                const tableElement = tempContainer.querySelector('table');
                if (!(tableElement instanceof HTMLTableElement)) {
                    return false;
                }

                output.focus();
                restorePreviewSelectionRange();
                const selection = window.getSelection();
                if (!selection) {
                    return false;
                }
                if (selection.rangeCount === 0) {
                    const fallbackRange = document.createRange();
                    fallbackRange.selectNodeContents(output);
                    fallbackRange.collapse(false);
                    selection.removeAllRanges();
                    selection.addRange(fallbackRange);
                    previewSavedRange = fallbackRange.cloneRange();
                }

                const selectedRange = selection.rangeCount > 0 ? selection.getRangeAt(0) : null;
                if (!selectedRange) {
                    return false;
                }
                const insertionRange = selectedRange.cloneRange();
                if (!isRangeInsideElement(insertionRange, output)) {
                    insertionRange.selectNodeContents(output);
                    insertionRange.collapse(false);
                }

                insertionRange.deleteContents();
                insertionRange.insertNode(tableElement);
                const firstCell = tableElement.querySelector('th,td');
                movePreviewCaretToCell(firstCell || tableElement);
                return true;
            };

            if (tryInsertTableIntoPreview()) {
                schedulePreviewSyncFromToolbar(editor);
                closeTablePopover();
                return;
            }

            editor.pushUndoStop();
            insertMarkdownSnippet(editor, snippet, selectionInSnippet);
            editor.pushUndoStop();
            closeTablePopover();
        };

        const applyTableColumnAction = (action) => {
            editor.pushUndoStop();
            const result = modifyTableColumn(editor, action);
            editor.pushUndoStop();
            if (!result.ok) {
                if (result.reason === 'min_columns') {
                    window.alert(t('tableMinColumns'));
                    return;
                }
                window.alert(t('tableNotFound'));
                return;
            }
            if (previewEditMode) {
                schedulePreviewSyncFromToolbar(editor);
            }
        };

        const applyTableRowAction = (action) => {
            editor.pushUndoStop();
            const result = modifyTableRow(editor, action);
            editor.pushUndoStop();
            if (!result.ok) {
                if (result.reason === 'no_data_rows') {
                    window.alert(t('tableNoDataRows'));
                    return;
                }
                window.alert(t('tableNotFound'));
                return;
            }
            if (previewEditMode) {
                schedulePreviewSyncFromToolbar(editor);
            }
        };

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
            link: () => toggleLinkPopover(),
            image: () => toggleImagePopover(),
            table: () => toggleTablePopover(),
            note: () => toggleNotePopover()
        };

        const clearButtonActiveStates = () => {
            buttonByFormat.forEach((_, formatType) => {
                if (formatType === 'table' || formatType === 'link' || formatType === 'image' || formatType === 'note') {
                    const isOpen = (formatType === 'table' && tablePopoverOpen)
                        || (formatType === 'link' && linkPopoverOpen)
                        || (formatType === 'image' && imagePopoverOpen)
                        || (formatType === 'note' && notePopoverOpen);
                    setButtonActiveState(formatType, isOpen);
                    return;
                }
                setButtonActiveState(formatType, false);
            });
        };

        const isCursorInsideDelimitedText = (lineText, cursorOffset, prefix, suffix) => {
            if (!lineText || !prefix || !suffix || cursorOffset < 0) {
                return false;
            }

            let openIndex = lineText.lastIndexOf(prefix, cursorOffset - 1);
            while (openIndex !== -1) {
                const contentStart = openIndex + prefix.length;
                const closeIndex = lineText.indexOf(suffix, contentStart);
                if (closeIndex !== -1 && closeIndex > contentStart && cursorOffset >= contentStart && cursorOffset <= closeIndex) {
                    return true;
                }
                openIndex = lineText.lastIndexOf(prefix, openIndex - 1);
            }
            return false;
        };

        const isInlineWrapActiveInEditor = (model, selection, prefix, suffix) => {
            if (!model || !selection) {
                return false;
            }

            const fullText = model.getValue();
            const startOffset = model.getOffsetAt({
                lineNumber: selection.startLineNumber,
                column: selection.startColumn
            });
            const endOffset = model.getOffsetAt({
                lineNumber: selection.endLineNumber,
                column: selection.endColumn
            });
            const selectedText = model.getValueInRange(selection);
            const hasSelection = startOffset !== endOffset;

            if (hasSelection) {
                return hasExactDelimiterAround(fullText, startOffset, endOffset, prefix, suffix)
                    || hasExactDelimiterInside(selectedText, prefix, suffix);
            }

            const lineText = model.getLineContent(selection.startLineNumber);
            const lineStartOffset = model.getOffsetAt({
                lineNumber: selection.startLineNumber,
                column: 1
            });
            const cursorOffsetInLine = startOffset - lineStartOffset;
            return isCursorInsideDelimitedText(lineText, cursorOffsetInLine, prefix, suffix);
        };

        const isLineFormatActiveInEditor = (model, selection, regex) => {
            if (!model || !selection || !(regex instanceof RegExp)) {
                return false;
            }
            let startLine = selection.startLineNumber;
            let endLine = selection.endLineNumber;
            if (endLine > startLine && selection.endColumn === 1) {
                endLine -= 1;
            }
            for (let line = startLine; line <= endLine; line += 1) {
                if (!regex.test(model.getLineContent(line))) {
                    return false;
                }
            }
            return true;
        };

        const isLinkActiveInEditor = (model, selection) => {
            if (!model || !selection) {
                return false;
            }
            const lineNumber = selection.startLineNumber;
            if (lineNumber !== selection.endLineNumber) {
                return false;
            }
            const lineText = model.getLineContent(lineNumber);
            const lineStartOffset = model.getOffsetAt({ lineNumber, column: 1 });
            const offsets = getSelectionOffsets(model, selection);
            if (!offsets) {
                return false;
            }
            const startInLine = offsets.startOffset - lineStartOffset;
            const endInLine = offsets.endOffset - lineStartOffset;
            return Boolean(findMarkdownLinkInLine(lineText, startInLine, endInLine));
        };

        const getPreviewSelectionContextElement = () => {
            const output = getPreviewOutputElement();
            if (!output) {
                return null;
            }

            let sourceRange = null;
            const selection = window.getSelection();
            if (selection && selection.rangeCount > 0) {
                const range = selection.getRangeAt(0);
                if (isRangeInsideElement(range, output)) {
                    sourceRange = range;
                }
            }
            if (!sourceRange && previewSavedRange && isRangeInsideElement(previewSavedRange, output)) {
                sourceRange = previewSavedRange;
            }
            if (!sourceRange) {
                return null;
            }

            const containerNode = sourceRange.startContainer;
            if (containerNode.nodeType === Node.ELEMENT_NODE) {
                return containerNode;
            }
            return containerNode.parentElement;
        };

        const getEditorFormatStates = () => {
            const model = editor.getModel();
            const selection = editor.getSelection();
            if (!model || !selection) {
                return {
                    link: linkPopoverOpen,
                    image: imagePopoverOpen,
                    table: tablePopoverOpen,
                    note: notePopoverOpen
                };
            }

            return {
                bold: isInlineWrapActiveInEditor(model, selection, '**', '**'),
                italic: isInlineWrapActiveInEditor(model, selection, '*', '*'),
                strikethrough: isInlineWrapActiveInEditor(model, selection, '~~', '~~'),
                code: isInlineWrapActiveInEditor(model, selection, '`', '`'),
                h1: isLineFormatActiveInEditor(model, selection, /^(\s{0,3})#\s+/),
                h2: isLineFormatActiveInEditor(model, selection, /^(\s{0,3})##\s+/),
                ul: isLineFormatActiveInEditor(model, selection, /^\s*[-*+]\s+/),
                ol: isLineFormatActiveInEditor(model, selection, /^\s*\d+\.\s+/),
                quote: isLineFormatActiveInEditor(model, selection, /^(\s{0,3})>\s?/),
                link: linkPopoverOpen || isLinkActiveInEditor(model, selection),
                image: imagePopoverOpen,
                table: tablePopoverOpen,
                note: notePopoverOpen
            };
        };

        const getPreviewFormatStates = () => {
            const contextElement = getPreviewSelectionContextElement();
            if (!contextElement || !(contextElement instanceof Element)) {
                return {
                    link: linkPopoverOpen,
                    image: imagePopoverOpen,
                    table: tablePopoverOpen,
                    note: notePopoverOpen
                };
            }
            const hasClosest = (selector) => Boolean(contextElement.closest(selector));
            const inInlineCode = hasClosest('code') && !hasClosest('pre code');

            return {
                bold: hasClosest('strong, b'),
                italic: hasClosest('em, i'),
                strikethrough: hasClosest('s, strike, del'),
                code: inInlineCode,
                h1: hasClosest('h1'),
                h2: hasClosest('h2'),
                ul: hasClosest('ul'),
                ol: hasClosest('ol'),
                quote: hasClosest('blockquote'),
                link: linkPopoverOpen || hasClosest('a[href]'),
                image: imagePopoverOpen || hasClosest('img'),
                table: tablePopoverOpen,
                note: notePopoverOpen || hasClosest('[data-lectr-note-id]')
            };
        };

        const applyFormatStates = (states) => {
            Object.keys(handlers).forEach((formatType) => {
                setButtonActiveState(formatType, Boolean(states[formatType]));
            });
        };

        let refreshFormatToolbarFrameId = null;
        const refreshFormatToolbarStateNow = () => {
            if (previewEditMode) {
                applyFormatStates(getPreviewFormatStates());
                return;
            }
            applyFormatStates(getEditorFormatStates());
        };

        const scheduleFormatToolbarStateRefresh = () => {
            if (refreshFormatToolbarFrameId !== null) {
                return;
            }
            refreshFormatToolbarFrameId = window.requestAnimationFrame(() => {
                refreshFormatToolbarFrameId = null;
                refreshFormatToolbarStateNow();
            });
        };

        refreshFormatToolbarState = scheduleFormatToolbarStateRefresh;
        scheduleFormatToolbarStateRefresh();

        const withUndoStop = (handler) => {
            editor.pushUndoStop();
            handler();
            editor.pushUndoStop();
            scheduleFormatToolbarStateRefresh();
        };
        refreshLinkPopoverActions();
        refreshNoteEditorHeader();

        if (tableInsertButton) {
            tableInsertButton.addEventListener('click', (event) => {
                event.preventDefault();
                insertGeneratedTable();
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (linkInsertButton) {
            linkInsertButton.addEventListener('click', (event) => {
                event.preventDefault();
                insertLinkFromPopover();
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (linkRemoveButton) {
            linkRemoveButton.addEventListener('click', (event) => {
                event.preventDefault();
                removeLinkFromPopover();
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (linkBrowseButton) {
            linkBrowseButton.addEventListener('click', (event) => {
                event.preventDefault();
                void browseLinkFile().then(() => {
                    scheduleFormatToolbarStateRefresh();
                });
            });
        }

        if (linkTargetInput) {
            linkTargetInput.addEventListener('input', () => {
                setLinkAnchorOptions([]);
            });
        }

        if (noteSearchInput) {
            noteSearchInput.addEventListener('input', () => {
                renderNoteEntriesList({
                    selectedId: selectedNoteEntryId,
                    searchValue: noteSearchInput.value
                });
            });
        }

        if (noteEntriesList) {
            noteEntriesList.addEventListener('click', (event) => {
                const target = event.target;
                if (!(target instanceof Element)) {
                    return;
                }
                const editButton = target.closest('.note-entry-edit');
                if (editButton instanceof HTMLButtonElement) {
                    const entryId = String(editButton.getAttribute('data-note-id') || '').trim();
                    if (!entryId) {
                        return;
                    }
                    selectedNoteEntryId = entryId;
                    renderNoteEntriesList({
                        selectedId: selectedNoteEntryId,
                        searchValue: noteSearchInput ? noteSearchInput.value : ''
                    });
                    openNoteEditorModal({ mode: 'edit', entryId });
                    return;
                }

                const button = target.closest('.note-entry-item');
                if (!(button instanceof HTMLButtonElement)) {
                    return;
                }
                const entryId = String(button.getAttribute('data-note-id') || '').trim();
                if (!entryId) {
                    return;
                }
                selectedNoteEntryId = entryId;
                renderNoteEntriesList({
                    selectedId: selectedNoteEntryId,
                    searchValue: noteSearchInput ? noteSearchInput.value : ''
                });
            });
        }

        if (noteNewEntryButton) {
            noteNewEntryButton.addEventListener('click', (event) => {
                event.preventDefault();
                openNoteEditorModal({ mode: 'create' });
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (noteBindButton) {
            noteBindButton.addEventListener('click', (event) => {
                event.preventDefault();
                bindNoteFromPopover();
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (noteEditorSaveButton) {
            noteEditorSaveButton.addEventListener('click', (event) => {
                event.preventDefault();
                saveNoteEntryFromModal();
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (noteEditorCancelButton) {
            noteEditorCancelButton.addEventListener('click', (event) => {
                event.preventDefault();
                closeNoteEditorModal();
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (noteEditorCloseButton) {
            noteEditorCloseButton.addEventListener('click', (event) => {
                event.preventDefault();
                closeNoteEditorModal();
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (noteEditorDeleteButton) {
            noteEditorDeleteButton.addEventListener('click', (event) => {
                event.preventDefault();
                if (!deleteNoteArchiveEntry(noteEditorEntryId)) {
                    return;
                }
                closeNoteEditorModal();
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (noteUnbindButton) {
            noteUnbindButton.addEventListener('click', (event) => {
                event.preventDefault();
                removeNoteBindingFromPopover();
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (imageInsertButton) {
            imageInsertButton.addEventListener('click', (event) => {
                event.preventDefault();
                insertImageFromPopover();
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (imageBrowseButton) {
            imageBrowseButton.addEventListener('click', (event) => {
                event.preventDefault();
                void browseImageFile().then(() => {
                    scheduleFormatToolbarStateRefresh();
                });
            });
        }

        if (tableAddColumnButton) {
            tableAddColumnButton.addEventListener('click', (event) => {
                event.preventDefault();
                applyTableColumnAction('add');
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (tableRemoveColumnButton) {
            tableRemoveColumnButton.addEventListener('click', (event) => {
                event.preventDefault();
                applyTableColumnAction('remove');
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (tableAddRowButton) {
            tableAddRowButton.addEventListener('click', (event) => {
                event.preventDefault();
                applyTableRowAction('add');
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (tableRemoveRowButton) {
            tableRemoveRowButton.addEventListener('click', (event) => {
                event.preventDefault();
                applyTableRowAction('remove');
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (linkPopover) {
            linkPopover.addEventListener('keydown', (event) => {
                if (event.key !== 'Enter') {
                    return;
                }
                if (event.target instanceof HTMLButtonElement) {
                    return;
                }
                event.preventDefault();
                insertLinkFromPopover();
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (imagePopover) {
            imagePopover.addEventListener('keydown', (event) => {
                if (event.key !== 'Enter') {
                    return;
                }
                if (event.target instanceof HTMLButtonElement) {
                    return;
                }
                event.preventDefault();
                insertImageFromPopover();
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (tablePopover) {
            tablePopover.addEventListener('keydown', (event) => {
                if (event.key !== 'Enter') {
                    return;
                }
                if (event.target instanceof HTMLButtonElement) {
                    return;
                }
                event.preventDefault();
                insertGeneratedTable();
                scheduleFormatToolbarStateRefresh();
            });
        }

        if (noteEditorOverlay) {
            noteEditorOverlay.addEventListener('pointerdown', (event) => {
                if (event.target === noteEditorOverlay) {
                    closeNoteEditorModal();
                    scheduleFormatToolbarStateRefresh();
                }
            });
            noteEditorOverlay.addEventListener('keydown', (event) => {
                if (event.key === 'Escape') {
                    event.preventDefault();
                    closeNoteEditorModal();
                    scheduleFormatToolbarStateRefresh();
                    return;
                }
                if (event.key !== 'Enter') {
                    return;
                }
                if (event.target instanceof HTMLTextAreaElement || event.target instanceof HTMLButtonElement) {
                    return;
                }
                event.preventDefault();
                saveNoteEntryFromModal();
                scheduleFormatToolbarStateRefresh();
            });
        }

        document.addEventListener('pointerdown', (event) => {
            if (noteEditorOpen) {
                return;
            }
            if (!linkPopoverOpen && !imagePopoverOpen && !tablePopoverOpen && !notePopoverOpen) {
                return;
            }
            const target = event.target;
            const isElementTarget = target instanceof Element;
            const insideLink = isElementTarget && !!target.closest('#link-tool');
            const insideImage = isElementTarget && !!target.closest('#image-tool');
            const insideTable = isElementTarget && !!target.closest('#table-tool');
            const insideNote = isElementTarget && !!target.closest('#note-tool');
            let closedAny = false;
            if (linkPopoverOpen && !insideLink) {
                closeLinkPopover();
                closedAny = true;
            }
            if (imagePopoverOpen && !insideImage) {
                closeImagePopover();
                closedAny = true;
            }
            if (tablePopoverOpen && !insideTable) {
                closeTablePopover();
                closedAny = true;
            }
            if (notePopoverOpen && !insideNote) {
                closeNotePopover();
                closedAny = true;
            }
            if (closedAny) {
                scheduleFormatToolbarStateRefresh();
            }
        });

        document.addEventListener('keydown', (event) => {
            if (event.key !== 'Escape') {
                return;
            }
            if (noteEditorOpen) {
                closeNoteEditorModal();
                scheduleFormatToolbarStateRefresh();
                return;
            }
            if (!linkPopoverOpen && !imagePopoverOpen && !tablePopoverOpen && !notePopoverOpen) {
                return;
            }
            closeInlinePopovers();
            scheduleFormatToolbarStateRefresh();
        });

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

            if (formatType === 'table' || formatType === 'link' || formatType === 'image' || formatType === 'note') {
                handlers[formatType]();
                scheduleFormatToolbarStateRefresh();
                return;
            }

            if (previewEditMode && applyPreviewFormat(formatType, editor)) {
                scheduleFormatToolbarStateRefresh();
                return;
            }
            withUndoStop(handlers[formatType]);
        });

        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyB, () => withUndoStop(handlers.bold));
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyI, () => withUndoStop(handlers.italic));
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyK, () => {
            handlers.link();
            scheduleFormatToolbarStateRefresh();
        });
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
                const activeTabElement = document.querySelector(`#tabs-list [data-tab-id="${activeTab.id}"]`);
                void closeTab(activeTab.id, editor, {
                    tabElement: activeTabElement,
                    animate: true
                });
            }
        });
        editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyT, () => {
            const newTabButton = document.querySelector('#new-tab-button');
            if (newTabButton instanceof HTMLButtonElement) {
                newTabButton.click();
                return;
            }
            createUntitledTab('');
        });

        editor.onDidChangeCursorSelection(() => {
            scheduleFormatToolbarStateRefresh();
        });
        editor.onDidChangeModelContent(() => {
            scheduleFormatToolbarStateRefresh();
        });
        editor.onDidFocusEditorText(() => {
            scheduleFormatToolbarStateRefresh();
        });
        editor.onDidBlurEditorText(() => {
            if (previewEditMode) {
                scheduleFormatToolbarStateRefresh();
                return;
            }
            clearButtonActiveStates();
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

    let loadOnboardingState = () => {
        try {
            const raw = localStorage.getItem(`${localStorageNamespace}_${localStorageOnboardingKey}`);
            return raw ? JSON.parse(raw) : null;
        } catch (error) {
            return null;
        }
    };

    let saveOnboardingState = (state) => {
        try {
            localStorage.setItem(`${localStorageNamespace}_${localStorageOnboardingKey}`, JSON.stringify(state));
        } catch (error) {
            // ignore storage errors
        }
    };

    let loadNotesArchive = () => {
        try {
            const raw = localStorage.getItem(`${localStorageNamespace}_${localStorageNotesArchiveKey}`);
            if (!raw) {
                return [];
            }
            const parsed = JSON.parse(raw);
            if (!Array.isArray(parsed)) {
                return [];
            }
            return parsed.map((entry) => {
                const id = typeof entry?.id === 'string' ? entry.id.trim() : '';
                const title = typeof entry?.title === 'string' ? entry.title.trim() : '';
                const body = typeof entry?.body === 'string' ? entry.body.trim() : '';
                if (!id || !title || !body) {
                    return null;
                }
                return {
                    id,
                    title,
                    body,
                    updatedAt: Number.isFinite(entry?.updatedAt) ? entry.updatedAt : Date.now()
                };
            }).filter(Boolean);
        } catch (_error) {
            return [];
        }
    };

    let saveNotesArchive = () => {
        try {
            localStorage.setItem(`${localStorageNamespace}_${localStorageNotesArchiveKey}`, JSON.stringify(notesArchive));
        } catch (_error) {
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
                    dirty: tab.dirty === true || content !== lastSavedContent,
                    editorScrollTop: Number.isFinite(tab.editorScrollTop) ? tab.editorScrollTop : 0,
                    previewScrollTop: Number.isFinite(tab.previewScrollTop) ? tab.previewScrollTop : 0
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
            captureActiveTabScrollState();
            const snapshot = {
                version: 1,
                activeTabId,
                tabs: tabs.map((tab) => ({
                    id: tab.id,
                    title: tab.title,
                    filePath: tab.filePath,
                    content: tab.content,
                    lastSavedContent: tab.lastSavedContent,
                    dirty: tab.dirty,
                    editorScrollTop: Number.isFinite(tab.editorScrollTop) ? tab.editorScrollTop : 0,
                    previewScrollTop: Number.isFinite(tab.previewScrollTop) ? tab.previewScrollTop : 0
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

    let initLanguageSetting = (language) => {
        const languageSelect = document.querySelector('#language-select');
        setLanguagePreference = (nextLanguage, { persist = true } = {}) => {
            const safeLanguage = nextLanguage === 'ru' ? 'ru' : 'en';
            currentLanguage = safeLanguage;
            if (languageSelect) {
                languageSelect.value = safeLanguage;
            }
            if (persist) {
                saveLanguageSettings(safeLanguage);
            }
            applyLocalization();
            refreshFormatToolbarState();
        };

        setLanguagePreference(language, { persist: false });

        if (!languageSelect || languageSelect.dataset.boundChange === '1') {
            return;
        }
        languageSelect.dataset.boundChange = '1';
        languageSelect.addEventListener('change', (event) => {
            const selected = event.currentTarget.value === 'ru' ? 'ru' : 'en';
            setLanguagePreference(selected, { persist: true });
        });
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
        setPreviewEditModePreference = (nextEnabled, {
            persist = true,
            syncRender = true,
            focusOutput = true
        } = {}) => {
            const checked = nextEnabled === true;
            previewEditMode = checked;

            if (checkbox) {
                checkbox.checked = checked;
            }

            if (persist) {
                savePreviewEditModeSetting(checked);
            }

            setPreviewEditableState(checked);
            setPreviewEditLayout(checked);
            refreshTopImmersionState();
            refreshFormatToolbarState();

            if (checked && focusOutput) {
                const output = getPreviewOutputElement();
                if (output) {
                    output.focus();
                }
            }

            if (previewEditSyncTimer !== null) {
                window.clearTimeout(previewEditSyncTimer);
                previewEditSyncTimer = null;
            }

            if (!syncRender) {
                return;
            }

            if (!checked) {
                const value = editor.getValue();
                scheduleConvert(value);
                return;
            }
            scheduleConvert(editor.getValue());
        };

        setPreviewEditModePreference(enabled, {
            persist: false,
            syncRender: false,
            focusOutput: false
        });

        if (!checkbox || checkbox.dataset.boundChange === '1') {
            return;
        }
        checkbox.dataset.boundChange = '1';
        checkbox.addEventListener('change', (event) => {
            const checked = event.currentTarget.checked === true;
            setPreviewEditModePreference(checked, {
                persist: true,
                syncRender: true,
                focusOutput: true
            });
        });
    };

    let setupFirstRunOnboarding = (editor) => {
        const overlay = document.querySelector('#onboarding-overlay');
        const stepLabel = document.querySelector('#onboarding-step-label');
        const backButton = document.querySelector('#onboarding-back-button');
        const nextButton = document.querySelector('#onboarding-next-button');
        const stepElements = Array.from(document.querySelectorAll('[data-onboarding-step]'));
        const languageButtons = Array.from(document.querySelectorAll('[data-onboarding-language]'));
        const modeButtons = Array.from(document.querySelectorAll('[data-onboarding-mode]'));
        const themeButtons = Array.from(document.querySelectorAll('[data-onboarding-theme]'));
        if (!overlay || !stepLabel || !backButton || !nextButton || stepElements.length === 0) {
            return;
        }

        const onboardingState = loadOnboardingState();
        if (onboardingState && onboardingState.completed === true) {
            return;
        }

        const state = {
            language: currentLanguage === 'ru' ? 'ru' : 'en',
            mode: previewEditMode ? 'simple' : 'advanced',
            theme: document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light'
        };
        let activeStep = 0;
        const totalSteps = stepElements.length;

        const renderStepButtons = () => {
            backButton.disabled = activeStep === 0;
            nextButton.textContent = activeStep === totalSteps - 1 ? t('onboardingFinish') : t('onboardingNext');
            stepLabel.textContent = t('onboardingStep', {
                step: activeStep + 1,
                total: totalSteps
            });
        };

        const renderSteps = () => {
            stepElements.forEach((stepElement, index) => {
                stepElement.classList.toggle('active', index === activeStep);
            });
            renderStepButtons();
        };

        const markActiveChoice = (buttons, key, expectedValue) => {
            buttons.forEach((button) => {
                const currentValue = button.getAttribute(key);
                button.classList.toggle('active', currentValue === expectedValue);
            });
        };

        const renderChoices = () => {
            markActiveChoice(languageButtons, 'data-onboarding-language', state.language);
            markActiveChoice(modeButtons, 'data-onboarding-mode', state.mode);
            markActiveChoice(themeButtons, 'data-onboarding-theme', state.theme);
        };

        const closeOverlay = () => {
            overlay.classList.remove('open');
            overlay.setAttribute('aria-hidden', 'true');
            document.body.classList.remove('onboarding-open');
            window.setTimeout(() => {
                overlay.hidden = true;
            }, 210);
        };

        const finishOnboarding = () => {
            setLanguagePreference(state.language, { persist: true });
            setPreviewEditModePreference(state.mode === 'simple', {
                persist: true,
                syncRender: true,
                focusOutput: state.mode === 'simple'
            });
            applyThemePreference(state.theme === 'dark', { persist: true });
            saveOnboardingState({
                completed: true,
                language: state.language,
                mode: state.mode,
                theme: state.theme,
                completedAt: Date.now()
            });
            closeOverlay();
        };

        refreshOnboardingLocalization = () => {
            if (overlay.hidden) {
                return;
            }
            renderStepButtons();
        };

        languageButtons.forEach((button) => {
            button.addEventListener('click', () => {
                const nextLanguage = button.getAttribute('data-onboarding-language') === 'ru' ? 'ru' : 'en';
                state.language = nextLanguage;
                renderChoices();
                setLanguagePreference(nextLanguage, { persist: false });
            });
        });

        modeButtons.forEach((button) => {
            button.addEventListener('click', () => {
                const nextMode = button.getAttribute('data-onboarding-mode') === 'simple' ? 'simple' : 'advanced';
                state.mode = nextMode;
                renderChoices();
                setPreviewEditModePreference(nextMode === 'simple', {
                    persist: false,
                    syncRender: true,
                    focusOutput: false
                });
            });
        });

        themeButtons.forEach((button) => {
            button.addEventListener('click', () => {
                const nextTheme = button.getAttribute('data-onboarding-theme') === 'dark' ? 'dark' : 'light';
                state.theme = nextTheme;
                renderChoices();
                applyThemePreference(nextTheme === 'dark', { persist: false });
            });
        });

        backButton.addEventListener('click', () => {
            activeStep = Math.max(0, activeStep - 1);
            renderSteps();
        });

        nextButton.addEventListener('click', () => {
            if (activeStep >= totalSteps - 1) {
                finishOnboarding();
                return;
            }
            activeStep += 1;
            renderSteps();
        });

        closeOpenMenus();
        overlay.hidden = false;
        overlay.setAttribute('aria-hidden', 'false');
        document.body.classList.add('onboarding-open');
        window.requestAnimationFrame(() => {
            overlay.classList.add('open');
        });
        renderChoices();
        renderSteps();
    };

    let setupDivider = () => {
        let lastLeftRatio = 0.5;
        const divider = document.getElementById('split-divider');
        const leftPane = document.getElementById('edit');
        const rightPane = document.getElementById('preview');
        const container = document.getElementById('container');
        if (!divider || !leftPane || !rightPane || !container) {
            return;
        }

        let isDragging = false;
        const stopDragging = () => {
            if (!isDragging) {
                return;
            }
            isDragging = false;
            divider.classList.remove('active');
            divider.classList.remove('hover');
            document.body.style.cursor = 'default';
            document.body.style.userSelect = '';
        };

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
            stopDragging();
        });

        window.addEventListener('blur', () => {
            stopDragging();
        });

        document.addEventListener('visibilitychange', () => {
            if (document.hidden) {
                stopDragging();
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
    notesArchive = loadNotesArchive();
    let editor = await setupEditor();
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
                    : (savedTab.content || ''),
                editorScrollTop: Number.isFinite(savedTab.editorScrollTop) ? savedTab.editorScrollTop : 0,
                previewScrollTop: Number.isFinite(savedTab.previewScrollTop) ? savedTab.previewScrollTop : 0
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
        window.requestAnimationFrame(() => {
            applyTabScrollState(activeTab);
        });
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
        window.requestAnimationFrame(() => {
            applyTabScrollState(initialTab);
        });
    }

    setupHeaderMenus();
    setupAboutDialog();
    setupOpenButton();
    setupPreviewLinkNavigation();
    setupPreviewNoteInteractions();
    setupPreviewTableActions(editor);
    setupSaveButton(editor);
    setupSaveAsButton(editor);
    setupResetButton();
    setupCopyButton(editor);
    setupExportButton();
    setupFormatToolbar(editor);
    setupGlobalShortcuts(editor);

    // initialize theme (dark/light)
    let themeSettings = loadThemeSettings();
    // normalize to boolean (saved value may be string or boolean)
    if (themeSettings === 'true' || themeSettings === true) {
        themeSettings = true;
    } else {
        themeSettings = false;
    }
    initThemeToggle(themeSettings);
    setupFirstRunOnboarding(editor);

    setupDivider();
    applyEditorViewportTopInset(editor);
    window.addEventListener('resize', () => {
        applyEditorViewportTopInset(editor);
    });

    window.addEventListener('beforeunload', () => {
        if (previewEditSyncTimer !== null) {
            window.clearTimeout(previewEditSyncTimer);
            previewEditSyncTimer = null;
        }
        if (persistTabsTimer !== null) {
            window.clearTimeout(persistTabsTimer);
            persistTabsTimer = null;
        }
        hideNoteTooltip();
        saveTabsState();
    });
};

window.addEventListener('load', () => {
    init().catch((error) => {
        // eslint-disable-next-line no-console
        console.error('Failed to initialize app', error);
    });
});
