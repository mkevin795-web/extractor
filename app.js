(() => {
    function bootstrap() {
        const collectDom = () => ({
            dropzone: document.getElementById('dropzone'),
            fileInput: document.getElementById('fileInput'),
            pickerInput: document.getElementById('picker'),
            manageUsersBtn: document.getElementById('manageUsersBtn'),
            pickerMeta: document.getElementById('pickerMeta'),
            pickerManager: document.getElementById('pickerManager'),
            pickerManagerCount: document.getElementById('pickerManagerCount'),
            newPickerInput: document.getElementById('newPickerInput'),
            addPickerBtn: document.getElementById('addPickerBtn'),
            pickerList: document.getElementById('pickerList'),
            pickerListTotal: document.getElementById('pickerListTotal'),
            resultsPanel: document.getElementById('results'),
            tableBody: document.querySelector('#dataTable tbody'),
            flexStatBtn: document.getElementById('flexStatBtn'),
            colectaStatBtn: document.getElementById('colectaStatBtn'),
            totalStatBtn: document.getElementById('totalStatBtn'),
            countFlex: document.getElementById('countFlex'),
            countColecta: document.getElementById('countColecta'),
            countTotal: document.getElementById('countTotal'),
            downloadBtn: document.getElementById('downloadBtn'),
            clearBtn: document.getElementById('clearBtn'),
            textInput: document.getElementById('textInput'),
            processTextBtn: document.getElementById('processTextBtn'),
            printZebraBtn: document.getElementById('printZebraBtn'),
            searchInput: document.getElementById('searchInput'),
            selectedInfo: document.getElementById('selectedInfo'),
            resultsMeta: document.getElementById('resultsMeta'),
            storageStatus: document.getElementById('storageStatus'),
            outsideHoursToggleBtn: document.getElementById('outsideHoursToggleBtn'),
            storageSummary: document.getElementById('storageSummary'),
            storageNote: document.getElementById('storageNote'),
            historyList: document.getElementById('historyList'),
            showStoredBtn: document.getElementById('showStoredBtn'),
            clearStoredBtn: document.getElementById('clearStoredBtn'),
            savePortableBtn: document.getElementById('savePortableBtn'),
            messageStack: document.getElementById('messageStack'),
            orderLookupFileInput: document.getElementById('orderLookupFileInput'),
            orderLookupLoadBtn: document.getElementById('orderLookupLoadBtn'),
            orderLookupFileName: document.getElementById('orderLookupFileName'),
            orderLookupInput: document.getElementById('orderLookupInput'),
            orderLookupSearchBtn: document.getElementById('orderLookupSearchBtn'),
            orderLookupStateFilter: document.getElementById('orderLookupStateFilter'),
            orderLookupRouteFilter: document.getElementById('orderLookupRouteFilter'),
            orderLookupClearFiltersBtn: document.getElementById('orderLookupClearFiltersBtn'),
            orderLookupSummary: document.getElementById('orderLookupSummary'),
            orderLookupTableWrap: document.getElementById('orderLookupTableWrap'),
            orderLookupTableHead: document.getElementById('orderLookupTableHead'),
            orderLookupTableBody: document.getElementById('orderLookupTableBody'),
            pasteTitle: document.querySelector('.input-methods .method-card:last-child h3')
        });

        const dom = collectDom();
        if (!dom.dropzone || !dom.fileInput || !dom.pickerInput || !dom.tableBody) {
            return;
        }

        const App = {
            config: {
                STORAGE_KEY: 'novapet_meli_documents_v2',
                PICKERS_STORAGE_KEY: 'novapet_meli_picker_options_v1',
                OUTSIDE_HOURS_HISTORY_KEY: 'novapet_meli_after_hours_history_v1',
                ORDER_LOOKUP_STORAGE_KEY: 'novapet_meli_order_lookup_v1',
                STORAGE_START_MINUTES: 7 * 60,
                STORAGE_END_MINUTES: 18 * 60,
                STORAGE_WINDOW_LABEL: '07:00 a 18:00',
                HISTORY_PREVIEW_LIMIT: 5,
                EMBEDDED_HISTORY_SCRIPT_ID: 'embeddedHistoryData',
                PORTABLE_FILENAME_PREFIX: 'ExtraccionMeli-portable',
                ORDER_LOOKUP_FIELDS: [
                    {
                        key: 'dato2',
                        label: 'DATO2',
                        required: true,
                        searchable: true,
                        aliases: ['DATO2']
                    },
                    {
                        key: 'estado',
                        label: 'ESTADO',
                        required: true,
                        aliases: ['ESTADO', 'ESTADOREV', 'ESTADOREVISION', 'ESTADOPEDIDO']
                    },
                    {
                        key: 'logisticsType',
                        label: 'RUTAS',
                        aliases: []
                    },
                    {
                        key: 'referenceNumber',
                        label: 'NRO.',
                        searchable: true,
                        aliases: ['NRO', 'DOCUMENTODEREFERENCIA', 'NROREF', 'NROREFERENCIA', 'PEDIDO', 'ORDEN']
                    },
                    {
                        key: 'buyerOrder',
                        label: 'COMPRADOR',
                        searchable: true,
                        aliases: ['COMPRADOR']
                    },
                    {
                        key: 'orderSource',
                        label: 'TIPO SOLICITUD',
                        aliases: ['TIPOSOLICITUD']
                    },
                    {
                        key: 'dispatchRoute',
                        label: 'RUTA DESPACHO',
                        aliases: ['RUTADESPACHO', 'RUTA']
                    }
                ]
            },

            dom,

            runtime: {
                storage: null,
                defaultPickerOptions: [],
                clockTimeFormatter: new Intl.DateTimeFormat('es-CL', {
                    hour: '2-digit',
                    minute: '2-digit',
                    second: '2-digit',
                    hour12: false
                }),
                dateTimeFormatter: new Intl.DateTimeFormat('es-CL', {
                    day: '2-digit',
                    month: '2-digit',
                    year: 'numeric',
                    hour: '2-digit',
                    minute: '2-digit',
                    hour12: false
                })
            },

            defaults: {
                selectionMessage: '',
                emptyResultsMessage: 'Carga un archivo, pega texto o usa la busqueda para consultar documentos guardados.'
            },

            state: {
                pickerOptions: [],
                embeddedDocuments: [],
                storedDocuments: [],
                extractedRows: [],
                rawZplData: '',
                selectedRowIds: new Set(),
                selectedRowsById: new Map(),
                activeBaseRows: [],
                activeBaseMessage: '',
                currentResultRows: [],
                currentResultMessage: '',
                activeRowTypeFilter: 'all',
                selectedHistoryDocumentIds: new Set(),
                outsideHoursHistoryEnabled: false,
                orderLookupRows: [],
                orderLookupDisplayColumns: [],
                orderLookupFileName: '',
                orderLookupMatches: [],
                orderLookupAvailableStates: [],
                orderLookupAvailableRoutes: []
            },

            helpers: {
                escapeHtml(value) {
                    return String(value ?? '')
                        .replace(/&/g, '&amp;')
                        .replace(/</g, '&lt;')
                        .replace(/>/g, '&gt;')
                        .replace(/"/g, '&quot;')
                        .replace(/'/g, '&#39;');
                },

                normalizePickerName(value) {
                    return String(value ?? '').trim();
                },

                normalizeComparable(value) {
                    const normalized = String(value ?? '').trim();
                    return /^-?\d+(\.0+)?$/.test(normalized)
                        ? normalized.replace(/\.0+$/, '')
                        : normalized;
                },

                normalizeHeaderKey(value) {
                    return App.helpers
                        .normalizeComparable(value)
                        .normalize('NFD')
                        .replace(/[\u0300-\u036f]/g, '')
                        .toUpperCase()
                        .replace(/[^A-Z0-9]+/g, '');
                },

                normalizePickerOptions(options, fallback = App.runtime.defaultPickerOptions) {
                    const source = Array.isArray(options) && options.length > 0 ? options : fallback;
                    const seen = new Set();

                    return source
                        .map(App.helpers.normalizePickerName)
                        .filter(option => {
                            const key = option.toLowerCase();
                            if (!option || seen.has(key)) {
                                return false;
                            }

                            seen.add(key);
                            return true;
                        });
                },

                summarizeDocumentType(rows = []) {
                    const counts = rows.reduce((accumulator, row) => {
                        if (row?.type === 'flex') {
                            accumulator.flex += 1;
                        } else if (row?.type === 'colecta') {
                            accumulator.colecta += 1;
                        } else {
                            accumulator.other += 1;
                        }

                        return accumulator;
                    }, { flex: 0, colecta: 0, other: 0 });

                    if (counts.flex > 0 && counts.colecta === 0 && counts.other === 0) {
                        return { key: 'flex', label: 'FLEX', detail: `${counts.flex} pedido(s)` };
                    }

                    if (counts.colecta > 0 && counts.flex === 0 && counts.other === 0) {
                        return { key: 'colecta', label: 'COLECTA', detail: `${counts.colecta} pedido(s)` };
                    }

                    if (counts.flex > 0 || counts.colecta > 0) {
                        return { key: 'mixto', label: 'MIXTO', detail: `Flex ${counts.flex} / Colecta ${counts.colecta}` };
                    }

                    return { key: 'sin-tipo', label: 'SIN TIPO', detail: `${rows.length} pedido(s)` };
                }
            },

            ui: {
                normalizeMessageType(type = 'info') {
                    return ['error', 'success', 'warning', 'info'].includes(type) ? type : 'info';
                },

                getMessageTimeout(type = 'info') {
                    const normalizedType = App.ui.normalizeMessageType(type);
                    if (normalizedType === 'error') {
                        return 5000;
                    }

                    if (normalizedType === 'success') {
                        return 3800;
                    }

                    if (normalizedType === 'warning') {
                        return 4200;
                    }

                    return 3500;
                },

                dismissMessage(item, options = {}) {
                    if (!item) {
                        return;
                    }

                    const { immediate = false } = options;
                    const removeItem = () => {
                        if (item.parentNode) {
                            item.remove();
                        }
                    };

                    if (item._messageTimer) {
                        window.clearTimeout(item._messageTimer);
                        item._messageTimer = null;
                    }

                    if (item._messageCloseTimer) {
                        window.clearTimeout(item._messageCloseTimer);
                        item._messageCloseTimer = null;
                    }

                    if (immediate) {
                        removeItem();
                        return;
                    }

                    if (item.dataset.toastState === 'leaving') {
                        return;
                    }

                    item.dataset.toastState = 'leaving';
                    item.classList.remove('is-visible');
                    item.classList.add('is-leaving');
                    item._messageCloseTimer = window.setTimeout(removeItem, 240);
                },

                showToast(message, type = 'info', options = {}) {
                    const stack = App.dom.messageStack;
                    if (!stack || !message) {
                        return;
                    }

                    const {
                        title = '',
                        sticky = false,
                        timeout = App.ui.getMessageTimeout(type)
                    } = options;
                    const normalizedType = App.ui.normalizeMessageType(type);

                    const item = document.createElement('article');
                    item.className = `app-message is-${normalizedType}`;
                    item.setAttribute('role', normalizedType === 'error' ? 'alert' : 'status');
                    item.dataset.toastState = 'entering';

                    const content = document.createElement('div');
                    content.className = 'app-message-content';

                    if (title) {
                        const titleNode = document.createElement('strong');
                        titleNode.textContent = title;
                        content.appendChild(titleNode);
                    }

                    const textNode = document.createElement('div');
                    textNode.className = 'app-message-text';
                    textNode.textContent = String(message);
                    content.appendChild(textNode);

                    const closeButton = document.createElement('button');
                    closeButton.type = 'button';
                    closeButton.className = 'app-message-close';
                    closeButton.setAttribute('aria-label', 'Cerrar mensaje');
                    closeButton.innerHTML = '&times;';
                    closeButton.addEventListener('click', () => App.ui.dismissMessage(item));

                    item.append(content, closeButton);
                    stack.prepend(item);
                    window.requestAnimationFrame(() => {
                        item.dataset.toastState = 'visible';
                        item.classList.add('is-visible');
                    });

                    while (stack.children.length > 5) {
                        App.ui.dismissMessage(stack.lastElementChild, { immediate: true });
                    }

                    if (!sticky) {
                        item._messageTimer = window.setTimeout(() => {
                            App.ui.dismissMessage(item);
                        }, timeout);
                    }
                },

                showMessage(message, type = 'info', options = {}) {
                    App.ui.showToast(message, type, options);
                },

                reportError(error, fallbackMessage, title = 'Error') {
                    console.error(error);
                    App.ui.showMessage(error?.message || fallbackMessage, 'error', { title });
                },

                updateResultsMeta(message) {
                    App.dom.resultsMeta.textContent = message || '';
                },

                setResultsVisible(visible) {
                    if (!App.dom.resultsPanel) {
                        return;
                    }

                    App.dom.resultsPanel.style.display = 'block';
                    App.dom.resultsPanel.dataset.resultsState = visible ? 'active' : 'idle';
                },

                setClearButtonVisible(visible) {
                    App.dom.clearBtn.classList.toggle('is-hidden', !visible);
                }
            },

            parser: {
                hashString(value) {
                    let hash = 2166136261;

                    for (let index = 0; index < value.length; index += 1) {
                        hash ^= value.charCodeAt(index);
                        hash += (hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24);
                    }

                    return (hash >>> 0).toString(16).padStart(8, '0');
                },

                buildDocumentFingerprint(rows) {
                    if (!Array.isArray(rows) || rows.length === 0) {
                        return '';
                    }

                    const signature = rows
                        .map(row => [
                            row.numero || '',
                            row.type || '',
                            row.picking || '',
                            row.revision || '',
                            row.zpl || ''
                        ].join('|'))
                        .join('||');

                    return `doc-${rows.length}-${App.parser.hashString(signature)}`;
                },

                getDocumentFingerprint(documentItem) {
                    if (!documentItem) {
                        return '';
                    }

                    return documentItem.fingerprint || App.parser.buildDocumentFingerprint(documentItem.rows);
                },

                buildDuplicateFilesMessage(fileNames) {
                    if (fileNames.length === 1) {
                        return `El archivo "${fileNames[0]}" ya habia sido guardado anteriormente y no se volvio a cargar.`;
                    }

                    return `Los siguientes archivos ya habian sido guardados anteriormente y no se volvieron a cargar:\n- ${fileNames.join('\n- ')}`;
                },

                readFileAsText(file) {
                    return new Promise((resolve, reject) => {
                        const reader = new FileReader();
                        reader.onload = event => resolve(String(event.target?.result || ''));
                        reader.onerror = () => reject(new Error(`No se pudo leer el archivo "${file.name}".`));
                        reader.readAsText(file);
                    });
                },

                extractBlocksFromText(text) {
                    const normalizedText = String(text ?? '').replace(/\r/g, '').trim();
                    if (!normalizedText) {
                        return [];
                    }

                    const rawBlocks = normalizedText.includes('^XA')
                        ? normalizedText.split(/(?=\^XA)/gi)
                        : normalizedText.split(/\^XZ/gi);

                    return rawBlocks
                        .map(block => block.trim())
                        .filter(Boolean)
                        .map(block => block.endsWith('^XZ') ? block : `${block}^XZ`)
                        .filter(block => /"id"\s*:\s*"?\d+"?/i.test(block) || /"shipping_id"\s*:\s*"?\d+"?/i.test(block));
                },

                buildRowFromBlock(block, picker) {
                    const match = /"id"\s*:\s*"?(?<id>\d+)"?/i.exec(block)
                        || /"shipping_id"\s*:\s*"?(?<id>\d+)"?/i.exec(block);

                    if (!match) {
                        return null;
                    }

                    const normalizedBlock = String(block ?? '').trim();
                    const isFlex = /\bflex\b/i.test(normalizedBlock) || /MESA05|REVISION05/i.test(normalizedBlock);

                    return {
                        proceso: 'INT-PICK-MERCADOLIBRE-MAS',
                        numero: match.groups?.id || match[1],
                        picker,
                        picking: isFlex ? 'MESA05_CD' : 'MESA03_CD',
                        revision: isFlex ? 'REVISION05_CD' : 'REVISION03_CD',
                        type: isFlex ? 'flex' : 'colecta',
                        zpl: normalizedBlock.endsWith('^XZ') ? normalizedBlock : `${normalizedBlock}^XZ`
                    };
                },

                extractRowsFromText(text, picker) {
                    return App.parser.extractBlocksFromText(text).reduce((rows, block) => {
                        const row = App.parser.buildRowFromBlock(block, picker);
                        if (row) {
                            rows.push(row);
                        }
                        return rows;
                    }, []);
                },

                createDocumentRecord({ picker, sourceName, sourceType, text }) {
                    const rows = App.parser.extractRowsFromText(text, picker);
                    if (rows.length === 0) {
                        return null;
                    }

                    return {
                        id: `doc-${Date.now()}-${Math.random().toString(16).slice(2, 10)}`,
                        picker,
                        sourceName,
                        sourceType,
                        storedAt: new Date().toISOString(),
                        fingerprint: App.parser.buildDocumentFingerprint(rows),
                        rows
                    };
                }
            },
            storage: {
                getBrowserStorage() {
                    try {
                        const testKey = '__novapet_storage_test__';
                        window.localStorage.setItem(testKey, '1');
                        window.localStorage.removeItem(testKey);
                        return window.localStorage;
                    } catch (error) {
                        return null;
                    }
                },

                loadEmbeddedPayload() {
                    const embeddedHistoryNode = document.getElementById(App.config.EMBEDDED_HISTORY_SCRIPT_ID);
                    if (!embeddedHistoryNode) {
                        return {};
                    }

                    try {
                        const parsed = JSON.parse(embeddedHistoryNode.textContent || '{}');
                        return parsed && typeof parsed === 'object' ? parsed : {};
                    } catch (error) {
                        return {};
                    }
                },

                loadStoredPickerOptions() {
                    if (!App.runtime.storage) {
                        return [];
                    }

                    try {
                        const parsed = JSON.parse(App.runtime.storage.getItem(App.config.PICKERS_STORAGE_KEY) || '[]');
                        return App.helpers.normalizePickerOptions(parsed, []);
                    } catch (error) {
                        return [];
                    }
                },

                loadEmbeddedPickerOptions() {
                    const parsed = App.storage.loadEmbeddedPayload();
                    return App.helpers.normalizePickerOptions(parsed.users, []);
                },

                normalizeDocuments(documents) {
                    if (!Array.isArray(documents)) {
                        return [];
                    }

                    const seen = new Set();

                    return [...documents]
                        .filter(documentItem => documentItem && Array.isArray(documentItem.rows))
                        .map(documentItem => ({
                            ...documentItem,
                            fingerprint: App.parser.getDocumentFingerprint(documentItem)
                        }))
                        .sort((a, b) => new Date(b.storedAt || 0) - new Date(a.storedAt || 0))
                        .filter(documentItem => {
                            const key = documentItem.fingerprint
                                || documentItem.id
                                || `${documentItem.sourceName || 'sin-nombre'}-${documentItem.storedAt || ''}-${documentItem.rows.length}`;

                            if (seen.has(key)) {
                                return false;
                            }

                            seen.add(key);
                            return true;
                        });
                },

                loadStoredDocuments() {
                    if (!App.runtime.storage) {
                        return [];
                    }

                    try {
                        const parsed = JSON.parse(App.runtime.storage.getItem(App.config.STORAGE_KEY) || '[]');
                        return App.storage.normalizeDocuments(parsed);
                    } catch (error) {
                        return [];
                    }
                },

                loadEmbeddedDocuments() {
                    const parsed = App.storage.loadEmbeddedPayload();
                    return App.storage.normalizeDocuments(parsed.documents);
                },

                getVisibleStoredDocuments(...documentGroups) {
                    const now = new Date();
                    const todayKey = App.storage.getDateKey(now);
                    const mergedDocuments = App.storage.normalizeDocuments(documentGroups.flat());

                    if (!App.storage.isHistoryEnabledNow(now)) {
                        return [];
                    }

                    return mergedDocuments.filter(documentItem => App.storage.getDateKey(documentItem.storedAt) === todayKey);
                },

                loadOutsideHoursHistoryPreference() {
                    if (!App.runtime.storage) {
                        return false;
                    }

                    try {
                        return App.runtime.storage.getItem(App.config.OUTSIDE_HOURS_HISTORY_KEY) === '1';
                    } catch (error) {
                        return false;
                    }
                },

                normalizeOrderLookupRows(rows) {
                    if (!Array.isArray(rows)) {
                        return [];
                    }

                    const allowedKeys = new Set([
                        ...App.orderLookup.getFieldDefinitions().map(field => field.key),
                        'logisticsType'
                    ]);

                    return rows
                        .filter(row => row && typeof row === 'object')
                        .map(row => Array.from(allowedKeys).reduce((normalizedRow, key) => {
                            normalizedRow[key] = App.helpers.normalizeComparable(row[key]);
                            return normalizedRow;
                        }, {}))
                        .filter(row => (
                            App.orderLookup.getSearchableFields().some(field => App.helpers.normalizeComparable(row[field.key]) !== '')
                            || App.helpers.normalizeComparable(row.estado) !== ''
                        ));
                },

                loadStoredOrderLookupState() {
                    if (!App.runtime.storage) {
                        return null;
                    }

                    try {
                        const parsed = JSON.parse(App.runtime.storage.getItem(App.config.ORDER_LOOKUP_STORAGE_KEY) || 'null');
                        if (!parsed || typeof parsed !== 'object') {
                            return null;
                        }

                        const rows = App.storage.normalizeOrderLookupRows(parsed.rows);
                        if (rows.length === 0) {
                            return null;
                        }

                        return {
                            rows,
                            fileName: String(parsed.fileName || '').trim(),
                            filters: {
                                query: App.helpers.normalizeComparable(parsed.filters?.query),
                                state: App.helpers.normalizeComparable(parsed.filters?.state),
                                route: App.helpers.normalizeComparable(parsed.filters?.route)
                            }
                        };
                    } catch (error) {
                        return null;
                    }
                },

                persistStoredOrderLookupState({
                    rows = App.state.orderLookupRows,
                    fileName = App.state.orderLookupFileName,
                    filters = App.orderLookup.getActiveFilters()
                } = {}) {
                    if (!App.runtime.storage) {
                        return false;
                    }

                    try {
                        const normalizedRows = App.storage.normalizeOrderLookupRows(rows);
                        if (normalizedRows.length === 0) {
                            App.storage.clearStoredOrderLookupState();
                            return true;
                        }

                        App.runtime.storage.setItem(
                            App.config.ORDER_LOOKUP_STORAGE_KEY,
                            JSON.stringify({
                                version: 1,
                                fileName: String(fileName || '').trim(),
                                rows: normalizedRows,
                                filters: {
                                    query: App.helpers.normalizeComparable(filters?.query),
                                    state: App.helpers.normalizeComparable(filters?.state),
                                    route: App.helpers.normalizeComparable(filters?.route)
                                }
                            })
                        );
                        return true;
                    } catch (error) {
                        console.error(error);
                        return false;
                    }
                },

                clearStoredOrderLookupState() {
                    if (!App.runtime.storage) {
                        return;
                    }

                    try {
                        App.runtime.storage.removeItem(App.config.ORDER_LOOKUP_STORAGE_KEY);
                    } catch (error) {
                        console.error(error);
                    }
                },

                writeEmbeddedState({ documents = App.state.embeddedDocuments, users = App.state.pickerOptions } = {}) {
                    const normalizedDocuments = App.storage.normalizeDocuments(documents);
                    const normalizedUsers = App.helpers.normalizePickerOptions(users, []);
                    const embeddedHistoryNode = document.getElementById(App.config.EMBEDDED_HISTORY_SCRIPT_ID);

                    App.state.embeddedDocuments = normalizedDocuments;

                    if (!embeddedHistoryNode) {
                        return;
                    }

                    embeddedHistoryNode.textContent = JSON.stringify(
                        { version: 2, documents: normalizedDocuments, users: normalizedUsers },
                        null,
                        0
                    ).replace(/</g, '\\u003c');
                },

                writeEmbeddedDocuments(documents) {
                    App.storage.writeEmbeddedState({ documents, users: App.state.pickerOptions });
                },

                writeEmbeddedPickerOptions(users) {
                    App.storage.writeEmbeddedState({ documents: App.state.embeddedDocuments, users });
                },

                persistStoredDocuments(nextDocuments) {
                    if (!App.runtime.storage) {
                        return false;
                    }

                    try {
                        const normalizedDocuments = App.storage.normalizeDocuments(nextDocuments);
                        App.runtime.storage.setItem(App.config.STORAGE_KEY, JSON.stringify(normalizedDocuments));
                        App.storage.writeEmbeddedDocuments(normalizedDocuments);
                        App.state.storedDocuments = App.storage.getVisibleStoredDocuments(normalizedDocuments);
                        App.storage.updateStorageUI();
                        App.orderLookup.syncReferenceScope();
                        return true;
                    } catch (error) {
                        console.error(error);
                        App.ui.showMessage(
                            'No fue posible guardar el historial local. Revisa el espacio disponible del navegador e intenta nuevamente.',
                            'error',
                            { title: 'Historial' }
                        );
                        return false;
                    }
                },

                getCurrentMinutes(date = new Date()) {
                    return (date.getHours() * 60) + date.getMinutes();
                },

                isWithinStorageWindow(date = new Date()) {
                    const currentMinutes = App.storage.getCurrentMinutes(date);
                    return currentMinutes >= App.config.STORAGE_START_MINUTES && currentMinutes < App.config.STORAGE_END_MINUTES;
                },

                isHistoryEnabledNow(date = new Date()) {
                    return App.storage.isWithinStorageWindow(date) || App.state.outsideHoursHistoryEnabled;
                },

                canStoreNow() {
                    return Boolean(App.runtime.storage) && App.storage.isHistoryEnabledNow();
                },

                getDateKey(value = new Date()) {
                    const date = new Date(value);
                    const year = date.getFullYear();
                    const month = String(date.getMonth() + 1).padStart(2, '0');
                    const day = String(date.getDate()).padStart(2, '0');
                    return `${year}-${month}-${day}`;
                },

                formatClockTime(value) {
                    return App.runtime.clockTimeFormatter.format(new Date(value));
                },

                formatDateTime(value) {
                    return App.runtime.dateTimeFormatter.format(new Date(value));
                },

                getNextStorageBoundary(date = new Date()) {
                    const currentMinutes = App.storage.getCurrentMinutes(date);
                    const boundary = new Date(date);

                    if (currentMinutes < App.config.STORAGE_START_MINUTES) {
                        boundary.setHours(
                            Math.floor(App.config.STORAGE_START_MINUTES / 60),
                            App.config.STORAGE_START_MINUTES % 60,
                            0,
                            0
                        );
                        return boundary;
                    }

                    if (currentMinutes < App.config.STORAGE_END_MINUTES) {
                        boundary.setHours(
                            Math.floor(App.config.STORAGE_END_MINUTES / 60),
                            App.config.STORAGE_END_MINUTES % 60,
                            0,
                            0
                        );
                        return boundary;
                    }

                    boundary.setDate(boundary.getDate() + 1);
                    boundary.setHours(
                        Math.floor(App.config.STORAGE_START_MINUTES / 60),
                        App.config.STORAGE_START_MINUTES % 60,
                        0,
                        0
                    );
                    return boundary;
                },

                formatBoundaryCountdown(targetDate, now = new Date()) {
                    const diffMs = Math.max(0, targetDate.getTime() - now.getTime());
                    const totalMinutes = Math.ceil(diffMs / 60000);
                    const hours = Math.floor(totalMinutes / 60);
                    const minutes = totalMinutes % 60;

                    if (hours > 0) {
                        return `${hours}h ${String(minutes).padStart(2, '0')}m`;
                    }

                    return `${Math.max(1, minutes)}m`;
                },

                buildStorageStatusMarkup(label, timeValue, meta) {
                    return `
                        <span class="status-pill-dot" aria-hidden="true"></span>
                        <span class="status-pill-content">
                            <span class="status-pill-label">${App.helpers.escapeHtml(label)}</span>
                            <span class="status-pill-time">${App.helpers.escapeHtml(App.storage.formatClockTime(timeValue))}</span>
                            <span class="status-pill-meta">${App.helpers.escapeHtml(meta)}</span>
                        </span>
                    `;
                },

                updateOutsideHoursToggleButton(now = new Date()) {
                    const button = App.dom.outsideHoursToggleBtn;
                    if (!button) {
                        return;
                    }

                    const enabled = Boolean(App.runtime.storage);
                    const insideWindow = App.storage.isWithinStorageWindow(now);
                    const isActive = enabled && App.state.outsideHoursHistoryEnabled;

                    button.disabled = !enabled;
                    button.classList.toggle('is-active', isActive);
                    button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
                    button.setAttribute(
                        'title',
                        !enabled
                            ? 'No disponible sin almacenamiento local'
                            : isActive
                                ? insideWindow
                                    ? 'Historial fuera de horario activo para cuando cierre la franja'
                                    : 'Desactivar historial fuera de horario'
                                : insideWindow
                                    ? 'Mantener historial activo tambien fuera de horario'
                                    : 'Activar historial fuera de horario'
                    );
                },

                toggleOutsideHoursHistory() {
                    if (!App.runtime.storage) {
                        return;
                    }

                    const nextValue = !App.state.outsideHoursHistoryEnabled;

                    try {
                        if (nextValue) {
                            App.runtime.storage.setItem(App.config.OUTSIDE_HOURS_HISTORY_KEY, '1');
                        } else {
                            App.runtime.storage.removeItem(App.config.OUTSIDE_HOURS_HISTORY_KEY);
                        }

                        App.state.outsideHoursHistoryEnabled = nextValue;

                        const localDocuments = App.storage.loadStoredDocuments();
                        const previousVisibleDocumentIds = new Set(App.state.storedDocuments.map(documentItem => documentItem.id));
                        App.state.storedDocuments = App.storage.getVisibleStoredDocuments(localDocuments, App.state.embeddedDocuments);
                        const shouldHideCurrentBase = !nextValue
                            && App.state.activeBaseRows.length > 0
                            && App.state.activeBaseRows.every(row => previousVisibleDocumentIds.has(row.documentId))
                            && App.state.storedDocuments.length === 0;

                        if (shouldHideCurrentBase) {
                            App.state.activeBaseRows = [];
                            App.state.activeBaseMessage = '';
                        }

                        App.storage.updateStorageUI();
                        App.orderLookup.syncReferenceScope();

                        if (App.dom.searchInput.value.trim()) {
                            App.table.handleSearch();
                        } else if (shouldHideCurrentBase) {
                            App.table.resetWorkspaceView(false);
                        }

                        App.ui.showMessage(
                            nextValue
                                ? 'El historial seguira activo tambien fuera del horario.'
                                : 'El historial fuera de horario fue desactivado.',
                            nextValue ? 'success' : 'info',
                            { title: 'Historial' }
                        );
                    } catch (error) {
                        console.error(error);
                        App.ui.showMessage(
                            'No fue posible actualizar la preferencia del historial fuera de horario.',
                            'error',
                            { title: 'Historial' }
                        );
                    }
                },

                refreshStorageStatus(now = new Date()) {
                    if (!App.dom.storageStatus) {
                        return;
                    }

                    const hasVisibleHistory = App.state.storedDocuments.length > 0;
                    App.storage.updateOutsideHoursToggleButton(now);

                    if (!App.runtime.storage) {
                        App.dom.storageStatus.className = 'status-pill is-disabled';
                        App.dom.storageStatus.innerHTML = App.storage.buildStorageStatusMarkup(
                            hasVisibleHistory ? 'Lectura' : 'Sin storage',
                            now,
                            hasVisibleHistory ? 'portable' : 'sin storage'
                        );
                        return;
                    }

                    const insideWindow = App.storage.isWithinStorageWindow(now);
                    const outsideHoursActive = !insideWindow && App.state.outsideHoursHistoryEnabled;
                    const nextBoundary = App.storage.getNextStorageBoundary(now);
                    const metaText = outsideHoursActive
                        ? '24h activo'
                        : insideWindow
                        ? `cierra ${App.storage.formatBoundaryCountdown(nextBoundary, now)}`
                        : App.storage.getCurrentMinutes(now) < App.config.STORAGE_START_MINUTES
                            ? `abre ${App.storage.formatBoundaryCountdown(nextBoundary, now)}`
                            : `reabre ${App.storage.formatBoundaryCountdown(nextBoundary, now)}`;

                    App.dom.storageStatus.className = `status-pill ${outsideHoursActive ? 'is-extended' : insideWindow ? 'is-open' : 'is-closed'}`;
                    App.dom.storageStatus.innerHTML = App.storage.buildStorageStatusMarkup(
                        outsideHoursActive ? '24h' : insideWindow ? 'Activo' : 'Cerrado',
                        now,
                        metaText
                    );
                },

                renderHistoryPreview(documents) {
                    const selectedIds = App.state.selectedHistoryDocumentIds;
                    return documents
                        .map(documentItem => {
                            const typeSummary = App.helpers.summarizeDocumentType(documentItem.rows);
                            return `
                                <button
                                    class="history-item${selectedIds.has(documentItem.id) ? ' is-selected' : ''}"
                                    type="button"
                                    data-history-document-id="${App.helpers.escapeHtml(documentItem.id || '')}"
                                    aria-pressed="${selectedIds.has(documentItem.id) ? 'true' : 'false'}"
                                    title="Haz clic para sumar o quitar este archivo de la seleccion"
                                >
                                    <strong>${App.helpers.escapeHtml(documentItem.sourceName || 'Documento sin nombre')}</strong>
                                    <div class="history-type-row">
                                        <span class="history-type-badge is-${App.helpers.escapeHtml(typeSummary.key)}">${App.helpers.escapeHtml(typeSummary.label)}</span>
                                        <span class="history-type-meta">${App.helpers.escapeHtml(typeSummary.detail)}</span>
                                    </div>
                                    <span>${documentItem.rows.length} numero(s) - Picker: ${App.helpers.escapeHtml(documentItem.picker || 'sin picker')}</span>
                                    <span>Guardado: ${App.helpers.escapeHtml(App.storage.formatDateTime(documentItem.storedAt))}</span>
                                </button>
                            `;
                        })
                        .join('');
                },

                syncSelectedHistoryDocumentIds() {
                    const validDocumentIds = new Set(App.state.storedDocuments.map(documentItem => documentItem.id));
                    App.state.selectedHistoryDocumentIds = new Set(
                        Array.from(App.state.selectedHistoryDocumentIds).filter(documentId => validDocumentIds.has(documentId))
                    );
                },

                toggleHistoryDocumentSelection(documentId) {
                    const normalizedId = String(documentId || '').trim();
                    if (!normalizedId) {
                        return;
                    }

                    const availableDocuments = App.state.storedDocuments.filter(documentItem => documentItem?.id);
                    if (!availableDocuments.some(documentItem => documentItem.id === normalizedId)) {
                        return;
                    }

                    const nextSelectedIds = new Set(App.state.selectedHistoryDocumentIds);
                    if (nextSelectedIds.has(normalizedId)) {
                        nextSelectedIds.delete(normalizedId);
                    } else {
                        nextSelectedIds.add(normalizedId);
                    }

                    const selectedDocuments = availableDocuments.filter(documentItem => nextSelectedIds.has(documentItem.id));
                    if (selectedDocuments.length === 0) {
                        App.table.showStoredDocuments();
                        return;
                    }

                    const totalRows = selectedDocuments.reduce((sum, documentItem) => sum + documentItem.rows.length, 0);
                    App.table.setActiveBaseFromDocuments(
                        selectedDocuments,
                        `Mostrando ${selectedDocuments.length} archivo(s) guardado(s) seleccionado(s) con ${totalRows} numero(s).`,
                        { selectedHistoryDocumentIds: nextSelectedIds }
                    );
                },

                updateStorageUI() {
                    const now = new Date();
                    const hasVisibleHistory = App.state.storedDocuments.length > 0;
                    App.storage.syncSelectedHistoryDocumentIds();
                    const hasSelectedHistoryDocuments = App.state.selectedHistoryDocumentIds.size > 0;

                    App.storage.refreshStorageStatus(now);

                    if (!App.runtime.storage) {
                        App.dom.storageSummary.textContent = hasVisibleHistory
                            ? `${App.state.storedDocuments.length} documento(s) recuperado(s) desde el HTML portable.`
                            : 'Historial local no disponible.';
                        App.dom.storageNote.textContent = hasVisibleHistory
                            ? 'Puedes consultar estos datos y guardar otra copia portable, pero los nuevos documentos no quedaran guardados localmente. Haz clic en uno o varios archivos guardados para ver sus pedidos.'
                            : 'Para usar busquedas sobre documentos guardados, abre el archivo en un navegador con almacenamiento local habilitado.';
                        App.dom.historyList.innerHTML = hasVisibleHistory
                            ? App.storage.renderHistoryPreview(App.state.storedDocuments.slice(0, App.config.HISTORY_PREVIEW_LIMIT))
                            : '<div class="history-empty">No se puede acceder al historial guardado desde este navegador.</div>';
                        App.dom.showStoredBtn.textContent = hasSelectedHistoryDocuments ? 'Mostrar todos' : 'Mostrar guardados';
                        App.dom.showStoredBtn.disabled = !hasVisibleHistory;
                        App.dom.clearStoredBtn.disabled = true;
                        App.dom.savePortableBtn.disabled = !App.storage.canSavePortableState();
                        return;
                    }

                    const insideWindow = App.storage.isWithinStorageWindow(now);
                    const outsideHoursActive = !insideWindow && App.state.outsideHoursHistoryEnabled;

                    if (!hasVisibleHistory) {
                        App.dom.storageSummary.textContent = outsideHoursActive
                            ? 'Sin documentos guardados. El modo 24h esta activo.'
                            : 'Sin documentos guardados.';
                        App.dom.historyList.innerHTML = '<div class="history-empty">Aun no hay documentos guardados en este navegador.</div>';
                    } else {
                        const latestDocument = App.state.storedDocuments[0];
                        const totalRows = App.state.storedDocuments.reduce((sum, documentItem) => sum + documentItem.rows.length, 0);
                        App.dom.storageSummary.textContent = `${App.state.storedDocuments.length} doc(s) · ${totalRows} num(s) · Ult.: ${App.storage.formatDateTime(latestDocument.storedAt)}.`;
                        App.dom.historyList.innerHTML = App.storage.renderHistoryPreview(
                            App.state.storedDocuments.slice(0, App.config.HISTORY_PREVIEW_LIMIT)
                        );
                    }

                    App.dom.storageNote.textContent = insideWindow
                        ? App.state.outsideHoursHistoryEnabled
                            ? `Entre ${App.config.STORAGE_WINDOW_LABEL} las cargas se guardan normalmente y el modo 24h quedara activo cuando termine la franja.${hasVisibleHistory ? ' Haz clic en uno o varios archivos guardados para ver sus pedidos.' : ''}`
                            : `Entre ${App.config.STORAGE_WINDOW_LABEL} las cargas quedan guardadas aunque cierres la ventana. Si quieres moverlas a otro computador, usa "Guardar HTML portable".${hasVisibleHistory ? ' Haz clic en uno o varios archivos guardados para ver sus pedidos.' : ''}`
                        : outsideHoursActive
                            ? `Modo 24h activo. Las cargas nuevas se seguiran agregando al historial hasta que lo desactives o cambie el dia.${hasVisibleHistory ? ' Haz clic en uno o varios archivos guardados para ver sus pedidos.' : ''}`
                            : 'Fuera de la franja operativa puedes procesar archivos, pero no se agregan al historial guardado. Usa 24h si quieres activarlo.';
                    App.dom.showStoredBtn.textContent = hasSelectedHistoryDocuments ? 'Mostrar todos' : 'Mostrar guardados';
                    App.dom.showStoredBtn.disabled = !hasVisibleHistory;
                    App.dom.clearStoredBtn.disabled = !hasVisibleHistory;
                    App.dom.savePortableBtn.disabled = !App.storage.canSavePortableState();
                },

                pruneHistoryBySchedule() {
                    const localDocuments = App.storage.loadStoredDocuments();
                    App.state.storedDocuments = App.storage.getVisibleStoredDocuments(localDocuments, App.state.embeddedDocuments);

                    if (!App.runtime.storage) {
                        return;
                    }

                    const now = new Date();
                    const todayKey = App.storage.getDateKey(now);
                    const hasDocumentsFromAnotherDay = localDocuments.some(
                        documentItem => App.storage.getDateKey(documentItem.storedAt) !== todayKey
                    );

                    if (!hasDocumentsFromAnotherDay) {
                        return;
                    }

                    const previousDocumentIds = new Set(localDocuments.map(documentItem => documentItem.id));
                    const currentDayDocuments = localDocuments.filter(
                        documentItem => App.storage.getDateKey(documentItem.storedAt) === todayKey
                    );

                    try {
                        App.runtime.storage.setItem(App.config.STORAGE_KEY, JSON.stringify(currentDayDocuments));
                    } catch (error) {
                        console.error(error);
                    }

                    App.state.storedDocuments = App.storage.getVisibleStoredDocuments(currentDayDocuments, App.state.embeddedDocuments);
                    App.storage.updateStorageUI();
                    App.orderLookup.syncReferenceScope();

                    if (
                        App.state.activeBaseRows.length > 0
                        && App.state.activeBaseRows.every(row => previousDocumentIds.has(row.documentId))
                    ) {
                        App.table.resetWorkspaceView(false);
                    }
                },

                updateWindowControls() {
                    const enabled = Boolean(App.runtime.storage);
                    App.dom.dropzone.classList.toggle('is-disabled', !enabled);
                    App.dom.processTextBtn.disabled = !enabled;
                    App.dom.processTextBtn.classList.toggle('is-disabled', !enabled);
                    App.dom.dropzone.title = enabled
                        ? 'Subir archivo .txt'
                        : 'No se pudo habilitar el almacenamiento local del navegador.';
                },

                canSavePortableState() {
                    return App.state.storedDocuments.length > 0 || App.users.hasCustomizedPickerOptions();
                },

                getPortableEmbeddedState() {
                    return JSON.stringify(
                        { version: 2, documents: App.state.storedDocuments, users: App.state.pickerOptions },
                        null,
                        0
                    ).replace(/</g, '\\u003c');
                },

                readAppStylesheetText() {
                    const stylesheetLink = document.querySelector('link[data-app-stylesheet]');
                    if (!stylesheetLink) {
                        return '';
                    }

                    const appSheet = Array.from(document.styleSheets).find(sheet => sheet.href === stylesheetLink.href);
                    if (!appSheet) {
                        return '';
                    }

                    try {
                        return Array.from(appSheet.cssRules).map(rule => rule.cssText).join('\n');
                    } catch (error) {
                        console.error(error);
                        return '';
                    }
                },

                buildPortableHtml() {
                    const clone = document.documentElement.cloneNode(true);
                    const cloneEmbeddedState = clone.querySelector(`#${App.config.EMBEDDED_HISTORY_SCRIPT_ID}`);
                    const cloneMessageStack = clone.querySelector('#messageStack');
                    const cloneAppScript = clone.querySelector('script[data-app-script]');
                    const cloneStylesheetLink = clone.querySelector('link[data-app-stylesheet]');
                    const portableStyles = App.storage.readAppStylesheetText();

                    if (cloneEmbeddedState) {
                        cloneEmbeddedState.textContent = App.storage.getPortableEmbeddedState();
                    }

                    if (cloneMessageStack) {
                        cloneMessageStack.innerHTML = '';
                    }

                    if (cloneAppScript) {
                        cloneAppScript.remove();
                    }

                    if (cloneStylesheetLink && portableStyles) {
                        const inlineStyles = document.createElement('style');
                        inlineStyles.textContent = portableStyles;
                        cloneStylesheetLink.replaceWith(inlineStyles);
                    }

                    const portableScript = document.createElement('script');
                    portableScript.textContent = `(${bootstrap.toString()})();`.replace(/<\/script/gi, '<\\/script');
                    clone.querySelector('body')?.appendChild(portableScript);

                    return `<!DOCTYPE html>\n${clone.outerHTML}`;
                }
            },
            users: {
                loadPickerOptions() {
                    const storedPickerOptions = App.storage.loadStoredPickerOptions();
                    if (storedPickerOptions.length > 0) {
                        return storedPickerOptions;
                    }

                    const embeddedPickerOptions = App.storage.loadEmbeddedPickerOptions();
                    if (embeddedPickerOptions.length > 0) {
                        return embeddedPickerOptions;
                    }

                    return [...App.runtime.defaultPickerOptions];
                },

                persistPickerOptions(nextPickerOptions) {
                    const normalizedPickerOptions = App.helpers.normalizePickerOptions(nextPickerOptions, []);
                    if (normalizedPickerOptions.length === 0) {
                        App.ui.showMessage('Debe quedar al menos un usuario disponible.', 'warning', { title: 'Usuarios' });
                        return false;
                    }

                    if (App.runtime.storage) {
                        try {
                            App.runtime.storage.setItem(
                                App.config.PICKERS_STORAGE_KEY,
                                JSON.stringify(normalizedPickerOptions)
                            );
                        } catch (error) {
                            console.error(error);
                            App.ui.showMessage(
                                'No fue posible guardar la lista de usuarios. Intenta nuevamente.',
                                'error',
                                { title: 'Usuarios' }
                            );
                            return false;
                        }
                    }

                    App.state.pickerOptions = normalizedPickerOptions;
                    App.users.renderPickerOptions();
                    App.users.renderPickerManager();
                    App.storage.writeEmbeddedPickerOptions(normalizedPickerOptions);
                    App.storage.updateStorageUI();
                    return true;
                },

                renderPickerOptions() {
                    const previousValue = App.dom.pickerInput.value;
                    App.dom.pickerInput.innerHTML = '';

                    const placeholderOption = document.createElement('option');
                    placeholderOption.value = '';
                    placeholderOption.disabled = true;
                    placeholderOption.textContent = 'Seleccione una opcion...';
                    App.dom.pickerInput.appendChild(placeholderOption);

                    App.state.pickerOptions.forEach(option => {
                        const optionNode = document.createElement('option');
                        optionNode.value = option;
                        optionNode.textContent = option;
                        App.dom.pickerInput.appendChild(optionNode);
                    });

                    App.dom.pickerInput.value = App.state.pickerOptions.includes(previousValue) ? previousValue : '';
                    if (!App.dom.pickerInput.value) {
                        App.dom.pickerInput.selectedIndex = 0;
                    }

                    App.dom.pickerMeta.textContent = `${App.state.pickerOptions.length} usuario(s) disponible(s)`;
                },

                renderPickerManager() {
                    App.dom.pickerManagerCount.textContent = String(App.state.pickerOptions.length);
                    App.dom.pickerListTotal.textContent = String(App.state.pickerOptions.length);
                    App.dom.pickerList.innerHTML = App.state.pickerOptions.length === 0
                        ? '<div class="picker-empty">No hay usuarios disponibles.</div>'
                        : App.state.pickerOptions
                            .map(option => `
                                <div class="picker-row">
                                    <div class="picker-row-main">
                                        <span class="picker-row-badge">${App.helpers.escapeHtml(option.charAt(0).toUpperCase())}</span>
                                        <div class="picker-row-info">
                                            <strong>${App.helpers.escapeHtml(option)}</strong>
                                        </div>
                                    </div>
                                    <button
                                        class="btn btn-secondary picker-remove-btn"
                                        type="button"
                                        data-remove-picker="${App.helpers.escapeHtml(option)}"
                                        aria-label="Eliminar ${App.helpers.escapeHtml(option)}"
                                        title="Eliminar ${App.helpers.escapeHtml(option)}"
                                    >X</button>
                                </div>
                            `)
                            .join('');

                    App.dom.manageUsersBtn.textContent = App.dom.pickerManager.hidden ? '+' : '\u2212';
                    App.dom.manageUsersBtn.title = App.dom.pickerManager.hidden
                        ? 'Mostrar lista de usuarios'
                        : 'Cerrar lista de usuarios';
                    App.dom.manageUsersBtn.setAttribute('aria-label', App.dom.manageUsersBtn.title);
                    App.dom.manageUsersBtn.classList.toggle('is-active', !App.dom.pickerManager.hidden);
                },

                togglePickerManager(forceOpen) {
                    const shouldOpen = typeof forceOpen === 'boolean' ? forceOpen : App.dom.pickerManager.hidden;
                    App.dom.pickerManager.hidden = !shouldOpen;
                    App.users.renderPickerManager();

                    if (shouldOpen) {
                        App.dom.newPickerInput.focus();
                    }
                },

                handleAddPicker() {
                    const nextPickerName = App.helpers.normalizePickerName(App.dom.newPickerInput.value);
                    if (!nextPickerName) {
                        App.dom.newPickerInput.focus();
                        return;
                    }

                    const alreadyExists = App.state.pickerOptions.some(
                        option => option.toLowerCase() === nextPickerName.toLowerCase()
                    );
                    if (alreadyExists) {
                        App.ui.showMessage(`El usuario "${nextPickerName}" ya existe.`, 'warning', { title: 'Usuarios' });
                        App.dom.newPickerInput.focus();
                        App.dom.newPickerInput.select();
                        return;
                    }

                    if (!App.users.persistPickerOptions([...App.state.pickerOptions, nextPickerName])) {
                        return;
                    }

                    App.dom.pickerInput.value = nextPickerName;
                    App.dom.newPickerInput.value = '';
                    App.users.togglePickerManager(true);
                },

                removePickerOption(optionToRemove) {
                    const normalizedTarget = App.helpers.normalizePickerName(optionToRemove);
                    if (!normalizedTarget) {
                        return;
                    }

                    if (App.state.pickerOptions.length === 1) {
                        App.ui.showMessage('No se puede eliminar el ultimo usuario disponible.', 'warning', { title: 'Usuarios' });
                        return;
                    }

                    const confirmed = window.confirm(
                        `Se eliminara el usuario "${normalizedTarget}" de la lista. Deseas continuar?`
                    );
                    if (!confirmed) {
                        return;
                    }

                    const remainingPickerOptions = App.state.pickerOptions.filter(
                        option => option.toLowerCase() !== normalizedTarget.toLowerCase()
                    );
                    if (!App.users.persistPickerOptions(remainingPickerOptions)) {
                        return;
                    }

                    if (App.dom.pickerInput.value.toLowerCase() === normalizedTarget.toLowerCase()) {
                        App.dom.pickerInput.value = '';
                    }
                },

                hasCustomizedPickerOptions() {
                    const normalizedDefaults = App.helpers.normalizePickerOptions(App.runtime.defaultPickerOptions, []);
                    if (App.state.pickerOptions.length !== normalizedDefaults.length) {
                        return true;
                    }

                    return App.state.pickerOptions.some((option, index) => option !== normalizedDefaults[index]);
                }
            },

            orderLookup: {
                getFieldDefinitions() {
                    return App.config.ORDER_LOOKUP_FIELDS;
                },

                getSearchableFields() {
                    return App.orderLookup.getFieldDefinitions().filter(field => field.searchable);
                },

                getVisibleColumns() {
                    return [
                        { key: 'dato2', label: 'DATO2' },
                        { key: 'estado', label: 'ESTADO' },
                        { key: 'logisticsType', label: 'RUTAS' }
                    ];
                },

                getDefaultDisplayColumns() {
                    return App.orderLookup.getVisibleColumns();
                },

                getReferenceOrderDescriptors() {
                    const descriptors = [];
                    const seen = new Set();

                    App.table.getSearchableRows().forEach(row => {
                        const orderNumber = App.helpers.normalizeComparable(row.numero);
                        if (!orderNumber || seen.has(orderNumber)) {
                            return;
                        }

                        seen.add(orderNumber);
                        descriptors.push({
                            orderNumber,
                            logisticsType: row.type === 'flex'
                                ? 'Flex'
                                : row.type === 'colecta'
                                    ? 'Colecta'
                                    : ''
                        });
                    });

                    return descriptors;
                },

                getReferenceOrderNumbers() {
                    return new Set(
                        App.orderLookup.getReferenceOrderDescriptors()
                            .map(descriptor => descriptor.orderNumber)
                    );
                },

                getReferenceOrderTypeMap() {
                    return new Map(
                        App.orderLookup.getReferenceOrderDescriptors()
                            .filter(descriptor => descriptor.logisticsType)
                            .map(descriptor => [descriptor.orderNumber, descriptor.logisticsType])
                    );
                },

                rowMatchesReferenceOrders(row, referenceOrderNumbers = App.orderLookup.getReferenceOrderNumbers()) {
                    if (!(referenceOrderNumbers instanceof Set) || referenceOrderNumbers.size === 0) {
                        return false;
                    }

                    return App.orderLookup.getSearchableFields().some(field =>
                        referenceOrderNumbers.has(App.helpers.normalizeComparable(row[field.key]))
                    );
                },

                getMatchedReferenceOrderNumber(
                    row,
                    referenceOrderNumbers = App.orderLookup.getReferenceOrderNumbers()
                ) {
                    if (!row || !(referenceOrderNumbers instanceof Set) || referenceOrderNumbers.size === 0) {
                        return '';
                    }

                    return App.orderLookup.getSearchableFields()
                        .map(field => App.helpers.normalizeComparable(row[field.key]))
                        .find(value => value && referenceOrderNumbers.has(value)) || '';
                },

                getMatchedExcelRows(
                    rows = App.state.orderLookupRows,
                    referenceOrderNumbers = App.orderLookup.getReferenceOrderNumbers(),
                    referenceOrderTypeMap = App.orderLookup.getReferenceOrderTypeMap()
                ) {
                    if (!Array.isArray(rows) || rows.length === 0 || referenceOrderNumbers.size === 0) {
                        return [];
                    }

                    return rows
                        .filter(row => App.orderLookup.rowMatchesReferenceOrders(row, referenceOrderNumbers))
                        .map(row => {
                            const matchedReferenceNumber = App.orderLookup.getMatchedReferenceOrderNumber(
                                row,
                                referenceOrderNumbers
                            );

                            return {
                                ...row,
                                matchedReferenceNumber,
                                logisticsType: App.orderLookup.resolveLogisticsType(row, referenceOrderTypeMap)
                            };
                        });
                },

                getMissingReferenceRows(
                    referenceOrderDescriptors = App.orderLookup.getReferenceOrderDescriptors(),
                    matchedRows = []
                ) {
                    if (!Array.isArray(referenceOrderDescriptors) || referenceOrderDescriptors.length === 0) {
                        return [];
                    }

                    const matchedReferenceNumbers = new Set(
                        matchedRows
                            .map(row => App.helpers.normalizeComparable(row.matchedReferenceNumber))
                            .filter(Boolean)
                    );

                    return referenceOrderDescriptors
                        .filter(descriptor => !matchedReferenceNumbers.has(descriptor.orderNumber))
                        .map(descriptor => ({
                            dato2: descriptor.orderNumber,
                            estado: 'N/D',
                            logisticsType: descriptor.logisticsType,
                            referenceNumber: descriptor.orderNumber,
                            buyerOrder: '',
                            orderSource: '',
                            dispatchRoute: '',
                            matchedReferenceNumber: descriptor.orderNumber,
                            isMissingInWorkbook: true
                        }));
                },

                getScopedRows(
                    rows = App.state.orderLookupRows,
                    referenceOrderNumbers = App.orderLookup.getReferenceOrderNumbers(),
                    referenceOrderTypeMap = App.orderLookup.getReferenceOrderTypeMap(),
                    referenceOrderDescriptors = App.orderLookup.getReferenceOrderDescriptors()
                ) {
                    if (!Array.isArray(rows) || rows.length === 0 || referenceOrderNumbers.size === 0) {
                        return [];
                    }

                    const matchedRows = App.orderLookup.getMatchedExcelRows(
                        rows,
                        referenceOrderNumbers,
                        referenceOrderTypeMap
                    );
                    const missingRows = App.orderLookup.getMissingReferenceRows(
                        referenceOrderDescriptors,
                        matchedRows
                    );

                    return [...matchedRows, ...missingRows];
                },

                getSummaryMarkup(title, body, tone = 'default') {
                    const bodyMarkup = String(body || '').trim()
                        ? `<span>${App.helpers.escapeHtml(body)}</span>`
                        : '';

                    return {
                        toneClass: tone,
                        html: `
                            <strong>${App.helpers.escapeHtml(title)}</strong>
                            ${bodyMarkup}
                        `
                    };
                },

                updateSummary(title, body, tone = 'default') {
                    const summary = App.dom.orderLookupSummary;
                    if (!summary) {
                        return;
                    }

                    const { html, toneClass } = App.orderLookup.getSummaryMarkup(title, body, tone);
                    summary.className = `order-lookup-summary${toneClass === 'default' ? '' : ` is-${toneClass}`}`;
                    summary.classList.remove('is-hidden');
                    summary.innerHTML = html;
                },

                hideSummary() {
                    App.dom.orderLookupSummary?.classList.add('is-hidden');
                },

                setTableColumns(columns = App.state.orderLookupDisplayColumns) {
                    const normalizedColumns = Array.isArray(columns) && columns.length > 0
                        ? columns
                        : App.orderLookup.getDefaultDisplayColumns();

                    if (App.dom.orderLookupTableHead) {
                        App.dom.orderLookupTableHead.innerHTML = `
                            <tr>
                                ${normalizedColumns.map(column => `
                                    <th class="order-lookup-col order-lookup-col--${App.helpers.escapeHtml(column.key)}">${App.helpers.escapeHtml(column.label)}</th>
                                `).join('')}
                            </tr>
                        `;
                    }

                    return normalizedColumns;
                },

                renderNoMatchesRow(message = 'No hay coincidencias para mostrar.') {
                    const columns = App.orderLookup.setTableColumns(
                        App.state.orderLookupDisplayColumns.length > 0
                            ? App.state.orderLookupDisplayColumns
                            : App.orderLookup.getDefaultDisplayColumns()
                    );

                    if (App.dom.orderLookupTableBody) {
                        App.dom.orderLookupTableBody.innerHTML = `
                            <tr>
                                <td colspan="${columns.length}" class="order-lookup-empty-row">${App.helpers.escapeHtml(message)}</td>
                            </tr>
                        `;
                    }

                    if (App.dom.orderLookupTableWrap) {
                        App.dom.orderLookupTableWrap.classList.add('is-hidden');
                    }
                },

                resetResults() {
                    App.state.orderLookupMatches = [];
                    App.state.orderLookupDisplayColumns = App.orderLookup.getVisibleColumns();
                    App.orderLookup.setTableColumns(App.state.orderLookupDisplayColumns);
                    App.orderLookup.renderNoMatchesRow('No hay coincidencias para mostrar.');
                    App.orderLookup.updateSummary(
                        'Sin resultados',
                        'Carga un archivo Excel para ver solo pedidos de Mercado Libre que coincidan con los numeros guardados o cargados y filtrarlos por estado, ruta o numero.',
                        'empty'
                    );
                },

                resetData() {
                    App.state.orderLookupRows = [];
                    App.state.orderLookupDisplayColumns = App.orderLookup.getVisibleColumns();
                    App.state.orderLookupAvailableStates = [];
                    App.state.orderLookupAvailableRoutes = [];
                    App.state.orderLookupMatches = [];
                    App.orderLookup.setLoadedFileName('');
                    if (App.dom.orderLookupInput) {
                        App.dom.orderLookupInput.value = '';
                    }
                    if (App.dom.orderLookupRouteFilter) {
                        App.dom.orderLookupRouteFilter.value = '';
                    }
                    App.orderLookup.populateStateFilter([]);
                    App.orderLookup.populateRouteFilter([]);
                    App.storage.clearStoredOrderLookupState();
                },

                restoreStoredState() {
                    const storedState = App.storage.loadStoredOrderLookupState();
                    if (!storedState) {
                        return;
                    }

                    App.state.orderLookupRows = storedState.rows;
                    App.state.orderLookupDisplayColumns = App.orderLookup.getVisibleColumns();
                    App.orderLookup.setLoadedFileName(storedState.fileName || 'Consulta recuperada');

                    if (App.dom.orderLookupInput) {
                        App.dom.orderLookupInput.value = storedState.filters.query || '';
                    }

                    App.orderLookup.populateStateFilter(App.orderLookup.getScopedRows());
                    App.orderLookup.populateRouteFilter(App.orderLookup.getScopedRows());

                    if (App.dom.orderLookupStateFilter) {
                        const normalizedState = App.helpers.normalizeComparable(storedState.filters.state);
                        const hasStoredOption = Array.from(App.dom.orderLookupStateFilter.options)
                            .some(option => App.helpers.normalizeComparable(option.value) === normalizedState);
                        App.dom.orderLookupStateFilter.value = hasStoredOption ? normalizedState : '';
                    }

                    if (App.dom.orderLookupRouteFilter) {
                        const normalizedRoute = App.helpers.normalizeComparable(storedState.filters.route);
                        const hasStoredOption = Array.from(App.dom.orderLookupRouteFilter.options)
                            .some(option => App.helpers.normalizeComparable(option.value) === normalizedRoute);
                        App.dom.orderLookupRouteFilter.value = hasStoredOption ? normalizedRoute : '';
                    }

                    App.orderLookup.applyFilters({ showWarningOnEmptySource: false });
                },

                setLoadedFileName(fileName) {
                    App.state.orderLookupFileName = fileName || '';

                    if (App.dom.orderLookupFileName) {
                        App.dom.orderLookupFileName.textContent = fileName || 'No hay archivo Excel cargado.';
                    }
                },

                populateStateFilter(rows = App.orderLookup.getScopedRows()) {
                    const filterSelect = App.dom.orderLookupStateFilter;
                    if (!filterSelect) {
                        return;
                    }

                    const previousValue = App.helpers.normalizeComparable(filterSelect.value);
                    const states = Array.from(new Set(
                        rows
                            .map(row => App.helpers.normalizeComparable(row.estado))
                            .filter(Boolean)
                    )).sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));

                    App.state.orderLookupAvailableStates = states;
                    filterSelect.innerHTML = `
                        <option value="">Todos los estados</option>
                        ${states.map(state => `<option value="${App.helpers.escapeHtml(state)}">${App.helpers.escapeHtml(state)}</option>`).join('')}
                    `;

                    filterSelect.value = states.includes(previousValue) ? previousValue : '';
                },

                populateRouteFilter(rows = App.orderLookup.getScopedRows()) {
                    const filterSelect = App.dom.orderLookupRouteFilter;
                    if (!filterSelect) {
                        return;
                    }

                    const previousValue = App.helpers.normalizeComparable(filterSelect.value);
                    const detectedRoutes = new Set(
                        rows
                            .map(row => App.helpers.normalizeComparable(row.logisticsType) || App.orderLookup.resolveLogisticsType(row))
                            .filter(Boolean)
                    );
                    const routes = ['Flex', 'Colecta']
                        .filter(route => detectedRoutes.has(route))
                        .concat(Array.from(detectedRoutes).filter(route => !['Flex', 'Colecta'].includes(route)));

                    App.state.orderLookupAvailableRoutes = routes;
                    filterSelect.innerHTML = `
                        <option value="">Todas las rutas</option>
                        ${routes.map(route => `<option value="${App.helpers.escapeHtml(route)}">${App.helpers.escapeHtml(route)}</option>`).join('')}
                    `;

                    filterSelect.value = routes.includes(previousValue) ? previousValue : '';
                },

                isMercadoLibreRow(row) {
                    const orderSource = App.helpers.normalizeHeaderKey(row.orderSource);
                    const dispatchRoute = App.helpers.normalizeHeaderKey(row.dispatchRoute);

                    return orderSource.includes('MERCADOLIBRE') || dispatchRoute.includes('MERCADOLIBRE');
                },

                resolveLogisticsType(row = {}, referenceOrderTypeMap = App.orderLookup.getReferenceOrderTypeMap()) {
                    if (referenceOrderTypeMap instanceof Map && referenceOrderTypeMap.size > 0) {
                        const matchedReferenceOrder = App.orderLookup.getSearchableFields()
                            .map(field => App.helpers.normalizeComparable(row[field.key]))
                            .find(value => value && referenceOrderTypeMap.has(value));

                        if (matchedReferenceOrder) {
                            return referenceOrderTypeMap.get(matchedReferenceOrder) || '';
                        }
                    }

                    const normalizedSignals = [
                        row.logisticsType,
                        row.dispatchRoute,
                        row.orderSource,
                        row.estado
                    ]
                        .map(App.helpers.normalizeHeaderKey)
                        .filter(Boolean)
                        .join(' ');

                    if (
                        normalizedSignals.includes('FLEX')
                        || normalizedSignals.includes('MESA05')
                        || normalizedSignals.includes('REVISION05')
                    ) {
                        return 'Flex';
                    }

                    if (
                        normalizedSignals.includes('COLECTA')
                        || normalizedSignals.includes('MESA03')
                        || normalizedSignals.includes('REVISION03')
                    ) {
                        return 'Colecta';
                    }

                    return '';
                },

                updateReferenceScopeEmptyState(reason = 'missingReference') {
                    const missingReference = reason === 'missingReference';
                    const title = missingReference ? 'Sin pedidos base' : 'Sin coincidencias';
                    const body = missingReference
                        ? 'Primero procesa etiquetas o usa el historial guardado para tener numeros con los que comparar el Excel.'
                        : 'El Excel cargado no contiene pedidos de Mercado Libre que coincidan con los numeros guardados o cargados actualmente.';
                    const rowMessage = missingReference
                        ? 'Primero carga etiquetas guardadas o pega texto para habilitar la consulta.'
                        : 'No hay pedidos del Excel que coincidan con los numeros guardados o cargados.';

                    App.state.orderLookupMatches = [];
                    App.orderLookup.updateSummary(title, body, missingReference ? 'empty' : 'warning');
                    App.orderLookup.renderNoMatchesRow(rowMessage);
                },

                findHeaderRowIndex(matrix) {
                    const requiredFields = App.orderLookup.getFieldDefinitions().filter(field => field.required);

                    return matrix.reduce((bestIndex, row, rowIndex) => {
                        if (!Array.isArray(row) || row.length === 0) {
                            return bestIndex;
                        }

                        const normalizedRow = row.map(App.helpers.normalizeHeaderKey);
                        const hasRequiredFields = requiredFields.every(field =>
                            normalizedRow.some(header => field.aliases.includes(header))
                        );
                        if (!hasRequiredFields) {
                            return bestIndex;
                        }

                        const score = App.orderLookup.getFieldDefinitions().reduce((total, field) => (
                            total + (normalizedRow.some(header => field.aliases.includes(header)) ? 1 : 0)
                        ), 0);

                        if (bestIndex === -1) {
                            return rowIndex;
                        }

                        const bestNormalizedRow = (matrix[bestIndex] || []).map(App.helpers.normalizeHeaderKey);
                        const bestScore = App.orderLookup.getFieldDefinitions().reduce((total, field) => (
                            total + (bestNormalizedRow.some(header => field.aliases.includes(header)) ? 1 : 0)
                        ), 0);

                        if (score > bestScore || (score === bestScore && rowIndex > bestIndex)) {
                            return rowIndex;
                        }

                        return bestIndex;
                    }, -1);
                },

                resolveColumns(headers) {
                    return App.orderLookup.getFieldDefinitions().reduce((fieldMap, field) => {
                        const columnIndex = headers.findIndex(header =>
                            field.aliases.includes(App.helpers.normalizeHeaderKey(header))
                        );

                        fieldMap[field.key] = {
                            header: columnIndex >= 0 ? headers[columnIndex] : '',
                            index: columnIndex
                        };
                        return fieldMap;
                    }, {});
                },

                extractSheetRows(sheet) {
                    const matrix = globalThis.XLSX.utils.sheet_to_json(sheet, {
                        header: 1,
                        defval: '',
                        raw: false,
                        blankrows: false
                    });
                    const headerRowIndex = App.orderLookup.findHeaderRowIndex(matrix);

                    if (headerRowIndex === -1) {
                        return {
                            rows: [],
                            displayColumns: []
                        };
                    }

                    const headers = matrix[headerRowIndex].map((header, index) =>
                        App.helpers.normalizeComparable(header) || `Columna ${index + 1}`
                    );
                    const resolvedColumns = App.orderLookup.resolveColumns(headers);
                    const requiredFields = App.orderLookup.getFieldDefinitions().filter(field => field.required);

                    if (requiredFields.some(field => resolvedColumns[field.key].index < 0)) {
                        return {
                            rows: [],
                            displayColumns: []
                        };
                    }

                    // Convertimos cada fila a un objeto canonico para que luego sea facil
                    // ampliar la consulta con mas columnas sin depender del encabezado original.
                    const rows = matrix
                        .slice(headerRowIndex + 1)
                        .filter(row => Array.isArray(row) && row.some(cell => App.helpers.normalizeComparable(cell) !== ''))
                        .map(row => {
                            const item = {};

                            App.orderLookup.getFieldDefinitions().forEach(field => {
                                item[field.key] = resolvedColumns[field.key].index >= 0
                                    ? App.helpers.normalizeComparable(row[resolvedColumns[field.key].index] ?? '')
                                    : '';
                            });

                            item.logisticsType = App.orderLookup.resolveLogisticsType(item);

                            return item;
                        })
                        .filter(row => (
                            App.orderLookup.getSearchableFields().some(field => App.helpers.normalizeComparable(row[field.key]) !== '')
                            || App.helpers.normalizeComparable(row.estado) !== ''
                        ));

                    return {
                        rows,
                        displayColumns: App.orderLookup.getVisibleColumns()
                    };
                },

                extractWorkbookRows(workbook) {
                    const workbookSheets = workbook.Workbook?.Sheets || [];
                    const visibleSheetNames = workbook.SheetNames.filter((sheetName, index) => !workbookSheets[index]?.Hidden);
                    const sheetNamesToRead = visibleSheetNames.length > 0 ? visibleSheetNames : workbook.SheetNames;

                    const parsedWorkbook = sheetNamesToRead.reduce((accumulator, sheetName) => {
                        const sheet = workbook.Sheets[sheetName];
                        const { rows, displayColumns } = App.orderLookup.extractSheetRows(sheet);

                        if (rows.length > 0) {
                            accumulator.rows.push(...rows);
                            if (displayColumns.length > accumulator.displayColumns.length) {
                                accumulator.displayColumns = displayColumns;
                            }
                        }

                        return accumulator;
                    }, {
                        rows: [],
                        displayColumns: []
                    });

                    const mercadoLibreRows = parsedWorkbook.rows.filter(App.orderLookup.isMercadoLibreRow);

                    if (parsedWorkbook.rows.length === 0) {
                        throw new Error('El archivo Excel no contiene columnas reconocibles para numero de pedido y estado. Se buscan encabezados como DATO2 o Nro.Ref., y ESTADO o Estado Rev.');
                    }

                    if (mercadoLibreRows.length === 0) {
                        throw new Error('El archivo Excel no contiene pedidos de Mercado Libre para mostrar.');
                    }

                    return parsedWorkbook;
                },

                renderRows(rows) {
                    if (!App.dom.orderLookupTableBody || !App.dom.orderLookupTableWrap) {
                        return;
                    }

                    const columns = App.orderLookup.setTableColumns(App.orderLookup.getVisibleColumns());
                    App.dom.orderLookupTableBody.innerHTML = rows
                        .map(row => `
                            <tr>
                                ${columns.map(column => `
                                    <td class="order-lookup-col order-lookup-col--${App.helpers.escapeHtml(column.key)}">
                                        ${App.orderLookup.getCellMarkup(row, column)}
                                    </td>
                                `).join('')}
                            </tr>
                        `)
                        .join('');

                    App.dom.orderLookupTableWrap.classList.remove('is-hidden');
                },

                getLogisticsBadgeMarkup(value) {
                    const normalizedValue = App.helpers.normalizeComparable(value);
                    if (!normalizedValue) {
                        return '<span class="order-lookup-badge is-unknown">Sin definir</span>';
                    }

                    const toneClass = normalizedValue === 'Flex'
                        ? 'is-flex'
                        : normalizedValue === 'Colecta'
                            ? 'is-colecta'
                            : 'is-unknown';

                    return `<span class="order-lookup-badge ${toneClass}">${App.helpers.escapeHtml(normalizedValue)}</span>`;
                },

                getCellMarkup(row, column) {
                    const rawValue = row[column.key] ?? '';

                    if (column.key === 'estado') {
                        const normalizedValue = App.helpers.normalizeComparable(rawValue) || 'Sin estado';
                        const toneClass = normalizedValue === 'N/D'
                            ? 'is-nd'
                            : normalizedValue === 'Sin estado'
                                ? 'is-empty'
                                : '';

                        return `<span class="order-lookup-state${toneClass ? ` ${toneClass}` : ''}">${App.helpers.escapeHtml(normalizedValue)}</span>`;
                    }

                    if (column.key === 'logisticsType') {
                        return App.orderLookup.getLogisticsBadgeMarkup(rawValue);
                    }

                    return App.helpers.escapeHtml(rawValue);
                },

                getActiveFilters() {
                    return {
                        query: App.helpers.normalizeComparable(App.dom.orderLookupInput?.value),
                        state: App.helpers.normalizeComparable(App.dom.orderLookupStateFilter?.value),
                        route: App.helpers.normalizeComparable(App.dom.orderLookupRouteFilter?.value)
                    };
                },

                filterRows(filters, rows = App.orderLookup.getScopedRows()) {
                    const searchableFields = App.orderLookup.getSearchableFields();

                    return rows.filter(row => {
                        const matchesQuery = !filters.query || searchableFields.some(field =>
                            App.helpers.normalizeComparable(row[field.key]) === filters.query
                        );
                        const matchesState = !filters.state || App.helpers.normalizeComparable(row.estado) === filters.state;
                        const matchesRoute = !filters.route || (
                            App.helpers.normalizeComparable(row.logisticsType)
                            || App.orderLookup.resolveLogisticsType(row)
                        ) === filters.route;

                        return matchesQuery && matchesState && matchesRoute;
                    });
                },

                buildNotFoundMessage(filters) {
                    const criteria = [];

                    if (filters.query) {
                        criteria.push(`el pedido "${filters.query}"`);
                    }

                    if (filters.state) {
                        criteria.push(`el estado "${filters.state}"`);
                    }

                    if (filters.route) {
                        criteria.push(`la ruta "${filters.route}"`);
                    }

                    if (criteria.length > 0) {
                        const criteriaLabel = criteria.length === 1
                            ? criteria[0]
                            : `${criteria.slice(0, -1).join(', ')} y ${criteria[criteria.length - 1]}`;

                        return `No se encontraron pedidos de Mercado Libre para ${criteriaLabel}.`;
                    }

                    return 'No hay pedidos de Mercado Libre que coincidan con los numeros guardados o cargados usando los filtros actuales.';
                },

                updateNotFoundState(filters) {
                    App.orderLookup.updateSummary(
                        'Sin coincidencias',
                        App.orderLookup.buildNotFoundMessage(filters),
                        'warning'
                    );
                    App.orderLookup.renderNoMatchesRow('No hay pedidos de Mercado Libre que coincidan con los numeros guardados o cargados usando los filtros actuales.');
                },

                updateFoundState(matches) {
                    const totals = matches.reduce((accumulator, row) => {
                        const normalizedState = App.helpers.normalizeComparable(row.estado);
                        if (normalizedState === 'N/D') {
                            accumulator.nd += 1;
                        } else {
                            accumulator.withWorkbook += 1;
                        }

                        accumulator.total += 1;
                        return accumulator;
                    }, { total: 0, withWorkbook: 0, nd: 0 });

                    const detailParts = [
                        `${totals.total} pedido${totals.total === 1 ? '' : 's'} en la consulta`
                    ];

                    if (totals.withWorkbook > 0) {
                        detailParts.push(`${totals.withWorkbook} con estado obtenido del Excel`);
                    }

                    if (totals.nd > 0) {
                        detailParts.push(`${totals.nd} marcado${totals.nd === 1 ? '' : 's'} como N/D`);
                    }

                    App.orderLookup.updateSummary(
                        'Consulta actualizada',
                        `${detailParts.join('. ')}.`,
                        totals.nd > 0 ? 'warning' : 'default'
                    );
                    App.orderLookup.renderRows(matches);
                },

                async loadWorkbook(file) {
                    if (!file) {
                        return;
                    }

                    if (!globalThis.XLSX) {
                        throw new Error('SheetJS no se pudo cargar correctamente. Recarga la pagina e intenta nuevamente.');
                    }

                    const buffer = await file.arrayBuffer();
                    const workbook = globalThis.XLSX.read(buffer, { type: 'array' });

                    if (!workbook.SheetNames.length) {
                        throw new Error('El archivo Excel no contiene hojas disponibles para consultar.');
                    }

                    const { rows } = App.orderLookup.extractWorkbookRows(workbook);
                    App.state.orderLookupRows = rows.filter(App.orderLookup.isMercadoLibreRow);
                    App.state.orderLookupDisplayColumns = App.orderLookup.getVisibleColumns();
                    App.orderLookup.setLoadedFileName(file.name);
                    if (App.dom.orderLookupInput) {
                        App.dom.orderLookupInput.value = '';
                    }
                    if (App.dom.orderLookupStateFilter) {
                        App.dom.orderLookupStateFilter.value = '';
                    }
                    if (App.dom.orderLookupRouteFilter) {
                        App.dom.orderLookupRouteFilter.value = '';
                    }
                    const referenceOrderDescriptors = App.orderLookup.getReferenceOrderDescriptors();
                    const referenceOrderNumbers = new Set(
                        referenceOrderDescriptors.map(descriptor => descriptor.orderNumber)
                    );
                    const referenceOrderTypeMap = new Map(
                        referenceOrderDescriptors
                            .filter(descriptor => descriptor.logisticsType)
                            .map(descriptor => [descriptor.orderNumber, descriptor.logisticsType])
                    );
                    const matchedRows = App.orderLookup.getMatchedExcelRows(
                        App.state.orderLookupRows,
                        referenceOrderNumbers,
                        referenceOrderTypeMap
                    );
                    const scopedRows = App.orderLookup.getScopedRows(
                        App.state.orderLookupRows,
                        referenceOrderNumbers,
                        referenceOrderTypeMap,
                        referenceOrderDescriptors
                    );
                    const missingRows = scopedRows.filter(row => App.helpers.normalizeComparable(row.estado) === 'N/D');

                    App.orderLookup.populateStateFilter(scopedRows);
                    App.orderLookup.populateRouteFilter(scopedRows);
                    App.orderLookup.applyFilters({ showWarningOnEmptySource: false, scopedRows });

                    let message = `Archivo "${file.name}" cargado correctamente. Se detectaron ${App.state.orderLookupRows.length} pedidos de Mercado Libre.`;
                    let messageType = 'success';

                    if (referenceOrderNumbers.size === 0) {
                        message = `Archivo "${file.name}" cargado correctamente. Se detectaron ${App.state.orderLookupRows.length} pedidos de Mercado Libre, pero aun no hay pedidos guardados o cargados para compararlos.`;
                        messageType = 'warning';
                    } else if (matchedRows.length === 0 && missingRows.length === 0) {
                        message = `Archivo "${file.name}" cargado correctamente. Se detectaron ${App.state.orderLookupRows.length} pedidos de Mercado Libre, pero ninguno coincide con los numeros guardados o cargados.`;
                        messageType = 'warning';
                    } else if (missingRows.length > 0) {
                        message = `Archivo "${file.name}" cargado correctamente. Se detectaron ${App.state.orderLookupRows.length} pedidos de Mercado Libre, ${matchedRows.length} coinciden con los numeros guardados o cargados y ${missingRows.length} se marcaron como N/D por no aparecer en el Excel.`;
                    } else {
                        message = `Archivo "${file.name}" cargado correctamente. Se detectaron ${App.state.orderLookupRows.length} pedidos de Mercado Libre y ${matchedRows.length} coinciden con los numeros guardados o cargados.`;
                    }

                    App.ui.showMessage(
                        message,
                        messageType,
                        { title: 'Consulta Pedido' }
                    );
                },

                applyFilters(options = {}) {
                    const { showWarningOnEmptySource = true, scopedRows = null } = options;
                    if (App.state.orderLookupRows.length === 0) {
                        if (showWarningOnEmptySource) {
                            App.ui.showMessage('Primero debes cargar un archivo Excel para consultar pedidos.', 'warning', { title: 'Consulta Pedido' });
                        }
                        return;
                    }

                    const referenceOrderNumbers = App.orderLookup.getReferenceOrderNumbers();
                    if (referenceOrderNumbers.size === 0) {
                        App.orderLookup.populateStateFilter([]);
                        App.orderLookup.populateRouteFilter([]);
                        App.storage.persistStoredOrderLookupState({ filters: App.orderLookup.getActiveFilters() });
                        App.orderLookup.updateReferenceScopeEmptyState('missingReference');
                        return;
                    }

                    const referenceOrderTypeMap = App.orderLookup.getReferenceOrderTypeMap();
                    const availableRows = Array.isArray(scopedRows)
                        ? scopedRows
                        : App.orderLookup.getScopedRows(App.state.orderLookupRows, referenceOrderNumbers, referenceOrderTypeMap);

                    App.orderLookup.populateStateFilter(availableRows);
                    App.orderLookup.populateRouteFilter(availableRows);
                    const filters = App.orderLookup.getActiveFilters();
                    App.storage.persistStoredOrderLookupState({ filters });

                    if (availableRows.length === 0) {
                        App.orderLookup.updateReferenceScopeEmptyState('noMatches');
                        return;
                    }

                    const matches = App.orderLookup.filterRows(filters, availableRows)
                        .map(row => ({
                            ...row,
                            dato2: App.helpers.normalizeComparable(row.dato2),
                            estado: App.helpers.normalizeComparable(row.estado) || 'Sin estado',
                            logisticsType: App.helpers.normalizeComparable(row.logisticsType)
                                || App.orderLookup.resolveLogisticsType(row, referenceOrderTypeMap)
                                || ''
                        }));

                    App.state.orderLookupMatches = matches;

                    if (matches.length === 0) {
                        App.orderLookup.updateNotFoundState(filters);
                        return;
                    }

                    App.orderLookup.updateFoundState(matches);
                },

                clearFilters() {
                    if (App.dom.orderLookupInput) {
                        App.dom.orderLookupInput.value = '';
                    }
                    if (App.dom.orderLookupStateFilter) {
                        App.dom.orderLookupStateFilter.value = '';
                    }
                    if (App.dom.orderLookupRouteFilter) {
                        App.dom.orderLookupRouteFilter.value = '';
                    }

                    App.orderLookup.applyFilters({ showWarningOnEmptySource: false });
                },

                syncReferenceScope(options = {}) {
                    if (App.state.orderLookupRows.length === 0) {
                        return;
                    }

                    const scopedRows = App.orderLookup.getScopedRows();
                    App.orderLookup.populateStateFilter(scopedRows);
                    App.orderLookup.populateRouteFilter(scopedRows);
                    App.orderLookup.applyFilters({
                        showWarningOnEmptySource: false,
                        ...options,
                        scopedRows
                    });
                }
            },

            table: {
                flattenDocumentRows(documents) {
                    return documents.flatMap(documentItem =>
                        documentItem.rows.map((row, index) => ({
                            ...row,
                            documentId: documentItem.id,
                            sourceName: documentItem.sourceName || 'Documento sin nombre',
                            storedAt: documentItem.storedAt,
                            storedLabel: App.storage.formatDateTime(documentItem.storedAt),
                            viewId: `${documentItem.id}-${index}-${row.numero}`
                        }))
                    );
                },

                buildRawZplFromRows(rows) {
                    return rows.map(row => row.zpl).join('\n');
                },

                getSelectedRows() {
                    return Array.from(App.state.selectedRowsById.values());
                },

                clearSelectedRows() {
                    App.state.selectedRowIds = new Set();
                    App.state.selectedRowsById = new Map();
                },

                setActiveBaseFromDocuments(documents, message, options = {}) {
                    const { selectedHistoryDocumentIds = null } = options;
                    App.state.activeBaseRows = App.table.flattenDocumentRows(documents);
                    App.state.activeBaseMessage = message;
                    App.state.selectedHistoryDocumentIds = selectedHistoryDocumentIds instanceof Set
                        ? new Set(selectedHistoryDocumentIds)
                        : Array.isArray(selectedHistoryDocumentIds)
                            ? new Set(selectedHistoryDocumentIds)
                            : new Set();
                    App.dom.searchInput.value = '';
                    App.table.renderResults(App.state.activeBaseRows, App.state.activeBaseMessage);
                    App.storage.updateStorageUI();
                    App.orderLookup.syncReferenceScope();
                },

                getSearchableRows() {
                    const combinedRows = [
                        ...App.state.activeBaseRows,
                        ...App.table.flattenDocumentRows(App.state.storedDocuments)
                    ];
                    const seen = new Set();

                    return combinedRows.filter(row => {
                        const key = `${row.documentId || 'sin-doc'}-${row.numero}-${row.zpl}`;
                        if (seen.has(key)) {
                            return false;
                        }

                        seen.add(key);
                        return true;
                    });
                },

                normalizeRowTypeFilter(filter) {
                    return filter === 'flex' || filter === 'colecta' ? filter : 'all';
                },

                getRowTypeFilterLabel(filter = App.state.activeRowTypeFilter) {
                    return filter === 'flex'
                        ? 'FLEX'
                        : filter === 'colecta'
                            ? 'COLECTA'
                            : 'TODOS';
                },

                countRowsByType(rows = App.state.currentResultRows) {
                    return rows.reduce((counts, row) => {
                        if (row.type === 'flex') {
                            counts.flex += 1;
                        } else if (row.type === 'colecta') {
                            counts.colecta += 1;
                        }

                        counts.total += 1;
                        return counts;
                    }, { flex: 0, colecta: 0, total: 0 });
                },

                getRowsForActiveTypeFilter(rows = App.state.currentResultRows) {
                    const filter = App.table.normalizeRowTypeFilter(App.state.activeRowTypeFilter);
                    if (filter === 'all') {
                        return rows;
                    }

                    return rows.filter(row => row.type === filter);
                },

                buildResultsMetaMessage(metaMessage, sourceRows, visibleRows) {
                    const filter = App.table.normalizeRowTypeFilter(App.state.activeRowTypeFilter);
                    if (filter === 'all' || sourceRows.length === 0) {
                        return metaMessage;
                    }

                    const filterMessage = `Filtro activo: ${App.table.getRowTypeFilterLabel(filter)} (${visibleRows.length} de ${sourceRows.length}).`;
                    return metaMessage ? `${metaMessage} ${filterMessage}` : filterMessage;
                },

                updateTypeFilterButtons(counts = App.table.countRowsByType()) {
                    const filter = App.table.normalizeRowTypeFilter(App.state.activeRowTypeFilter);
                    const buttonDefinitions = [
                        { button: App.dom.flexStatBtn, filter: 'flex', count: counts.flex },
                        { button: App.dom.colectaStatBtn, filter: 'colecta', count: counts.colecta },
                        { button: App.dom.totalStatBtn, filter: 'all', count: counts.total }
                    ];

                    buttonDefinitions.forEach(({ button, filter: buttonFilter, count }) => {
                        if (!button) {
                            return;
                        }

                        const isActive = filter === buttonFilter;
                        const shouldDim = filter !== 'all' && buttonFilter !== filter && count > 0;

                        button.classList.toggle('is-active', isActive);
                        button.classList.toggle('is-dimmed', shouldDim);
                        button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
                    });
                },

                setRowTypeFilter(nextFilter) {
                    if (App.state.currentResultRows.length === 0 && !App.state.currentResultMessage) {
                        return;
                    }

                    const normalizedFilter = App.table.normalizeRowTypeFilter(nextFilter);
                    App.state.activeRowTypeFilter = normalizedFilter === App.state.activeRowTypeFilter
                        ? 'all'
                        : normalizedFilter;
                    App.table.renderCurrentResults();
                },

                updateSelectionState() {
                    const rows = App.dom.tableBody.querySelectorAll('tr');

                    rows.forEach(tr => {
                        const rowId = tr.dataset.rowId;
                        const isSelected = App.state.selectedRowIds.has(rowId);
                        tr.classList.toggle('is-selected', isSelected);

                        const checkbox = tr.querySelector('input[type="checkbox"]');
                        if (checkbox) {
                            checkbox.checked = isSelected;
                        }
                    });

                    const selectedRows = App.table.getSelectedRows();
                    const visibleSelectedCount = App.state.extractedRows.filter(
                        row => App.state.selectedRowIds.has(row.viewId)
                    ).length;

                    if (selectedRows.length === 0) {
                        App.dom.selectedInfo.textContent = App.defaults.selectionMessage;
                        App.dom.printZebraBtn.disabled = App.state.extractedRows.length === 0;
                        return;
                    }

                    if (selectedRows.length === 1) {
                        const selectedRow = selectedRows[0];
                        App.dom.selectedInfo.textContent = visibleSelectedCount === 0
                            ? `1 bulto seleccionado: ${selectedRow.numero} (fuera del filtro actual).`
                            : `1 bulto seleccionado: ${selectedRow.numero}`;
                        App.dom.printZebraBtn.disabled = false;
                        return;
                    }

                    const hiddenSelectedCount = selectedRows.length - visibleSelectedCount;
                    App.dom.selectedInfo.textContent = hiddenSelectedCount > 0
                        ? `${selectedRows.length} bultos seleccionados (${hiddenSelectedCount} fuera del filtro actual).`
                        : `${selectedRows.length} bultos seleccionados para impresion.`;
                    App.dom.printZebraBtn.disabled = false;
                },

                toggleRowSelection(row, forceSelected) {
                    if (!row || !row.viewId) {
                        return;
                    }

                    const shouldSelect = typeof forceSelected === 'boolean'
                        ? forceSelected
                        : !App.state.selectedRowIds.has(row.viewId);

                    if (shouldSelect) {
                        App.state.selectedRowIds.add(row.viewId);
                        App.state.selectedRowsById.set(row.viewId, row);
                    } else {
                        App.state.selectedRowIds.delete(row.viewId);
                        App.state.selectedRowsById.delete(row.viewId);
                    }

                    App.table.updateSelectionState();
                },

                renderResults(rows, metaMessage) {
                    App.state.currentResultRows = Array.isArray(rows) ? [...rows] : [];
                    App.state.currentResultMessage = metaMessage || '';
                    App.table.renderCurrentResults();
                },

                renderCurrentResults() {
                    const sourceRows = App.state.currentResultRows.map((row, index) => ({
                        ...row,
                        viewId: row.viewId || `view-${index}-${row.numero}`
                    }));
                    const counts = App.table.countRowsByType(sourceRows);
                    const visibleRows = App.table.getRowsForActiveTypeFilter(sourceRows);
                    const activeFilter = App.table.normalizeRowTypeFilter(App.state.activeRowTypeFilter);
                    const hasActiveTypeFilter = activeFilter !== 'all';

                    App.state.extractedRows = visibleRows;
                    App.state.rawZplData = App.table.buildRawZplFromRows(App.state.extractedRows);
                    App.dom.tableBody.innerHTML = '';

                    App.state.extractedRows.forEach(row => {
                        if (App.state.selectedRowIds.has(row.viewId)) {
                            App.state.selectedRowsById.set(row.viewId, row);
                        }
                    });

                    App.state.extractedRows.forEach(row => {
                        const tr = document.createElement('tr');
                        tr.dataset.rowId = row.viewId;
                        tr.innerHTML = `
                            <td><input class="row-selector" type="checkbox" aria-label="Seleccionar ${App.helpers.escapeHtml(row.numero)}"></td>
                            <td>${App.helpers.escapeHtml(row.proceso)}</td>
                            <td><strong>${App.helpers.escapeHtml(row.numero)}</strong></td>
                            <td>${App.helpers.escapeHtml(row.picker)}</td>
                            <td>${App.helpers.escapeHtml(row.picking)}</td>
                            <td><span style="color: ${row.type === 'flex' ? 'var(--primary)' : 'var(--success)'}">${App.helpers.escapeHtml(row.revision)}</span></td>
                            <td>
                                <div class="meta-cell">
                                    <strong>${App.helpers.escapeHtml(row.storedLabel || 'Carga actual')}</strong>
                                    <span>${App.helpers.escapeHtml(row.sourceName || 'Documento actual')}</span>
                                </div>
                            </td>
                        `;

                        tr.addEventListener('click', () => {
                            App.table.toggleRowSelection(row);
                        });

                        const checkbox = tr.querySelector('input[type="checkbox"]');
                        checkbox.addEventListener('click', event => {
                            event.stopPropagation();
                            App.table.toggleRowSelection(row, checkbox.checked);
                        });

                        App.dom.tableBody.appendChild(tr);
                    });

                    App.dom.countFlex.textContent = String(counts.flex);
                    App.dom.countColecta.textContent = String(counts.colecta);
                    App.dom.countTotal.textContent = String(counts.total);
                    App.table.updateTypeFilterButtons(counts);

                    App.ui.setResultsVisible(true);
                    App.ui.setClearButtonVisible(true);
                    App.ui.updateResultsMeta(
                        App.table.buildResultsMetaMessage(
                            App.state.currentResultMessage,
                            sourceRows,
                            App.state.extractedRows
                        )
                    );

                    if (App.state.extractedRows.length === 0) {
                        App.table.resetActionButtons(true);
                        App.table.updateSelectionState();
                        if (App.state.selectedRowsById.size === 0) {
                            App.dom.selectedInfo.textContent = hasActiveTypeFilter
                                ? `No se encontraron bultos ${App.table.getRowTypeFilterLabel(activeFilter)} para mostrar.`
                                : 'No se encontraron numeros para mostrar.';
                        }
                        return;
                    }

                    App.dom.downloadBtn.disabled = false;
                    App.dom.printZebraBtn.disabled = false;
                    App.dom.downloadBtn.innerHTML = App.table.getDownloadButtonMarkup();
                    App.table.updateSelectionState();
                },

                resetWorkspaceView(resetInputs) {
                    App.state.activeBaseRows = [];
                    App.state.activeBaseMessage = '';
                    App.state.currentResultRows = [];
                    App.state.currentResultMessage = '';
                    App.state.activeRowTypeFilter = 'all';
                    App.state.selectedHistoryDocumentIds = new Set();
                    App.state.extractedRows = [];
                    App.state.rawZplData = '';
                    App.table.clearSelectedRows();
                    App.dom.tableBody.innerHTML = '';
                    App.dom.countFlex.textContent = '0';
                    App.dom.countColecta.textContent = '0';
                    App.dom.countTotal.textContent = '0';
                    App.table.updateTypeFilterButtons({ flex: 0, colecta: 0, total: 0 });
                    App.ui.setResultsVisible(false);
                    App.ui.setClearButtonVisible(false);
                    App.dom.searchInput.value = '';
                    App.dom.selectedInfo.textContent = App.defaults.selectionMessage;
                    App.ui.updateResultsMeta(App.defaults.emptyResultsMessage);
                    App.table.resetActionButtons();

                    if (resetInputs) {
                        App.dom.fileInput.value = '';
                        App.dom.textInput.value = '';
                        App.dom.pickerInput.value = '';
                    }

                    App.storage.updateStorageUI();
                    App.orderLookup.syncReferenceScope();
                },

                getDownloadButtonMarkup() {
                    return '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg> Descargar Planilla CSV';
                },

                resetActionButtons(emptyState = false) {
                    App.dom.downloadBtn.disabled = true;
                    App.dom.printZebraBtn.disabled = App.state.extractedRows.length === 0 && App.state.selectedRowsById.size === 0;
                    App.dom.downloadBtn.innerHTML = emptyState
                        ? 'No se encontraron numeros'
                        : App.table.getDownloadButtonMarkup();
                },

                downloadCSV(data, filename) {
                    if (data.length === 0) {
                        return;
                    }

                    let csvContent = 'Proceso;NUMERO;PICKER;PICKING;REVISION\r\n';
                    data.forEach(row => {
                        csvContent += `${row.proceso};${row.numero};${row.picker};${row.picking};${row.revision}\r\n`;
                    });

                    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                    const url = URL.createObjectURL(blob);
                    const link = document.createElement('a');
                    link.setAttribute('href', url);
                    link.setAttribute('download', filename);
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    URL.revokeObjectURL(url);
                },

                handleSearch() {
                    const query = App.dom.searchInput.value.trim().toLowerCase();

                    if (!query) {
                        if (App.state.activeBaseRows.length === 0) {
                            if (App.state.selectedRowsById.size > 0) {
                                App.table.renderResults(
                                    [],
                                    'La seleccion actual se conserva. Realiza otra busqueda para seguir agregando bultos.'
                                );
                                return;
                            }

                            App.table.resetWorkspaceView(false);
                            return;
                        }

                        App.table.renderResults(
                            App.state.activeBaseRows,
                            App.state.activeBaseMessage || 'Mostrando la ultima carga procesada.'
                        );
                        return;
                    }

                    const matchedRows = App.table.getSearchableRows().filter(
                        row => row.numero.toLowerCase().includes(query)
                    );

                    if (matchedRows.length === 0) {
                        App.table.renderResults(
                            [],
                            `No se encontraron coincidencias guardadas para "${App.dom.searchInput.value.trim()}".`
                        );
                        return;
                    }

                    const documentCount = new Set(matchedRows.map(row => row.documentId)).size;
                    App.table.renderResults(
                        matchedRows,
                        `Mostrando ${matchedRows.length} coincidencia(s) encontradas en ${documentCount} documento(s) guardado(s).`
                    );
                },

                showStoredDocuments() {
                    if (App.state.storedDocuments.length === 0) {
                        App.ui.showMessage('Todavia no hay documentos guardados para mostrar.', 'info', { title: 'Historial' });
                        return;
                    }

                    App.table.setActiveBaseFromDocuments(
                        App.state.storedDocuments,
                        `Mostrando ${App.state.storedDocuments.length} documento(s) guardado(s) en el historial local.`
                    );
                }
            },
            actions: {
                preventDefaults(event) {
                    event.preventDefault();
                    event.stopPropagation();
                },

                validateProcessingReady(actionLabel) {
                    if (!App.runtime.storage) {
                        App.ui.showMessage(
                            'No fue posible habilitar el historial local del navegador. Sin ese historial no se puede guardar ni buscar documentos previos.',
                            'error',
                            { title: 'Historial' }
                        );
                        return false;
                    }

                    if (!App.dom.pickerInput.value.trim()) {
                        App.ui.showMessage(
                            `Por favor, selecciona un usuario antes de ${actionLabel}.`,
                            'warning',
                            { title: 'Usuario' }
                        );
                        App.dom.pickerInput.focus();
                        return false;
                    }

                    return true;
                },

                appendDocumentsToHistory(documents) {
                    const nextDocuments = [...documents, ...App.state.storedDocuments].sort(
                        (a, b) => new Date(b.storedAt) - new Date(a.storedAt)
                    );

                    return App.storage.persistStoredDocuments(nextDocuments);
                },

                restoreStoredDocumentsOnOpen() {
                    if (App.state.storedDocuments.length === 0) {
                        return;
                    }

                    App.table.setActiveBaseFromDocuments(
                        App.state.storedDocuments,
                        `Historial recuperado automaticamente al abrir el archivo (${App.state.storedDocuments.length} documento(s) disponibles).`
                    );
                },

                async handleDrop(event) {
                    const dataTransfer = event.dataTransfer;
                    if (!dataTransfer) {
                        return;
                    }

                    await App.actions.handleFiles(dataTransfer.files);
                },

                async handlePastedText() {
                    if (!App.actions.validateProcessingReady('procesar el texto')) {
                        return;
                    }

                    const text = App.dom.textInput.value.trim();
                    if (!text) {
                        App.ui.showMessage('Por favor, pega el contenido del archivo en el area de texto.', 'warning', { title: 'Texto' });
                        App.dom.textInput.focus();
                        return;
                    }

                    const picker = App.dom.pickerInput.value.trim();
                    const shouldPersist = App.storage.canStoreNow();
                    const documentRecord = App.parser.createDocumentRecord({
                        picker,
                        sourceName: `Texto pegado ${App.storage.formatDateTime(new Date())}`,
                        sourceType: 'text',
                        text
                    });

                    if (!documentRecord) {
                        App.ui.showMessage('No se encontraron numeros validos dentro del texto pegado.', 'warning', { title: 'Texto' });
                        return;
                    }

                    if (shouldPersist && !App.actions.appendDocumentsToHistory([documentRecord])) {
                        return;
                    }

                    App.dom.fileInput.value = '';
                    App.dom.textInput.value = '';
                    App.table.setActiveBaseFromDocuments(
                        [documentRecord],
                        shouldPersist
                            ? `Mostrando 1 documento recien guardado (${documentRecord.rows.length} numero(s)).`
                            : `Mostrando 1 documento procesado (${documentRecord.rows.length} numero(s)). Fuera del horario de historial, por eso no se guardo.`
                    );
                },

                async handleFiles(files) {
                    if (!App.actions.validateProcessingReady('cargar archivos')) {
                        App.dom.fileInput.value = '';
                        return;
                    }

                    const txtFiles = Array.from(files || []).filter(
                        file => file.name.toLowerCase().endsWith('.txt')
                    );

                    if (txtFiles.length === 0) {
                        App.ui.showMessage('Solo se admiten archivos .txt de MercadoLibre.', 'warning', { title: 'Archivos' });
                        App.dom.fileInput.value = '';
                        return;
                    }

                    const picker = App.dom.pickerInput.value.trim();
                    const shouldPersist = App.storage.canStoreNow();
                    const newDocuments = [];
                    const knownFingerprints = new Set(
                        App.state.storedDocuments
                            .map(documentItem => App.parser.getDocumentFingerprint(documentItem))
                            .filter(Boolean)
                    );
                    const duplicateFiles = [];

                    for (const file of txtFiles) {
                        const text = await App.parser.readFileAsText(file);
                        const documentRecord = App.parser.createDocumentRecord({
                            picker,
                            sourceName: file.name,
                            sourceType: 'file',
                            text
                        });

                        if (!documentRecord) {
                            App.ui.showMessage(
                                `No se encontraron numeros validos en el archivo "${file.name}".`,
                                'warning',
                                { title: 'Archivos' }
                            );
                            App.dom.fileInput.value = '';
                            return;
                        }

                        if (documentRecord.fingerprint && knownFingerprints.has(documentRecord.fingerprint)) {
                            duplicateFiles.push(file.name);
                            continue;
                        }

                        if (documentRecord.fingerprint) {
                            knownFingerprints.add(documentRecord.fingerprint);
                        }

                        newDocuments.push(documentRecord);
                    }

                    if (newDocuments.length === 0) {
                        if (duplicateFiles.length > 0) {
                            App.ui.showMessage(
                                App.parser.buildDuplicateFilesMessage(duplicateFiles),
                                'warning',
                                { title: 'Archivos duplicados', timeout: 9000 }
                            );
                        }

                        App.dom.fileInput.value = '';
                        return;
                    }

                    if (shouldPersist && !App.actions.appendDocumentsToHistory(newDocuments)) {
                        App.dom.fileInput.value = '';
                        return;
                    }

                    App.dom.textInput.value = '';
                    App.dom.fileInput.value = '';

                    const duplicateMessage = duplicateFiles.length > 0
                        ? ` Se omitieron ${duplicateFiles.length} archivo(s) ya guardado(s).`
                        : '';

                    App.table.setActiveBaseFromDocuments(
                        newDocuments,
                        shouldPersist
                            ? `Mostrando ${newDocuments.length} documento(s) recien guardado(s) con ${newDocuments.reduce((sum, documentItem) => sum + documentItem.rows.length, 0)} numero(s).${duplicateMessage}`
                            : `Mostrando ${newDocuments.length} documento(s) procesado(s) con ${newDocuments.reduce((sum, documentItem) => sum + documentItem.rows.length, 0)} numero(s). Fuera del horario de historial, por eso no se guardaron.${duplicateMessage}`
                    );

                    if (duplicateFiles.length > 0) {
                        App.ui.showMessage(
                            App.parser.buildDuplicateFilesMessage(duplicateFiles),
                            'warning',
                            { title: 'Archivos duplicados', timeout: 9000 }
                        );
                    }
                },

                clearStoredHistory() {
                    if (!App.runtime.storage || App.state.storedDocuments.length === 0) {
                        return;
                    }

                    const confirmed = window.confirm(
                        'Se eliminara el historial guardado de este navegador. Deseas continuar?'
                    );
                    if (!confirmed) {
                        return;
                    }

                    if (!App.storage.persistStoredDocuments([])) {
                        return;
                    }

                    App.state.activeBaseRows = [];
                    App.state.activeBaseMessage = '';

                    if (App.dom.searchInput.value.trim()) {
                        App.table.renderResults([], 'El historial fue limpiado. No hay coincidencias para mostrar.');
                        return;
                    }

                    App.table.resetWorkspaceView(false);
                },

                async handlePortableSave() {
                    if (!App.storage.canSavePortableState()) {
                        App.ui.showMessage(
                            'No hay historial ni usuarios personalizados para guardar dentro del HTML portable.',
                            'info',
                            { title: 'Portable' }
                        );
                        return;
                    }

                    App.storage.writeEmbeddedState({
                        documents: App.state.storedDocuments,
                        users: App.state.pickerOptions
                    });

                    const portableHtml = App.storage.buildPortableHtml();
                    const timestamp = new Date();
                    const filename = `${App.config.PORTABLE_FILENAME_PREFIX}-${timestamp.getFullYear()}-${String(timestamp.getMonth() + 1).padStart(2, '0')}-${String(timestamp.getDate()).padStart(2, '0')}-${String(timestamp.getHours()).padStart(2, '0')}${String(timestamp.getMinutes()).padStart(2, '0')}.html`;

                    if (window.showSaveFilePicker) {
                        try {
                            const fileHandle = await window.showSaveFilePicker({
                                suggestedName: filename,
                                types: [
                                    {
                                        description: 'Archivo HTML',
                                        accept: {
                                            'text/html': ['.html']
                                        }
                                    }
                                ]
                            });
                            const writable = await fileHandle.createWritable();
                            await writable.write(portableHtml);
                            await writable.close();
                            App.ui.showMessage(
                                'Se guardo una copia portable del HTML con el historial y los usuarios actuales.',
                                'success',
                                { title: 'Portable' }
                            );
                            return;
                        } catch (error) {
                            if (error && error.name === 'AbortError') {
                                return;
                            }

                            App.ui.reportError(error, 'No fue posible guardar la copia portable.', 'Portable');
                            return;
                        }
                    }

                    const blob = new Blob([portableHtml], { type: 'text/html;charset=utf-8;' });
                    const url = URL.createObjectURL(blob);
                    const link = document.createElement('a');
                    link.setAttribute('href', url);
                    link.setAttribute('download', filename);
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    URL.revokeObjectURL(url);
                    App.ui.showMessage(
                        'Se descargo una copia portable del HTML con el historial y los usuarios actuales.',
                        'success',
                        { title: 'Portable' }
                    );
                },

                handleCsvDownload() {
                    if (App.state.extractedRows.length === 0) {
                        return;
                    }

                    const flexData = App.state.extractedRows.filter(row => row.type === 'flex');
                    const colectaData = App.state.extractedRows.filter(row => row.type === 'colecta');

                    if (flexData.length > 0) {
                        App.table.downloadCSV(flexData, '1 f.csv');
                    }
                    if (colectaData.length > 0) {
                        window.setTimeout(() => {
                            App.table.downloadCSV(colectaData, '1 c.csv');
                        }, 300);
                    }
                },

                async handleZebraPrint() {
                    const selectedRows = App.table.getSelectedRows();
                    const zplToPrint = selectedRows.length > 0
                        ? App.table.buildRawZplFromRows(selectedRows)
                        : App.state.rawZplData;

                    if (!zplToPrint) {
                        return;
                    }

                    const originalText = App.dom.printZebraBtn.innerHTML;
                    App.dom.printZebraBtn.innerHTML = 'Buscando Zebra...';
                    App.dom.printZebraBtn.disabled = true;

                    try {
                        let printer = null;

                        try {
                            const resDefault = await fetch('http://127.0.0.1:9100/default');
                            if (resDefault.ok) {
                                const defaultText = await resDefault.text();
                                if (defaultText && defaultText.includes('uid')) {
                                    try {
                                        printer = JSON.parse(defaultText);
                                    } catch (error) {
                                        printer = null;
                                    }
                                }
                            }
                        } catch (error) {
                            printer = null;
                        }

                        if (!printer || typeof printer !== 'object') {
                            const resAvailable = await fetch('http://127.0.0.1:9100/available');
                            if (!resAvailable.ok) {
                                throw new Error('No se pudo conectar a Zebra Browser Print.');
                            }

                            const devices = await resAvailable.json();
                            if (devices.printer && devices.printer.length > 0) {
                                printer = devices.printer[0];
                            }
                        }

                        if (!printer || typeof printer !== 'object' || !printer.uid) {
                            throw new Error(
                                'No se encontro ninguna impresora Zebra disponible.\nAsegurate de que este encendida, con papel y conectada por cable o red.'
                            );
                        }

                        App.dom.printZebraBtn.innerHTML = 'Enviando a Zebra...';

                        const resWrite = await fetch('http://127.0.0.1:9100/write', {
                            method: 'POST',
                            body: JSON.stringify({
                                device: printer,
                                data: zplToPrint
                            })
                        });

                        if (!resWrite.ok) {
                            let errText = 'Error desconocido';
                            try {
                                errText = await resWrite.text();
                            } catch (error) {
                                errText = 'Error desconocido';
                            }

                            throw new Error(
                                `Se encontro la impresora, pero fallo la transmision del texto hacia ella.\nDetalle: ${resWrite.status} ${errText}`
                            );
                        }

                        if (selectedRows.length > 0) {
                            App.table.clearSelectedRows();
                            App.table.updateSelectionState();
                        }

                        App.ui.showMessage(
                            'Las etiquetas fueron enviadas a la impresora Zebra exitosamente.',
                            'success',
                            { title: 'Impresion Zebra' }
                        );
                    } catch (error) {
                        App.ui.reportError(
                            error,
                            'Error de impresion Zebra.\n\n1. Instala Zebra Browser Print desde la web oficial.\n2. Verifica que el icono de Zebra este activo junto al reloj de Windows.\n3. Autoriza el localhost con "Yes" la primera vez que se ejecute Desktop Print.',
                            'Impresion Zebra'
                        );
                    } finally {
                        App.dom.printZebraBtn.innerHTML = originalText;
                        App.dom.printZebraBtn.disabled = App.state.extractedRows.length === 0 && App.state.selectedRowsById.size === 0;
                    }
                },

                bindEvents() {
                    App.dom.dropzone.addEventListener('click', () => {
                        if (App.runtime.storage) {
                            App.dom.fileInput.click();
                        }
                    });

                    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
                        App.dom.dropzone.addEventListener(eventName, App.actions.preventDefaults, false);
                    });

                    ['dragenter', 'dragover'].forEach(eventName => {
                        App.dom.dropzone.addEventListener(eventName, () => {
                            if (App.runtime.storage) {
                                App.dom.dropzone.classList.add('dragover');
                            }
                        }, false);
                    });

                    ['dragleave', 'drop'].forEach(eventName => {
                        App.dom.dropzone.addEventListener(eventName, () => {
                            App.dom.dropzone.classList.remove('dragover');
                        }, false);
                    });

                    App.dom.dropzone.addEventListener('drop', async event => {
                        try {
                            await App.actions.handleDrop(event);
                        } catch (error) {
                            App.ui.reportError(error, 'No fue posible procesar los archivos seleccionados.', 'Archivos');
                            App.dom.fileInput.value = '';
                        }
                    }, false);

                    App.dom.fileInput.addEventListener('change', async event => {
                        try {
                            await App.actions.handleFiles(event.target.files);
                        } catch (error) {
                            App.ui.reportError(error, 'No fue posible procesar los archivos seleccionados.', 'Archivos');
                            App.dom.fileInput.value = '';
                        }
                    }, false);

                    App.dom.manageUsersBtn.addEventListener('click', () => App.users.togglePickerManager());
                    App.dom.addPickerBtn.addEventListener('click', App.users.handleAddPicker);
                    App.dom.newPickerInput.addEventListener('keydown', event => {
                        if (event.key === 'Enter') {
                            event.preventDefault();
                            App.users.handleAddPicker();
                        }
                    });

                    App.dom.pickerList.addEventListener('click', event => {
                        const removeButton = event.target.closest('[data-remove-picker]');
                        if (!removeButton) {
                            return;
                        }

                        App.users.removePickerOption(removeButton.dataset.removePicker || '');
                    });

                    App.dom.historyList.addEventListener('click', event => {
                        const historyButton = event.target.closest('[data-history-document-id]');
                        if (!historyButton) {
                            return;
                        }

                        App.storage.toggleHistoryDocumentSelection(historyButton.dataset.historyDocumentId || '');
                    });

                    App.dom.searchInput.addEventListener('input', App.table.handleSearch);
                    App.dom.flexStatBtn?.addEventListener('click', () => App.table.setRowTypeFilter('flex'));
                    App.dom.colectaStatBtn?.addEventListener('click', () => App.table.setRowTypeFilter('colecta'));
                    App.dom.totalStatBtn?.addEventListener('click', () => App.table.setRowTypeFilter('all'));
                    App.dom.outsideHoursToggleBtn?.addEventListener('click', App.storage.toggleOutsideHoursHistory);
                    App.dom.showStoredBtn.addEventListener('click', App.table.showStoredDocuments);
                    App.dom.clearStoredBtn.addEventListener('click', App.actions.clearStoredHistory);
                    App.dom.savePortableBtn.addEventListener('click', App.actions.handlePortableSave);
                    App.dom.clearBtn.addEventListener('click', () => App.table.resetWorkspaceView(true));
                    App.dom.processTextBtn.addEventListener('click', async () => {
                        try {
                            await App.actions.handlePastedText();
                        } catch (error) {
                            App.ui.reportError(error, 'No fue posible procesar el texto pegado.', 'Texto');
                        }
                    });

                    App.dom.orderLookupLoadBtn?.addEventListener('click', () => {
                        App.dom.orderLookupFileInput?.click();
                    });

                    App.dom.orderLookupFileInput?.addEventListener('change', async event => {
                        const file = event.target.files?.[0];
                        if (!file) {
                            return;
                        }

                        if (!file.name.toLowerCase().endsWith('.xlsx')) {
                            App.ui.showMessage('Solo se admiten archivos Excel con extension .xlsx.', 'warning', { title: 'Consulta Pedido' });
                            event.target.value = '';
                            return;
                        }

                        try {
                            await App.orderLookup.loadWorkbook(file);
                        } catch (error) {
                            App.ui.reportError(error, 'No fue posible cargar el archivo Excel.', 'Consulta Pedido');
                            App.orderLookup.resetData();
                            App.orderLookup.resetResults();
                        } finally {
                            event.target.value = '';
                        }
                    });

                    App.dom.orderLookupSearchBtn?.addEventListener('click', () => {
                        App.orderLookup.applyFilters();
                    });

                    App.dom.orderLookupInput?.addEventListener('keydown', event => {
                        if (event.key === 'Enter') {
                            event.preventDefault();
                            App.orderLookup.applyFilters();
                        }
                    });

                    App.dom.orderLookupStateFilter?.addEventListener('change', () => {
                        App.orderLookup.applyFilters({ showWarningOnEmptySource: false });
                    });

                    App.dom.orderLookupRouteFilter?.addEventListener('change', () => {
                        App.orderLookup.applyFilters({ showWarningOnEmptySource: false });
                    });

                    App.dom.orderLookupClearFiltersBtn?.addEventListener('click', () => {
                        App.orderLookup.clearFilters();
                    });

                    App.dom.downloadBtn.addEventListener('click', App.actions.handleCsvDownload);
                    App.dom.printZebraBtn.addEventListener('click', App.actions.handleZebraPrint);
                },

                startTimers() {
                    window.setInterval(() => {
                        App.storage.pruneHistoryBySchedule();
                        App.storage.updateStorageUI();
                        App.storage.updateWindowControls();
                    }, 60000);

                    window.setInterval(() => {
                        App.storage.refreshStorageStatus(new Date());
                    }, 1000);
                }
            },

            init() {
                App.runtime.storage = App.storage.getBrowserStorage();
                App.runtime.defaultPickerOptions = Array.from(App.dom.pickerInput.querySelectorAll('option'))
                    .map(option => option.value.trim())
                    .filter(Boolean);

                App.state.embeddedDocuments = App.storage.loadEmbeddedDocuments();
                App.state.outsideHoursHistoryEnabled = App.storage.loadOutsideHoursHistoryPreference();
                App.state.pickerOptions = App.users.loadPickerOptions();
                App.state.storedDocuments = App.storage.getVisibleStoredDocuments(
                    App.storage.loadStoredDocuments(),
                    App.state.embeddedDocuments
                );

                if (App.dom.pasteTitle) {
                    App.dom.pasteTitle.textContent = 'Pegar Texto';
                }

                App.users.renderPickerOptions();
                App.users.renderPickerManager();
                App.dom.selectedInfo.textContent = App.defaults.selectionMessage;
                App.orderLookup.setLoadedFileName('');
                App.orderLookup.resetResults();

                App.storage.pruneHistoryBySchedule();
                App.ui.updateResultsMeta(App.defaults.emptyResultsMessage);
                App.storage.updateStorageUI();
                App.storage.updateWindowControls();
                App.storage.refreshStorageStatus(new Date());
                App.table.resetActionButtons();
                App.table.updateTypeFilterButtons({ flex: 0, colecta: 0, total: 0 });
                App.actions.restoreStoredDocumentsOnOpen();
                App.orderLookup.restoreStoredState();
                App.actions.bindEvents();
                App.actions.startTimers();
            }
        };

        App.init();
        window.App = App;
    }

    bootstrap();
})();
