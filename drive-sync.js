/**
 * drive-sync.js
 * Monitorea Google Drive (Shared Drive) y procesa archivos .txt de etiquetas ML.
 * Modos: Automático (kmendoza) | Manual (usuario elige y procesa)
 */

const DRIVE_CONFIG = {
    CLIENT_ID:      '539650020603-dcgj1e8u6ig305qosef3p49tkc8ml551.apps.googleusercontent.com',
    API_KEY:        null,
    ROOT_FOLDER_ID: '0AGCxqv4-HTgIUk9PVA',
    POLL_INTERVAL:  30_000,
    SCOPES:         'https://www.googleapis.com/auth/drive.readonly',
    DEFAULT_PICKER: 'kmendoza',
};

const MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
               'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

function _getApiKey() {
    const saved = localStorage.getItem('drive_api_key');
    if (saved) return saved;
    const entered = prompt('Ingresá tu Google API Key:\n(Se guardará para no pedirla de nuevo)');
    if (entered?.trim()) { localStorage.setItem('drive_api_key', entered.trim()); return entered.trim(); }
    return null;
}

const _state = {
    mode:         localStorage.getItem('drive_mode') || 'auto',
    isPolling:    false,
    isAuthorized: false,
    pollTimer:    null,
    knownIds:     new Set(),
    knownFiles:   [],
    pending:      [],
    processed:    [],
    tokenClient:  null,
    panelOpen:    false,
    browseDate:   new Date(),  /* fecha del panel de exploración */
    browseFiles:  null,          /* null = usa knownFiles del día actual */
};

function _getMes(d = new Date()) { return `${MESES[d.getMonth()]} ${d.getFullYear()}`; }
function _getDia(d = new Date()) { return `${d.getDate()} ${MESES[d.getMonth()]}`; }
function _timeStr(d = new Date()) { return d.toLocaleTimeString('es-CL', { hour:'2-digit', minute:'2-digit' }); }

function _setStatus(text, active = false) {
    const s = document.getElementById('drive-sync-status');
    if (!s) return;
    s.textContent = text;
    s.style.color = active ? '#00ac47' : '';
    s.style.fontWeight = active ? '600' : '';
}

function _setFileCount(n) {
    const el = document.getElementById('drive-file-count');
    if (!el) return;
    if (n === null || n === undefined) { el.style.display = 'none'; return; }
    el.style.display = '';
    el.textContent = '\uD83D\uDCC1 ' + n + ' archivo' + (n !== 1 ? 's' : '');
    el.style.color = n > 0 ? 'var(--text)' : 'var(--text-muted)';
    el.style.cursor = 'pointer';
    el.title = 'Ver archivos del día';
    el.onclick = () => DriveSync.togglePanel();
}

function _highlightNewestHistorialItem() {
    setTimeout(function() {
        var items = document.querySelectorAll('#historyList [data-history-document-id]');
        if (!items.length) return;
        var newest = items[0];
        newest.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        var prev = newest.style.background;
        newest.style.transition = 'background 0.4s';
        newest.style.background = '#e8f5e9';
        setTimeout(function() {
            newest.style.transition = 'background 1s';
            newest.style.background = prev || '';
        }, 1800);
        if (!newest.querySelector('.drive-badge')) {
            var badge = document.createElement('span');
            badge.className = 'drive-badge';
            badge.title = 'Procesado desde Google Drive';
            badge.style.cssText = 'display:inline-flex;align-items:center;margin-left:5px;vertical-align:middle;opacity:0.8;';
            badge.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="10" viewBox="0 0 87.3 78">'
                + '<path d="m6.6 66.85 3.85 6.65c.8 1.4 1.95 2.5 3.3 3.3l13.75-23.8h-27.5c0 1.55.4 3.1 1.2 4.5z" fill="#0066da"/>'
                + '<path d="m43.65 25-13.75-23.8c-1.35.8-2.5 1.9-3.3 3.3l-25.4 44a9.06 9.06 0 0 0-1.2 4.5h27.5z" fill="#00ac47"/>'
                + '<path d="m73.55 76.8c1.35-.8 2.5-1.9 3.3-3.3l1.6-2.75 7.65-13.25c.8-1.4 1.2-2.95 1.2-4.5h-27.502l5.852 11.5z" fill="#ea4335"/>'
                + '<path d="m43.65 25 13.75-23.8c-1.35-.8-2.9-1.2-4.5-1.2h-18.5c-1.6 0-3.15.45-4.5 1.2z" fill="#00832d"/>'
                + '<path d="m59.8 53h-32.3l-13.75 23.8c1.35.8 2.9 1.2 4.5 1.2h50.8c1.6 0 3.15-.45 4.5-1.2z" fill="#2684fc"/>'
                + '<path d="m73.4 26.5-12.7-22c-.8-1.4-1.95-2.5-3.3-3.3l-13.75 23.8 16.15 27h27.45c0-1.55-.4-3.1-1.2-4.5z" fill="#ffba00"/>'
                + '</svg>';
            var nameEl = newest.querySelector('strong, [class*="name"], [class*="source"]') || newest.firstElementChild || newest;
            nameEl.appendChild(badge);
        }
    }, 700);
}

function _setToast(text, ok = true) {
    const t = document.getElementById('drive-sync-toast');
    if (!t) return;
    t.textContent = text;
    t.style.borderLeftColor = ok ? '#00ac47' : '#ea4335';
    t.style.display = 'block';
    setTimeout(() => { if (t.textContent === text) t.style.display = 'none'; }, 6000);
}

function _renderPending() {
    const list    = document.getElementById('drive-pending-list');
    const noPend  = document.getElementById('drive-no-pending');
    const section = document.getElementById('drive-pending-section');
    if (!list) return;

    if (_state.mode !== 'manual') { section.style.display = 'none'; return; }
    section.style.display = 'block';

    [...list.querySelectorAll('.drive-pending-row')].forEach(r => r.remove());

    if (_state.pending.length === 0) { noPend.style.display = 'block'; return; }
    noPend.style.display = 'none';

    const pickerSel = document.getElementById('picker');
    const opts = pickerSel
        ? Array.from(pickerSel.options).filter(o => o.value).map(o => ({ v: o.value, l: o.text }))
        : [{ v: DRIVE_CONFIG.DEFAULT_PICKER, l: DRIVE_CONFIG.DEFAULT_PICKER }];

    for (const item of _state.pending) {
        const row = document.createElement('div');
        row.className = 'drive-pending-row';
        row.dataset.id = item.id;
        row.style.cssText = `background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:10px 12px;display:flex;flex-direction:column;gap:6px;font-size:12px;`;

        const optHtml = opts.map(o =>
            `<option value="${o.v}" ${o.v.toLowerCase().includes(DRIVE_CONFIG.DEFAULT_PICKER) ? 'selected' : ''}>${o.l}</option>`
        ).join('');

        row.innerHTML = `
            <div style="display:flex;justify-content:space-between;align-items:center;">
                <span style="font-weight:600;color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:155px;" title="${item.name}">📄 ${item.name}</span>
                <span style="color:var(--text-muted);font-size:11px;flex-shrink:0;">${_timeStr(item.detectedAt)}</span>
            </div>
            <div style="display:flex;gap:4px;">
                <span style="background:#e8f5e9;color:#2e7d32;border-radius:4px;padding:2px 7px;font-size:11px;">${item.folder || 'Drive'}</span>
                <span style="background:#e3f2fd;color:#1565c0;border-radius:4px;padding:2px 7px;font-size:11px;">${item.bulkCount} bulto${item.bulkCount!==1?'s':''}</span>
            </div>
            <div style="display:flex;gap:6px;align-items:center;">
                <select id="picker-${item.id}" style="flex:1;padding:5px 8px;border:1px solid var(--border);border-radius:6px;font-size:12px;background:var(--surface);color:var(--text);">${optHtml}</select>
                <button onclick="DriveSync.processItem('${item.id}')" style="padding:5px 12px;background:#00ac47;color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;">Procesar</button>
                <button onclick="DriveSync.dismissItem('${item.id}')" style="padding:5px 8px;background:transparent;color:var(--text-muted);border:1px solid var(--border);border-radius:6px;font-size:12px;cursor:pointer;">✕</button>
            </div>`;
        list.appendChild(row);
    }
}

function _renderProcessed() {
    const list    = document.getElementById('drive-processed-list');
    const section = document.getElementById('drive-processed-section');
    if (!list || !section) return;
    if (_state.processed.length === 0) { section.style.display = 'none'; return; }
    section.style.display = 'block';
    list.innerHTML = '';
    for (const item of [..._state.processed].reverse()) {
        const row = document.createElement('div');
        row.style.cssText = `display:flex;justify-content:space-between;align-items:center;padding:6px 10px;background:var(--surface);border:1px solid var(--border);border-radius:6px;font-size:11px;`;
        row.innerHTML = `
            <span style="color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:130px;" title="${item.name}">✅ ${item.name}</span>
            <div style="display:flex;gap:5px;align-items:center;flex-shrink:0;">
                <span style="color:var(--text-muted);">${item.picker}</span>
                <span style="background:#e8f5e9;color:#2e7d32;border-radius:4px;padding:1px 6px;">${item.bulkCount} bultos</span>
                <span style="color:var(--text-muted);">${_timeStr(item.processedAt)}</span>
            </div>`;
        list.appendChild(row);
    }
}



async function _loadBrowseFiles() {
    /* Mostrar Cargando... en el listado del panel */
    _state.browseFiles = 'loading';
    _renderExistingPanel();
    try {
        /* silent=true para no pisar el status del monitor activo */
        const dayId = await _getDayFolderId(_state.browseDate, true);
        if (!dayId) {
            _state.browseFiles = [];
            _renderExistingPanel();
            return;
        }
        const files = await _getAllTxtRecursive(dayId);
        _state.browseFiles = files;
    } catch(e) {
        _state.browseFiles = [];
        console.error('[DriveSync browse]', e);
    }
    _renderExistingPanel();
}

function _isProcessed(fileId, fileName) {
    return _state.processed.some(p => (p.id && p.id === fileId) || p.name === fileName);
}

function _getProcessedInfo(fileId, fileName) {
    return _state.processed.find(p => (p.id && p.id === fileId) || p.name === fileName);
}

function _updatePanelHeader() { _renderExistingPanel(); }

function _renderExistingPanel() {
    let panel = document.getElementById('drive-existing-panel');
    if (!panel) {
        panel = document.createElement('div');
        panel.id = 'drive-existing-panel';
        panel.style.cssText = [
            'display:none',
            'margin-top:10px',
            'border:1px solid rgba(15,23,42,0.10)',
            'border-radius:10px',
            'overflow:hidden',
            'background:rgba(255,255,255,0.96)',
        ].join(';');
        /* inyectar keyframe para el spinner si no existe */
        if (!document.getElementById('ds-spin-style')) {
            const s = document.createElement('style');
            s.id = 'ds-spin-style';
            s.textContent = '@keyframes spin{to{transform:rotate(360deg)}}';
            document.head.appendChild(s);
        }
        const syncPanel = document.getElementById('drive-sync-panel');
        if (syncPanel) syncPanel.appendChild(panel);
    }

    /* Determinar qué archivos mostrar */
    const today = new Date();
    const isToday = _getDia(_state.browseDate) === _getDia(today);
    const isTomorrow = (() => {
        const tom = new Date(today); tom.setDate(today.getDate() + 1);
        return _getDia(_state.browseDate) === _getDia(tom);
    })();
    const isLoading = _state.browseFiles === 'loading';
    const files = (isLoading || _state.browseFiles === null) ? _state.knownFiles : _state.browseFiles;

    const pickerSel = document.getElementById('picker');
    const opts = pickerSel
        ? Array.from(pickerSel.options).filter(o => o.value).map(o => ({ v: o.value, l: o.text }))
        : [{ v: DRIVE_CONFIG.DEFAULT_PICKER, l: DRIVE_CONFIG.DEFAULT_PICKER }];
    const optHtml = opts.map(o =>
        `<option value="${o.v}" ${o.v === DRIVE_CONFIG.DEFAULT_PICKER ? 'selected' : ''}>${o.l}</option>`
    ).join('');

    const processedCount = files.filter(f => _isProcessed(f.id, f.name)).length;
    const totalCount = files.length;
    const dayLabel = _getDia(_state.browseDate);
    const mesLabel = _getMes(_state.browseDate);

    const navBtn = (label, offset, active) => `
        <button onclick="DriveSync.browseDay(${offset})" style="
            padding:3px 9px;border-radius:6px;font-size:10px;font-weight:600;cursor:pointer;
            border:1px solid ${active ? '#1D9E75' : 'rgba(15,23,42,0.12)'};
            background:${active ? '#1D9E75' : 'transparent'};
            color:${active ? '#fff' : '#5C6F63'};white-space:nowrap;">${label}</button>`;

    const header = `
        <div style="display:flex;flex-direction:column;gap:0;
                    background:rgba(0,172,71,0.06);border-bottom:1px solid rgba(15,23,42,0.08);">
            <div style="display:flex;align-items:center;justify-content:space-between;padding:8px 12px;">
                <div style="display:flex;align-items:center;gap:7px;flex-wrap:wrap;">
                    <span style="font-size:11px;font-weight:700;color:#0F6E56;text-transform:uppercase;letter-spacing:.05em;">
                        📁 ${dayLabel}
                    </span>
                    <span style="font-size:10px;color:#5C6F63;">${mesLabel}</span>
                    ${totalCount > 0 ? `<span style="font-size:10px;background:rgba(0,172,71,0.12);color:#0B7A3D;
                                 padding:2px 8px;border-radius:999px;font-weight:600;">
                        ${processedCount}/${totalCount} procesados
                    </span>` : ''}
                </div>
                <button onclick="DriveSync.togglePanel()" style="
                    background:transparent;border:none;cursor:pointer;
                    color:#5C6F63;font-size:16px;line-height:1;padding:0 2px;flex-shrink:0;">×</button>
            </div>
            <div style="display:flex;gap:5px;padding:0 12px 8px;">
                ${navBtn('← Ayer', -1, false)}
                ${navBtn('Hoy', 0, isToday)}
                ${navBtn('Mañana →', 1, isTomorrow)}
            </div>
        </div>`;

    const rows = isLoading
        ? `<div style="padding:14px;text-align:center;color:#5C6F63;font-size:12px;">
               <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none"
                    stroke="#1D9E75" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"
                    style="animation:spin 1s linear infinite;vertical-align:middle;margin-right:5px">
                   <line x1="12" y1="2" x2="12" y2="6"/><line x1="12" y1="18" x2="12" y2="22"/>
                   <line x1="4.93" y1="4.93" x2="7.76" y2="7.76"/><line x1="16.24" y1="16.24" x2="19.07" y2="19.07"/>
                   <line x1="2" y1="12" x2="6" y2="12"/><line x1="18" y1="12" x2="22" y2="12"/>
                   <line x1="4.93" y1="19.07" x2="7.76" y2="16.24"/><line x1="16.24" y1="7.76" x2="19.07" y2="4.93"/>
               </svg>Cargando archivos...
           </div>`
        : files.length === 0
        ? `<div style="padding:14px;text-align:center;color:#5C6F63;font-size:12px;">Sin archivos en esta carpeta.</div>`
        : files.map(f => {
            const proc = _getProcessedInfo(f.id, f.name);
            const done = Boolean(proc);
            const timeStr = proc ? _timeStr(proc.processedAt) : '';
            return `
            <div style="display:flex;flex-direction:column;gap:6px;
                        padding:9px 12px;border-bottom:1px solid rgba(15,23,42,0.06);
                        background:${done ? 'rgba(0,172,71,0.04)' : 'transparent'};">
                <div style="display:flex;justify-content:space-between;align-items:center;gap:6px;">
                    <span style="font-size:12px;font-weight:600;
                                 color:${done ? '#0B7A3D' : '#17301F'};
                                 overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:150px;"
                          title="${f.name}">
                        ${done ? '✅' : '📄'} ${f.name}
                    </span>
                    <div style="display:flex;align-items:center;gap:5px;flex-shrink:0;">
                        ${done ? `<span style="font-size:10px;background:#e8f5e9;color:#2e7d32;
                                              border-radius:4px;padding:2px 7px;font-weight:600;">
                                    ✓ ${proc.picker} · ${timeStr}
                                  </span>` : `<span style="font-size:10px;color:#5C6F63;">${f.folder || 'Drive'}</span>`}
                    </div>
                </div>
                ${done ? '' : `
                <div style="display:flex;gap:5px;align-items:center;">
                    <select id="ep-picker-${f.id}" style="
                        flex:1;padding:4px 7px;border:1px solid rgba(15,23,42,0.12);
                        border-radius:6px;font-size:11px;background:#fff;color:#17301F;">${optHtml}</select>
                    <button onclick="DriveSync.processExisting('${f.id}')" style="
                        padding:4px 11px;background:#00ac47;color:#fff;border:none;
                        border-radius:6px;font-size:11px;font-weight:600;cursor:pointer;
                        white-space:nowrap;">Procesar</button>
                </div>`}
            </div>`;
        }).join('');

    panel.innerHTML = header + `<div style="max-height:260px;overflow-y:auto;">${rows}</div>`;
    panel.style.display = _state.panelOpen ? 'block' : 'none';
}

function _applyMode() {
    const bA = document.getElementById('drive-mode-auto');
    const bM = document.getElementById('drive-mode-manual');
    const d  = document.getElementById('drive-mode-desc');
    if (_state.mode === 'auto') {
        if (bA) { bA.style.background='#00ac47'; bA.style.color='#fff'; }
        if (bM) { bM.style.background='transparent'; bM.style.color='var(--text-muted)'; }
        if (d)  d.innerHTML = 'Asigna a <strong>kmendoza</strong> y procesa automáticamente';
    } else {
        if (bA) { bA.style.background='transparent'; bA.style.color='var(--text-muted)'; }
        if (bM) { bM.style.background='#1a73e8'; bM.style.color='#fff'; }
        if (d)  d.innerHTML = 'Elegí usuario y procesá manualmente cada archivo';
    }
    _renderPending();
}

function _loadScript(src) {
    return new Promise((res, rej) => {
        if (document.querySelector(`script[src="${src}"]`)) { res(); return; }
        const s = document.createElement('script');
        s.src = src; s.onload = res;
        s.onerror = () => rej(new Error('Error cargando ' + src));
        document.head.appendChild(s);
    });
}

async function _loadAll() {
    const timeout = (ms) => new Promise((_, rej) =>
        setTimeout(() => rej(new Error(`Timeout después de ${ms/1000}s — revisá tu conexión`)), ms)
    );

    await Promise.race([_loadScript('https://accounts.google.com/gsi/client'), timeout(10000)]);
    await Promise.race([_loadScript('https://apis.google.com/js/api.js'), timeout(10000)]);

    const apiKey = localStorage.getItem('drive_api_key');
    await Promise.race([
        new Promise((res, rej) =>
            gapi.load('client', {
                callback: () => gapi.client.init({
                    apiKey,
                    discoveryDocs: ['https://www.googleapis.com/discovery/v1/apis/drive/v3/rest']
                }).then(res).catch(rej),
                onerror: (e) => rej(new Error('gapi.load falló: ' + (e?.message || e)))
            })
        ),
        timeout(15000)
    ]);
}

async function _authorize() {
    const apiKey = _getApiKey();
    if (!apiKey) {
        _setStatus('Cancelado — sin API Key');
        _resetConnectBtn();
        return;
    }

    _setStatus('Cargando SDK...');
    try {
        await _loadAll();
    } catch(e) {
        _setStatus('❌ Error SDK: ' + e.message);
        _setToast('No se pudo cargar el SDK de Google. Revisa tu conexión.', false);
        _resetConnectBtn();
        return;
    }

    /* Verificar que la key funciona antes de pedir OAuth */
    try {
        await gapi.client.drive.files.list({ pageSize: 1, fields: 'files(id)', supportsAllDrives: true });
    } catch(e) {
        if (e.status === 400 || e.status === 403) {
            _setStatus('❌ API Key inválida');
            _setToast('La API Key no es válida o fue revocada. Bórrala y reconecta.', false);
            /* Ofrecer borrar la key para que el próximo intento la pida de nuevo */
            if (confirm('La API Key guardada no funciona.\n¿Querés ingresarla de nuevo?')) {
                localStorage.removeItem('drive_api_key');
            }
            _resetConnectBtn();
            return;
        }
        /* Error 401 = sin token aún, es normal — continuar con OAuth */
    }

    _state.tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: DRIVE_CONFIG.CLIENT_ID,
        scope:     DRIVE_CONFIG.SCOPES,
        callback:  (r) => {
            if (r.error) {
                _setStatus('❌ Error OAuth: ' + r.error);
                _setToast('Error de autorización: ' + r.error, false);
                _resetConnectBtn();
                return;
            }
            _state.isAuthorized = true;
            _startPolling();
        },
        error_callback: (e) => {
            _setStatus('❌ ' + (e.message || e.type || 'Error desconocido'));
            _setToast('Error al conectar con Google: ' + (e.message || e.type), false);
            _resetConnectBtn();
        },
    });

    _setStatus('Abriendo Google...');
    _state.tokenClient.requestAccessToken({ prompt: '' });
}

function _resetConnectBtn() {
    const btn  = document.getElementById('drive-sync-btn');
    const stop = document.getElementById('drive-sync-stop');
    if (btn)  { btn.disabled = false; btn.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg> Conectar Drive'; }
    if (stop) stop.style.display = 'none';
}

async function _searchFolder(name, parentId = null) {
    let q = `mimeType='application/vnd.google-apps.folder' and name='${name}' and trashed=false`;
    if (parentId) q += ` and '${parentId}' in parents`;
    const res = await gapi.client.drive.files.list({
        q, fields: 'files(id,name)', pageSize: 10,
        supportsAllDrives: true, includeItemsFromAllDrives: true, corpora: 'allDrives',
    });
    return res.result.files || [];
}

async function _getAllTxtRecursive(folderId, folderName = '') {
    const results = [];
    const res = await gapi.client.drive.files.list({
        q: `'${folderId}' in parents and mimeType='text/plain' and trashed=false`,
        fields: 'files(id,name)', pageSize: 100,
        supportsAllDrives: true, includeItemsFromAllDrives: true, corpora: 'allDrives',
    });
    (res.result.files || []).forEach(f => results.push({ ...f, folder: folderName }));
    const subs = await gapi.client.drive.files.list({
        q: `'${folderId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
        fields: 'files(id,name)', pageSize: 100,
        supportsAllDrives: true, includeItemsFromAllDrives: true, corpora: 'allDrives',
    });
    for (const sub of (subs.result.files || [])) {
        const nested = await _getAllTxtRecursive(sub.id, sub.name);
        results.push(...nested);
    }
    return results;
}

async function _getDayFolderId(date = new Date(), silent = false) {
    const mes = _getMes(date); const dia = _getDia(date);
    const meses = await _searchFolder(mes, DRIVE_CONFIG.ROOT_FOLDER_ID);
    const mesCarpeta = meses[0] || (await _searchFolder(mes))[0];
    if (!mesCarpeta) { if (!silent) _setStatus(`⚠ No encontré "${mes}"`); return null; }
    const dias = await _searchFolder(dia, mesCarpeta.id);
    if (!dias[0]) { if (!silent) _setStatus(`Esperando carpeta "${dia}"...`, true); return null; }
    return dias[0].id;
}

async function _downloadText(fileId) {
    const token = gapi.auth.getToken();
    if (!token?.access_token) throw new Error('Sin token');
    const r = await fetch(
        `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media&supportsAllDrives=true`,
        { headers: { Authorization: `Bearer ${token.access_token}` } }
    );
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    return r.text();
}

function _countBulks(text) {
    const zpl = (text.match(/\^XA/gi) || []).length;
    return zpl > 0 ? zpl : (text.match(/"id"\s*:\s*"?\d+"?/gi) || []).length;
}

async function _startPolling() {
    if (_state.isPolling) return;
    _state.isPolling = true;
    localStorage.setItem('drive_was_active', 'true');
    document.getElementById('drive-sync-btn').style.display = 'none';
    document.getElementById('drive-sync-stop').style.display = '';
    _setStatus('● Monitoreando en vivo', true);
    _applyMode();
    await _scan(true);
    _state.pollTimer = setInterval(() => _scan(false), DRIVE_CONFIG.POLL_INTERVAL);
}

function _stopPolling() {
    clearInterval(_state.pollTimer);
    _state.pollTimer = null; _state.isPolling = false;
    localStorage.setItem('drive_was_active', 'false');
    _setStatus('Detenido');
    const bc = document.getElementById('drive-sync-btn');
    const bs = document.getElementById('drive-sync-stop');
    if (bc) { bc.style.display = ''; bc.disabled = false; }
    if (bs) bs.style.display = 'none';
}

async function _scan(firstScan) {
    try {
        const dayId = await _getDayFolderId();
        if (!dayId) return;
        const files = await _getAllTxtRecursive(dayId);
        if (firstScan) {
            files.forEach(f => _state.knownIds.add(f.id));
            _state.knownFiles = files;
            _setFileCount(files.length);
            _setStatus(`● Monitoreando — ${files.length} existentes`, true);
            _renderExistingPanel();
            return;
        }
        const newFiles = files.filter(f => !_state.knownIds.has(f.id));
        for (const file of newFiles) {
            _state.knownIds.add(file.id);
            _state.knownFiles.push(file);   /* FIX 3: mantener knownFiles actualizado */
            await _handleNewFile(file);
        }
        if (newFiles.length > 0) {
            _setFileCount(_state.knownFiles.length);
            _renderExistingPanel();         /* refrescar panel con nuevos archivos */
        }
    } catch(err) {
        console.error('[DriveSync]', err);
        if (err.status === 401) { _state.isAuthorized = false; _stopPolling(); _setStatus('Sesión expirada'); }
    }
}

async function _handleNewFile(file) {
    let text = '';
    try { text = await _downloadText(file.id); } catch(e) { _setToast(`Error: ${file.name}`, false); return; }
    const bulkCount = _countBulks(text);

    if (Notification.permission === 'granted') {
        new Notification('Nuevo archivo en Drive', {
            body: `"${file.name}" — ${bulkCount} bultos`,
            icon: 'https://ssl.gstatic.com/images/branding/product/1x/drive_2020q4_48dp.png',
        });
    }

    if (_state.mode === 'auto') {
        await _processFile({ id: file.id, name: file.name, text, bulkCount, folder: file.folder || 'Drive', picker: DRIVE_CONFIG.DEFAULT_PICKER });
    } else {
        _state.pending.push({ id: file.id, name: file.name, folder: file.folder || 'Drive', detectedAt: new Date(), bulkCount, text });
        _setToast(`Pendiente: "${file.name}"`);
        _renderPending();
    }
}

async function _processFile({ id = null, name, text, bulkCount, folder, picker }) {
    // 1. Setear picker en el <select> del app
    const pickerInput = document.getElementById('picker');
    if (pickerInput) {
        const target = Array.from(pickerInput.options)
            .find(o => o.value.toLowerCase() === picker.toLowerCase())
            || Array.from(pickerInput.options)
                .find(o => o.value.toLowerCase().includes(picker.toLowerCase()));
        if (target) {
            pickerInput.value = target.value;
            pickerInput.dispatchEvent(new Event('change', { bubbles: true }));
        }
    }

    // 2. Pasar el archivo al App principal
    const blob = new Blob([text], { type: 'text/plain' });
    const f    = new File([blob], name, { type: 'text/plain' });

    if (window.App?.actions?.handleFiles) {
        await window.App.actions.handleFiles([f]);
    } else {
        console.warn('[DriveSync] window.App no disponible. ¿Agregaste window.App = App en app.js?');
    }

    // 3. Registrar como procesado en el panel Drive
    _state.processed.push({ id, name, folder, picker, bulkCount, processedAt: new Date() });
    _renderProcessed();
    _renderExistingPanel(); /* actualizar panel para marcar como procesado */
    _setToast(`✓ Procesado: "${name}" → ${picker}`);
    // 4. Sincronizar con historial guardado
    _highlightNewestHistorialItem();
}

// ─── API Pública ───────────────────────────────────────────────────────────
const DriveSync = {

    setMode(mode) {
        _state.mode = mode;
        localStorage.setItem('drive_mode', mode);
        _applyMode();
    },

    async processItem(id) {
        const idx = _state.pending.findIndex(p => p.id === id);
        if (idx === -1) return;
        const item   = _state.pending[idx];
        const sel    = document.getElementById(`picker-${id}`);
        const picker = sel?.value || DRIVE_CONFIG.DEFAULT_PICKER;
        const btn    = document.querySelector(`button[onclick="DriveSync.processItem('${id}')"]`);
        if (btn) { btn.disabled = true; btn.textContent = 'Procesando...'; }
        await _processFile({ id: item.id, name: item.name, text: item.text, bulkCount: item.bulkCount, folder: item.folder, picker });
        _state.pending.splice(idx, 1);
        _renderPending();
    },

    dismissItem(id) {
        _state.pending = _state.pending.filter(p => p.id !== id);
        _renderPending();
    },

    connect() {
        const btn = document.getElementById('drive-sync-btn');
        if (btn) { btn.disabled = true; btn.textContent = 'Conectando...'; }
        if (_state.isAuthorized) {
            _startPolling();
        } else {
            _authorize().catch((e) => {
                console.error('[DriveSync connect]', e);
                _setStatus('❌ Error inesperado');
                _setToast('Error inesperado al conectar. Revisá la consola.', false);
                _resetConnectBtn();
            });
        }
    },

    stop: _stopPolling,

    togglePanel() {
        _state.panelOpen = !_state.panelOpen;
        if (_state.panelOpen && _state.browseFiles === null) {
            /* primera apertura: usar knownFiles del día actual */
            _state.browseDate = new Date();
        }
        _renderExistingPanel();
    },

    browseDay(offset) {
        const base = new Date();
        base.setDate(base.getDate() + offset);
        _state.browseDate = base;
        if (offset === 0) {
            /* volver a hoy: usar los archivos ya cargados por el monitor */
            _state.browseFiles = null;
            _renderExistingPanel();
        } else {
            /* otro día: cargar desde Drive */
            _state.browseFiles = [];
            _renderExistingPanel();
            _loadBrowseFiles();
        }
    },

    async processExisting(fileId) {
        const allFiles = (_state.browseFiles !== null && _state.browseFiles !== 'loading')
            ? _state.browseFiles : _state.knownFiles;
        const file = allFiles.find(f => f.id === fileId);
        if (!file) return;
        const sel    = document.getElementById('ep-picker-' + fileId);
        const picker = sel?.value || DRIVE_CONFIG.DEFAULT_PICKER;

        /* Marcar el botón como procesando ANTES de await (la ref es válida aquí) */
        const btn = document.querySelector(`button[onclick="DriveSync.processExisting('${fileId}')"]`);
        if (btn) { btn.disabled = true; btn.textContent = 'Procesando...'; }

        try {
            const text = await _downloadText(fileId);
            const bulkCount = _countBulks(text);
            /* _processFile llama _renderExistingPanel internamente (reconstruye DOM)
               por eso la ref btn queda stale — no usar btn después de esta línea */
            await _processFile({ id: file.id, name: file.name, text, bulkCount, folder: file.folder || 'Drive', picker });
            /* El panel ya se re-renderizó con ✅ por _processFile, no hace falta más */
        } catch(e) {
            /* Botón puede estar stale si el error ocurre tarde — buscar de nuevo */
            const btnRetry = document.querySelector(`button[onclick="DriveSync.processExisting('${fileId}')"]`);
            if (btnRetry) { btnRetry.disabled = false; btnRetry.textContent = 'Procesar'; }
            _setToast('Error al procesar: ' + file.name, false);
        }
    },

    init() {
        const attach = () => {
            const btn  = document.getElementById('drive-sync-btn');
            const stop = document.getElementById('drive-sync-stop');
            if (!btn) { setTimeout(attach, 100); return; }
            btn.addEventListener('click',  () => DriveSync.connect());
            stop.addEventListener('click', () => DriveSync.stop());
            if (localStorage.getItem('drive_api_key')) {
                btn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg> Reconectar Drive`;
            }
            if (Notification.permission === 'default') Notification.requestPermission();
            _applyMode();
            if (localStorage.getItem('drive_was_active') === 'true') setTimeout(() => DriveSync.connect(), 800);
        };
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', attach);
        } else {
            attach();
        }
    }
};

DriveSync.init();
