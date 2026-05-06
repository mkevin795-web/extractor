/**
 * drive-sync.js
 * Monitorea Google Drive (Shared Drive) y procesa archivos .txt de etiquetas ML.
 * Modos: Automático (kmendoza) | Manual (usuario elige y procesa)
 *
 * ACTUALIZACIÓN: Ahora detecta carpetas con ambos formatos de fecha
 * - "4 MAYO" (sin padding)
 * - "04 MAYO" (con padding de ceros)
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
    // Mostrar input en el status en vez de prompt() (que puede ser bloqueado)
    return null;
}

function _restoreConnectBtn() {
    const btn = document.getElementById('drive-sync-btn');
    if (!btn) return;
    btn.disabled = false;
    const hasKey = !!localStorage.getItem('drive_api_key');
    btn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg> ${hasKey ? 'Reconectar Drive' : 'Conectar Drive'}`;
}

// Cargar procesados del día desde localStorage (se limpia automáticamente al cambiar de día)
function _todayKey() {
    const d = new Date();
    return `drive_processed_${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
}
function _loadProcessedIds() {
    try {
        const saved = localStorage.getItem(_todayKey());
        return saved ? new Set(JSON.parse(saved)) : new Set();
    } catch(e) { return new Set(); }
}
function _saveProcessedIds() {
    try {
        localStorage.setItem(_todayKey(), JSON.stringify([..._state.processedIds]));
    } catch(e) {}
}

const _state = {
    mode:         localStorage.getItem('drive_mode') || 'auto',
    isPolling:    false,
    isAuthorized: false,
    pollTimer:    null,
    knownIds:     new Set(),
    knownNames:   new Set(),   // "carpeta|nombre" — evita re-procesar mismo archivo con distinto casing
    knownFiles:   [],
    pending:      [],
    processed:    [],
    processedIds: _loadProcessedIds(), // IDs ya procesados para marcar en panel
    tokenClient:  null,
    panelOpen:    false,
    offsetDay:    0,  // 0 = hoy, -1 = ayer, +1 = mañana
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
    el.textContent = '📁 ' + n + ' archivo' + (n !== 1 ? 's' : '');
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
        const syncPanel = document.getElementById('drive-sync-panel');
        if (syncPanel) syncPanel.appendChild(panel);
    }

    const pickerSel = document.getElementById('picker');
    const opts = pickerSel
        ? Array.from(pickerSel.options).filter(o => o.value).map(o => ({ v: o.value, l: o.text }))
        : [{ v: DRIVE_CONFIG.DEFAULT_PICKER, l: DRIVE_CONFIG.DEFAULT_PICKER }];
    const optHtml = opts.map(o =>
        `<option value="${o.v}" ${o.v === DRIVE_CONFIG.DEFAULT_PICKER ? 'selected' : ''}>${o.l}</option>`
    ).join('');

    const offsetD = new Date();
    offsetD.setDate(offsetD.getDate() + _state.offsetDay);
    const diaStr  = _getDia(offsetD);
    const dayLabel = _state.offsetDay === 0  ? `HOY · ${diaStr}`
        : _state.offsetDay === -1 ? `AYER · ${diaStr}`
        : _state.offsetDay ===  1 ? `MAÑANA · ${diaStr}`
        : diaStr;

    const header = `
        <div style="display:flex;align-items:center;justify-content:space-between;
                    padding:8px 12px;background:rgba(0,172,71,0.06);
                    border-bottom:1px solid rgba(15,23,42,0.08);">
            <div style="display:flex;align-items:center;gap:6px;">
                <button onclick="DriveSync.changeDay(-1)" style="
                    background:transparent;border:1px solid rgba(15,23,42,0.12);
                    border-radius:4px;cursor:pointer;font-size:13px;
                    padding:1px 6px;color:#0F6E56;line-height:1.4;">◀</button>
                <span style="font-size:11px;font-weight:700;color:#0F6E56;text-transform:uppercase;letter-spacing:.05em;">
                    📁 ${dayLabel} — ${_state.knownFiles.length} archivo${_state.knownFiles.length !== 1 ? 's' : ''}
                </span>
                <button onclick="DriveSync.changeDay(1)" style="
                    background:transparent;border:1px solid rgba(15,23,42,0.12);
                    border-radius:4px;cursor:pointer;font-size:13px;
                    padding:1px 6px;color:#0F6E56;line-height:1.4;">▶</button>
            </div>
            <button onclick="DriveSync.togglePanel()" style="
                background:transparent;border:none;cursor:pointer;
                color:#5C6F63;font-size:16px;line-height:1;padding:0 2px;">×</button>
        </div>`;

    const rows = _state.knownFiles.length === 0
        ? `<div style="padding:14px;text-align:center;color:#5C6F63;font-size:12px;">Sin archivos en esta carpeta.</div>`
        : _state.knownFiles.map(f => {
            const isDone = _state.processedIds.has(f.id);
            return `
            <div style="display:flex;flex-direction:column;gap:6px;
                        padding:9px 12px;border-bottom:1px solid rgba(15,23,42,0.06);
                        ${isDone ? 'background:rgba(0,172,71,0.04);' : ''}">
                <div style="display:flex;justify-content:space-between;align-items:center;gap:6px;">
                    <span style="font-size:12px;font-weight:600;color:#17301F;
                                 overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:160px;"
                          title="${f.name}">📄 ${f.name}</span>
                    <span style="font-size:10px;color:#5C6F63;flex-shrink:0;">${f.folder || 'Drive'}</span>
                </div>
                <div style="display:flex;gap:5px;align-items:center;">
                    ${isDone
                        ? `<span style="flex:1;padding:4px 10px;background:#e8f5e9;color:#2e7d32;
                                        border-radius:6px;font-size:11px;font-weight:600;text-align:center;">
                                ✅ Procesado</span>`
                        : `<select id="ep-picker-${f.id}" style="
                                flex:1;padding:4px 7px;border:1px solid rgba(15,23,42,0.12);
                                border-radius:6px;font-size:11px;background:#fff;color:#17301F;">${optHtml}</select>
                            <button onclick="DriveSync.processExisting('${f.id}')" style="
                                padding:4px 11px;background:#00ac47;color:#fff;border:none;
                                border-radius:6px;font-size:11px;font-weight:600;cursor:pointer;
                                white-space:nowrap;">Procesar</button>`
                    }
                </div>
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

let _gapiReady = false;
async function _loadAll() {
    await _loadScript('https://accounts.google.com/gsi/client');
    await _loadScript('https://apis.google.com/js/api.js');
    if (_gapiReady) return; // ya inicializado, no re-inicializar
    const apiKey = localStorage.getItem('drive_api_key');
    await new Promise((res, rej) =>
        gapi.load('client', {
            callback: () => gapi.client.init({
                apiKey: apiKey || undefined,
                discoveryDocs: ['https://www.googleapis.com/discovery/v1/apis/drive/v3/rest']
            }).then(() => { _gapiReady = true; res(); }).catch(rej),
            onerror: (e) => rej(new Error('gapi.load falló: ' + (e?.message || e)))
        })
    );
}

async function _authorize() {
    _setStatus('⏳ Cargando SDK de Google...');
    try {
        await _loadAll();
    } catch(e) {
        console.error('[DriveSync] _loadAll falló:', e);
        _setStatus('❌ Error SDK: ' + e.message);
        _restoreConnectBtn();
        return;
    }

    _setStatus('⏳ Esperando autorización...');

    try {
        await new Promise((resolve, reject) => {
            // Timeout: si en 2 min no hay respuesta, abortar
            const timeout = setTimeout(() => {
                reject(new Error('Tiempo de espera agotado. ¿Se bloqueó el popup de Google?'));
            }, 120_000);

            _state.tokenClient = google.accounts.oauth2.initTokenClient({
                client_id: DRIVE_CONFIG.CLIENT_ID,
                scope:     DRIVE_CONFIG.SCOPES,
                callback: (r) => {
                    clearTimeout(timeout);
                    if (r.error) {
                        reject(new Error('OAuth: ' + r.error));
                    } else {
                        _state.isAuthorized = true;
                        resolve();
                    }
                },
                error_callback: (e) => {
                    clearTimeout(timeout);
                    reject(new Error(e.message || e.type || 'error desconocido'));
                },
            });

            _state.tokenClient.requestAccessToken({ prompt: 'consent' });
        });

        // Éxito
        _startPolling();

    } catch(e) {
        console.error('[DriveSync] OAuth falló:', e);
        _setStatus('❌ ' + e.message + ' — Intentá de nuevo');
        _restoreConnectBtn();
    }
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

async function _getAllTxtRecursive(folderId, folderName = '', _seenIds = new Set()) {
    const results = [];
    const res = await gapi.client.drive.files.list({
        q: `'${folderId}' in parents and mimeType='text/plain' and trashed=false`,
        fields: 'files(id,name,createdTime)', pageSize: 100,
        supportsAllDrives: true, includeItemsFromAllDrives: true, corpora: 'allDrives',
    });
    for (const f of (res.result.files || [])) {
        if (_seenIds.has(f.id)) continue;
        _seenIds.add(f.id);
        results.push({ ...f, folder: folderName });
    }
    const subs = await gapi.client.drive.files.list({
        q: `'${folderId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
        fields: 'files(id,name)', pageSize: 100,
        supportsAllDrives: true, includeItemsFromAllDrives: true, corpora: 'allDrives',
    });
    for (const sub of (subs.result.files || [])) {
        const nested = await _getAllTxtRecursive(sub.id, sub.name, _seenIds);
        results.push(...nested);
    }
    return results;
}

function _nameKey(f) {
    return `${(f.folder || '').toLowerCase().trim()}|${f.name.toLowerCase().trim()}`;
}

/**
 * FUNCIÓN ACTUALIZADA: Detecta ambos formatos de carpeta por día
 * - "4 MAYO" (sin padding)
 * - "04 MAYO" (con padding de ceros)
 */
async function _getDayFolderId(offset = 0) {
    const d = new Date();
    d.setDate(d.getDate() + offset);
    const mes = `${MESES[d.getMonth()]} ${d.getFullYear()}`; // "MAYO 2026"

    // Generar ambos formatos de día
    const diaPadded = `${String(d.getDate()).padStart(2, '0')} ${MESES[d.getMonth()]}`; // "04 MAYO"
    const diaNormal = `${d.getDate()} ${MESES[d.getMonth()]}`; // "4 MAYO"

    // Buscar carpeta de mes
    const meses = await _searchFolder(mes, DRIVE_CONFIG.ROOT_FOLDER_ID);
    const mesCarpeta = meses[0] || (await _searchFolder(mes))[0];
    if (!mesCarpeta) {
        _setStatus(`⚠ Sin carpeta "${mes}"`);
        return null;
    }

    // Intentar primero con padding (04 MAYO)
    let dias = await _searchFolder(diaPadded, mesCarpeta.id);

    // Si no encuentra, intentar sin padding (4 MAYO)
    if (!dias[0]) {
        dias = await _searchFolder(diaNormal, mesCarpeta.id);
    }

    if (!dias[0]) {
        _setStatus(`Sin carpeta "${diaPadded}" o "${diaNormal}" en ${mes}`, true);
        return null;
    }

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
    _state.offsetDay = 0;
    localStorage.setItem('drive_was_active', 'true');

    const bc = document.getElementById('drive-sync-btn');
    const bs = document.getElementById('drive-sync-stop');
    if (bc) bc.style.display = 'none';
    if (bs) bs.style.display = '';

    // Mostrar fila "Carpeta: 29 ABRIL  [Ver archivos]"
    const dayStatus = document.getElementById('drive-day-status');
    const dayName   = document.getElementById('drive-day-name');
    if (dayName)   dayName.textContent = 'Carpeta: ' + _getDia();
    if (dayStatus) dayStatus.style.display = 'flex';

    _setStatus('● Monitoreando en vivo', true);
    _applyMode();
    await _scan(true);
    _state.pollTimer = setInterval(() => _scan(false), DRIVE_CONFIG.POLL_INTERVAL);
}

function _stopPolling() {
    clearInterval(_state.pollTimer);
    _state.pollTimer = null;
    _state.isPolling = false;
    localStorage.setItem('drive_was_active', 'false');
    _setStatus('Desconectado');

    const bc = document.getElementById('drive-sync-btn');
    const bs = document.getElementById('drive-sync-stop');
    if (bc) { bc.style.display = ''; bc.disabled = false; }
    if (bs) bs.style.display = 'none';

    const dayStatus = document.getElementById('drive-day-status');
    if (dayStatus) dayStatus.style.display = 'none';
    _setFileCount(null);
    _restoreConnectBtn();
}

async function _scan(firstScan) {
    try {
        const dayId = await _getDayFolderId(0);
        if (!dayId) return;
        const files = await _getAllTxtRecursive(dayId);

        if (firstScan) {
            files.forEach(f => {
                _state.knownIds.add(f.id);
                _state.knownNames.add(_nameKey(f));
            });
            _state.knownFiles = files;
            _setFileCount(files.length);
            _setStatus(`● Monitoreando — ${files.length} existentes`, true);
            _renderExistingPanel();
            return;
        }

        const newFiles = files.filter(f =>
            !_state.knownIds.has(f.id) &&
            !_state.knownNames.has(_nameKey(f))
        );
        for (const file of newFiles) {
            _state.knownIds.add(file.id);
            _state.knownNames.add(_nameKey(file));
            await _handleNewFile(file);
        }
        if (newFiles.length > 0) {
            _state.knownFiles = files;
            _setFileCount(files.length);
        }

    } catch(err) {
        console.error('[DriveSync] _scan error:', err);
        const code = err.status ?? err.code ?? err.result?.error?.code;
        if (code === 401 || code === 403) {
            _state.isAuthorized = false;
            _gapiReady = false;
            _stopPolling();
            _setStatus('⚠️ Sesión expirada — clic en Reconectar');
        } else {
            _setStatus('⚠️ Error temporal, reintentando en 30s...');
        }
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
        await _processFile({ name: file.name, text, bulkCount, folder: file.folder || 'Drive', picker: DRIVE_CONFIG.DEFAULT_PICKER });
    } else {
        _state.pending.push({ id: file.id, name: file.name, folder: file.folder || 'Drive', detectedAt: new Date(), bulkCount, text });
        _setToast(`Pendiente: "${file.name}"`);
        _renderPending();
    }
}

async function _processFile({ name, text, bulkCount, folder, picker }) {
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

    const blob = new Blob([text], { type: 'text/plain' });
    const f    = new File([blob], name, { type: 'text/plain' });
    let realCount = bulkCount;

    if (window.App?.actions?.handleFiles) {
        const prevLen = window.App.state?.storedDocuments?.length ?? -1;
        await window.App.actions.handleFiles([f]);
        const docs = window.App.state?.storedDocuments;
        if (Array.isArray(docs) && docs.length > prevLen && docs[0]?.rows?.length > 0) {
            realCount = docs[0].rows.length;
        }
    } else {
        console.warn('[DriveSync] window.App no disponible.');
    }

    _state.processed.push({ name, folder, picker, bulkCount: realCount, processedAt: new Date() });
    _renderProcessed();
    _setToast(`✓ "${name}" — ${realCount} pedidos → ${picker}`);
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
        await _processFile({ name: item.name, text: item.text, bulkCount: item.bulkCount, folder: item.folder, picker });
        _state.pending.splice(idx, 1);
        _renderPending();
    },

    dismissItem(id) {
        _state.pending = _state.pending.filter(p => p.id !== id);
        _renderPending();
    },

    connect() {
        if (_state.isPolling) return; // ya conectado
        const btn = document.getElementById('drive-sync-btn');
        if (btn) { btn.disabled = true; btn.textContent = '⏳ Conectando...'; }

        const apiKey = localStorage.getItem('drive_api_key');
        if (!apiKey) {
            // Pedir API Key de forma visible (no prompt bloqueable)
            const key = window.prompt('Ingresá tu Google API Key para Drive:');
            if (!key?.trim()) {
                _setStatus('Cancelado — se requiere API Key');
                _restoreConnectBtn();
                return;
            }
            localStorage.setItem('drive_api_key', key.trim());
            _gapiReady = false; // forzar re-init con la nueva key
        }

        _authorize().catch((e) => {
            console.error('[DriveSync] connect error:', e);
            _setStatus('❌ Error: ' + (e.message || e));
            _restoreConnectBtn();
        });
    },

    stop: _stopPolling,

    togglePanel() {
        _state.panelOpen = !_state.panelOpen;
        if (_state.panelOpen) _state.offsetDay = 0;
        _renderExistingPanel();
    },

    async changeDay(delta) {
        _state.offsetDay += delta;
        _state.knownFiles = [];
        _setFileCount(null);
        const panel = document.getElementById('drive-existing-panel');
        if (panel) panel.innerHTML = '<div style="padding:14px;text-align:center;font-size:12px;color:#5C6F63;">Cargando...</div>';
        try {
            const dayId = await _getDayFolderId(_state.offsetDay);
            if (dayId) {
                const files = await _getAllTxtRecursive(dayId);
                _state.knownFiles = files;
                _setFileCount(files.length);
            } else {
                _setFileCount(0);
            }
        } catch(e) { console.error('[DriveSync] changeDay:', e); }
        _renderExistingPanel();
    },

    async processExisting(fileId) {
        const file = _state.knownFiles.find(f => f.id === fileId);
        if (!file) return;
        const sel    = document.getElementById('ep-picker-' + fileId);
        const picker = sel?.value || DRIVE_CONFIG.DEFAULT_PICKER;
        const btn    = document.querySelector(`button[onclick="DriveSync.processExisting('${fileId}')"]`);
        if (btn) { btn.disabled = true; btn.textContent = 'Procesando...'; }
        try {
            const text = await _downloadText(fileId);
            const bulkCount = _countBulks(text);
            await _processFile({ name: file.name, text, bulkCount, folder: file.folder || 'Drive', picker });
            _state.processedIds.add(fileId);
            _saveProcessedIds();
            _renderExistingPanel();
            if (btn) { btn.textContent = '✓ Listo'; btn.style.background = '#e8f5e9'; btn.style.color = '#2e7d32'; }
        } catch(e) {
            if (btn) { btn.disabled = false; btn.textContent = 'Procesar'; }
            _setToast('Error al procesar: ' + file.name, false);
        }
    },

    init() {
        const attach = () => {
            const btn = document.getElementById('drive-sync-btn');
            if (!btn) { setTimeout(attach, 100); return; }

            // Restaurar texto correcto del botón
            _restoreConnectBtn();
            if (Notification.permission === 'default') Notification.requestPermission();
            _applyMode();

            // Auto-reconectar si estaba activo (solo si hay API key guardada)
            const wasActive = localStorage.getItem('drive_was_active') === 'true';
            const hasKey    = !!localStorage.getItem('drive_api_key');
            if (wasActive && hasKey) {
                setTimeout(() => DriveSync.connect(), 1000);
            }
        };
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', attach);
        } else {
            attach();
        }
    }
};

DriveSync.init();
