/*
════════════════════════════════════════════════════════════════════════════════
app-ACTUALIZADO.js
Cambios principales: Nuevas funciones para historial mejorado con pestañas
════════════════════════════════════════════════════════════════════════════════
*/

// [Tu código existente de app.js se mantiene igual...]
// [Agrega las siguientes funciones al final del archivo]

/* ════════════════════════════════════════════════════════════════════════════════ */
/* HISTORIAL MEJORADO - NUEVAS FUNCIONES */
/* ════════════════════════════════════════════════════════════════════════════════ */

/**
 * Cambia entre pestañas del historial
 * @param {number} tabIndex - Índice de la pestaña (0 = Historial, 1 = Procesados)
 * @param {HTMLElement} btnElement - Elemento del botón clickeado
 */
function switchTab(tabIndex, btnElement) {
  // Desactivar todas las pestañas y contenido
  document.querySelectorAll('.history-tab').forEach(tab => {
    tab.classList.remove('active');
  });
  document.querySelectorAll('.tab-content').forEach(content => {
    content.classList.remove('active');
  });

  // Activar pestaña actual
  btnElement.classList.add('active');
  const tabs = document.querySelectorAll('.tab-content');
  if (tabs[tabIndex]) {
    tabs[tabIndex].classList.add('active');

    // Re-renderizar contenido según pestaña
    if (tabIndex === 0) {
      renderHistorialTab();
    } else {
      renderProcesadosTab();
    }
  }
}

/**
 * Renderiza la pestaña de Historial con documentos guardados
 */
function renderHistorialTab() {
  const grid = document.getElementById('historialGrid');
  const count = document.getElementById('historialCount');

  if (!grid || !count) return;

  // Obtener documentos del estado de la app
  const docs = App.state.storedDocuments || [];
  count.textContent = docs.length + ' total';

  // Generar HTML de tarjetas
  const html = docs.map((doc, idx) => {
    const typeSummary = App.helpers.summarizeDocumentType(doc.rows);
    const className = `history-card ${App.state.selectedHistoryDocumentIds && App.state.selectedHistoryDocumentIds.has(doc.id) ? 'selected' : ''}`;

    return `
      <div class="${className}" data-doc-id="${doc.id}" onclick="toggleDocumentSelection(this, '${doc.id}')">
        <div class="card-header">
          <span class="card-title">${App.helpers.escapeHtml(doc.sourceName || 'Documento sin nombre')}</span>
          <span class="card-time">${App.storage.formatClockTime(doc.storedAt)}</span>
        </div>
        <div class="card-badges">
          <span class="card-badge ${typeSummary.key}">${typeSummary.label}</span>
        </div>
        <div class="card-info">
          <div class="info-row">
            <span>${doc.rows.length} número(s)</span>
            <strong>${App.helpers.escapeHtml(doc.picker || 'sin picker')}</strong>
          </div>
        </div>
      </div>
    `;
  }).join('');

  grid.innerHTML = html || '<p style="padding: 2rem; text-align: center; color: var(--text-muted);">No hay documentos guardados</p>';
}

/**
 * Renderiza la pestaña de Procesados Hoy (desde Google Drive)
 */
function renderProcesadosTab() {
  const grid = document.getElementById('procesadosGrid');
  const count = document.getElementById('procesadosCount');

  if (!grid || !count) return;

  // Obtener documentos procesados del estado
  const processed = App.state.processed || [];
  count.textContent = processed.length + ' archivo' + (processed.length !== 1 ? 's' : '');

  // Generar HTML de tarjetas
  const html = processed.map((item, idx) => {
    return `
      <div class="history-card" data-processed-id="${idx}">
        <div class="card-header">
          <span class="card-title">${App.helpers.escapeHtml(item.name || 'Archivo sin nombre')}</span>
          <span class="card-time">${App.storage.formatClockTime(item.processedAt)}</span>
        </div>
        <div class="card-badges">
          <span class="card-badge flex">PROCESADO</span>
        </div>
        <div class="card-info">
          <div class="info-row">
            <span>${item.bulkCount || 0} número(s)</span>
            <strong>${App.helpers.escapeHtml(item.picker || 'Automático')}</strong>
          </div>
        </div>
      </div>
    `;
  }).join('');

  grid.innerHTML = html || '<p style="padding: 2rem; text-align: center; color: var(--text-muted);">No hay procesados aún</p>';
}

/**
 * Alterna la selección de un documento en el historial
 * @param {HTMLElement} element - Elemento de la tarjeta
 * @param {string} docId - ID del documento
 */
function toggleDocumentSelection(element, docId) {
  element.classList.toggle('selected');

  // Inicializar Set si no existe
  if (!App.state.selectedHistoryDocumentIds) {
    App.state.selectedHistoryDocumentIds = new Set();
  }

  if (App.state.selectedHistoryDocumentIds.has(docId)) {
    App.state.selectedHistoryDocumentIds.delete(docId);
  } else {
    App.state.selectedHistoryDocumentIds.add(docId);
  }
}

/**
 * Inicializa la interfaz del historial mejorado
 * Llama a esta función en App.actions.init()
 */
function initHistoryUI() {
  // Inicializar Set de documentos seleccionados si no existe
  if (!App.state.selectedHistoryDocumentIds) {
    App.state.selectedHistoryDocumentIds = new Set();
  }

  // Renderizar ambas pestañas
  renderHistorialTab();
  renderProcesadosTab();

  // Configurar primero la pestaña por defecto
  const firstTab = document.querySelector('.history-tab');
  if (firstTab) {
    firstTab.classList.add('active');
  }
  const firstContent = document.querySelector('.tab-content');
  if (firstContent) {
    firstContent.classList.add('active');
  }
}

/**
 * Función de paginación para puntos navegables
 * (Opcional - para agregar si quieres paginación funcional)
 */
function setupPagination() {
  const dots = document.querySelectorAll('.pagination-dot');

  dots.forEach((dot, idx) => {
    dot.addEventListener('click', () => {
      // Desactivar todos los puntos
      dots.forEach(d => d.classList.remove('active'));
      dot.classList.add('active');

      // Aquí puedes agregar lógica de cambio de página si lo necesitas
      // Por ahora es visual solamente
    });
  });
}

/* ════════════════════════════════════════════════════════════════════════════════ */
/* FIN DE NUEVAS FUNCIONES */
/* ════════════════════════════════════════════════════════════════════════════════ */

/*
INSTRUCCIONES DE INTEGRACIÓN:

1. Copia todas las funciones arriba y agrégalas al final de tu app.js original

2. En la función App.actions.init() (o tu función de inicialización), agrega:

   initHistoryUI();  // Inicializa el historial mejorado

3. Asegúrate que tu App.state tenga:
   - App.state.storedDocuments (array de documentos)
   - App.state.processed (array de procesados desde Drive)

4. Las funciones usan estas funciones auxiliares que DEBEN existir:
   - App.helpers.summarizeDocumentType(rows)
   - App.helpers.escapeHtml(text)
   - App.storage.formatClockTime(timestamp)

5. Si alguna función no existe, crea un stub:

   App.helpers.summarizeDocumentType = function(rows) {
     return { key: 'flex', label: 'FLEX' };
   };
*/
