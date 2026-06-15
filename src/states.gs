// ============================================================
// STATES.GS — Gestión de estados y datos pendientes por usuario
// ============================================================
//
// Estados posibles por usuario:
//   IDLE                — Esperando primera foto (albarán)
//   WAITING_PORTE       — Albarán recibido, esperando carta de porte
//   PENDING_CONFIRM     — Ambas fotos procesadas, esperando confirmación
//   PENDING_EDIT_FIELD  — Operario corrigiendo un campo concreto
//
// Datos guardados en PropertiesService (JSON por chatId):
//   pendingAlbaran_{chatId}  — datos extraídos del albarán
//   pendingPorte_{chatId}    — datos extraídos de la carta de porte
//   pendingMerged_{chatId}   — objeto combinado final
//   state_{chatId}           — estado actual + campo en edición
// ============================================================

// ─── Estado de sesión ────────────────────────────────────────

function getState(chatId) {
  const val = PropertiesService.getScriptProperties().getProperty('state_' + chatId);
  return val ? JSON.parse(val) : { mode: 'IDLE' };
}

function setState(chatId, state) {
  PropertiesService.getScriptProperties()
    .setProperty('state_' + chatId, JSON.stringify(state));
}

// ─── Datos parciales por documento ───────────────────────────

function saveAlbaranData(chatId, data) {
  PropertiesService.getScriptProperties()
    .setProperty('pendingAlbaran_' + chatId, JSON.stringify(data));
}

function getAlbaranData(chatId) {
  const val = PropertiesService.getScriptProperties().getProperty('pendingAlbaran_' + chatId);
  return val ? JSON.parse(val) : null;
}

function savePorteData(chatId, data) {
  PropertiesService.getScriptProperties()
    .setProperty('pendingPorte_' + chatId, JSON.stringify(data));
}

function getPorteData(chatId) {
  const val = PropertiesService.getScriptProperties().getProperty('pendingPorte_' + chatId);
  return val ? JSON.parse(val) : null;
}

// ─── Datos combinados (para confirmar/editar) ─────────────────

function saveMergedData(chatId, data) {
  PropertiesService.getScriptProperties()
    .setProperty('pendingMerged_' + chatId, JSON.stringify(data));
}

function getMergedData(chatId) {
  const val = PropertiesService.getScriptProperties().getProperty('pendingMerged_' + chatId);
  return val ? JSON.parse(val) : null;
}

// Alias para compatibilidad con webhook.gs
function getPendingData(chatId)       { return getMergedData(chatId); }
function savePendingData(chatId, d)   { saveMergedData(chatId, d); }

function clearAllPendingData(chatId) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('pendingAlbaran_' + chatId);
  props.deleteProperty('pendingPorte_'   + chatId);
  props.deleteProperty('pendingMerged_'  + chatId);
}

// ─── Merge de los dos documentos ─────────────────────────────

function mergeDocuments(albaranData, porteData) {
  return {
    // Del albarán: tipo, líneas de producto, fecha documento
    tipo:        albaranData.tipo        || 'ENTRADA',
    fecha:       albaranData.fecha       || porteData.fecha || null,
    lineas:      albaranData.lineas      || [],   // array de {referencia, descripcion, palets}
    cliente:     albaranData.cliente     || porteData.cliente || null,
    // De la carta de porte: orden de carga, matrícula
    orden_carga: porteData.orden_carga   || null,
    matricula:   porteData.matricula     || null,
    // Totales calculados
    total_palets: (albaranData.lineas || []).reduce((s, l) => s + (l.palets || 0), 0),
    observaciones: albaranData.observaciones || porteData.observaciones || null,
  };
}

// ─── Corrección de campo individual ──────────────────────────

function applyFieldCorrection(chatId, field, value) {
  const data = getMergedData(chatId);
  if (!data) {
    sendMessage(chatId, '⚠️ No hay datos pendientes. Envía el albarán de nuevo.');
    return;
  }

  switch (field) {
    case 'tipo':
      data.tipo = value.toUpperCase().includes('SALIDA') ? 'SALIDA' : 'ENTRADA';
      break;
    case 'fecha':
      if (!/^\d{2}\/\d{2}\/\d{4}$/.test(value)) {
        sendMessage(chatId, '❌ Formato incorrecto. Usa DD/MM/YYYY (ej: 16/06/2025)');
        return;
      }
      data.fecha = value;
      break;
    case 'cliente':
      data.cliente = value.trim();
      break;
    case 'orden_carga':
      data.orden_carga = value.trim();
      break;
    case 'matricula':
      data.matricula = value.trim().toUpperCase();
      break;
    case 'total_palets':
      const n = parseInt(value);
      if (isNaN(n) || n < 1) {
        sendMessage(chatId, '❌ El total de palets debe ser un número entero positivo.');
        return;
      }
      data.total_palets = n;
      break;
    case 'observaciones':
      data.observaciones = value.trim();
      break;
    default:
      data[field] = value.trim();
  }

  saveMergedData(chatId, data);
  setState(chatId, { mode: 'PENDING_CONFIRM' });
  sendMessage(chatId,
    `✅ Campo *${field}* actualizado a: *${data[field]}*`,
    { parse_mode: 'Markdown' }
  );
  sendConfirmationMessage(chatId, data);
}
