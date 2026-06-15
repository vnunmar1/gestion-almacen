// ============================================================
// STATES.GS — Gestión de estados y datos pendientes por usuario
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

// ─── Datos pendientes de confirmación ────────────────────────

function savePendingData(chatId, data) {
  PropertiesService.getScriptProperties()
    .setProperty('pending_' + chatId, JSON.stringify(data));
}

function getPendingData(chatId) {
  const val = PropertiesService.getScriptProperties().getProperty('pending_' + chatId);
  return val ? JSON.parse(val) : null;
}

function clearPendingData(chatId) {
  PropertiesService.getScriptProperties().deleteProperty('pending_' + chatId);
}

// ─── Corrección de campo individual ──────────────────────────

function applyFieldCorrection(chatId, field, value) {
  const data = getPendingData(chatId);
  if (!data) {
    sendMessage(chatId, '⚠️ No hay datos pendientes. Envía la foto de nuevo.');
    return;
  }

  // Normalizar según el campo
  switch (field) {
    case 'palets':
      data[field] = parseInt(value);
      if (isNaN(data[field])) {
        sendMessage(chatId, '❌ El número de palets debe ser un número entero. Inténtalo de nuevo.');
        return;
      }
      break;
    case 'tipo':
      data[field] = value.toUpperCase().includes('SALIDA') ? 'SALIDA' : 'ENTRADA';
      break;
    case 'fecha':
      // Validar formato DD/MM/YYYY
      if (!/^\d{2}\/\d{2}\/\d{4}$/.test(value)) {
        sendMessage(chatId, '❌ Formato de fecha incorrecto. Usa DD/MM/YYYY (ej: 15/06/2025)');
        return;
      }
      data[field] = value;
      break;
    default:
      data[field] = value.trim();
  }

  savePendingData(chatId, data);
  setState(chatId, { mode: 'PENDING_CONFIRM' });
  sendMessage(chatId, `✅ Campo *${field}* actualizado a: *${data[field]}*`, { parse_mode: 'Markdown' });
  sendConfirmationMessage(chatId, data);
}
