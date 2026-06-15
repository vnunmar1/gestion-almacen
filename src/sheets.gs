// ============================================================
// SHEETS.GS — Escritura en Google Sheets, config y Telegram API
// ============================================================

// ─── Configuración ───────────────────────────────────────────

function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG');
  const rows  = sheet.getDataRange().getValues();
  const config = {};
  rows.forEach(r => {
    if (r[0] && r[0] !== 'CLAVE') config[String(r[0]).trim()] = String(r[1]).trim();
  });
  return config;
}

// ─── Escritura de movimiento ──────────────────────────────────
// El registro combinado tiene:
//   tipo, fecha, cliente, lineas[], total_palets,
//   orden_carga, matricula, matricula_remolque, transportista, observaciones
//
// Estrategia de escritura:
//   — Una fila por LÍNEA DE PRODUCTO en ENTRADAS o SALIDAS
//   — Campos comunes (orden_carga, matricula, fecha, cliente) se repiten en cada fila
// ─────────────────────────────────────────────────────────────

function writeMovement(data, chatId) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = data.tipo === 'ENTRADA' ? 'ENTRADAS' : 'SALIDAS';
  const sheet     = ss.getSheetByName(sheetName);
  const now       = new Date();

  const lineas = (data.lineas && data.lineas.length > 0)
    ? data.lineas
    : [{ referencia: '—', descripcion: '—', palets: data.total_palets || 0 }];

  lineas.forEach(linea => {
    sheet.appendRow([
      now,                             // A: Timestamp confirmación
      data.fecha         || '',        // B: Fecha albarán
      data.orden_carga   || '',        // C: Nº orden de carga
      data.cliente       || '',        // D: Cliente
      linea.referencia   || '',        // E: Referencia producto
      linea.descripcion  || '',        // F: Descripción
      linea.palets       || 0,         // G: Palets de esta línea
      data.matricula     || '',        // H: Matrícula vehículo
      data.matricula_remolque || '',   // I: Matrícula remolque
      data.transportista || '',        // J: Transportista
      data.observaciones || '',        // K: Observaciones
      String(chatId),                  // L: Operario (Telegram ID)
    ]);

    // Formato de la última fila
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, 12).setBackground(lastRow % 2 === 0 ? '#F9F9F9' : '#FFFFFF');
    sheet.getRange(lastRow, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    sheet.getRange(lastRow, 2).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(lastRow, 7).setNumberFormat('#,##0');

    // Actualizar stock para cada referencia
    addStockRowIfMissing(ss, linea.referencia, data.cliente);
  });

  Logger.log(`Movimiento ${data.tipo} registrado: orden ${data.orden_carga}, ${lineas.length} líneas, ${data.total_palets} palets totales`);
}

// ─── Consultas de stock ──────────────────────────────────────

function sendStockSummaryAfterMovement(chatId, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lineas = data.lineas || [];

  let msg = `📊 *Stock tras el movimiento:*\n`;
  lineas.forEach(l => {
    const ent   = _sumMovements(ss.getSheetByName('ENTRADAS'), l.referencia, data.cliente);
    const sal   = _sumMovements(ss.getSheetByName('SALIDAS'),  l.referencia, data.cliente);
    const stock = ent - sal;
    const emoji = stock > 10 ? '🟢' : stock > 0 ? '🟡' : '🔴';
    msg += `${emoji} *${l.referencia}* (${data.cliente}): *${stock} palets*\n`;
  });

  sendMessage(chatId, msg, { parse_mode: 'Markdown' });
}

function sendFullStockToUser(chatId) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName('STOCK');
  const data       = stockSheet.getDataRange().getValues();
  const rows       = data.slice(3).filter(r => r[0] && r[1]);

  if (!rows.length) {
    sendMessage(chatId, '📦 El almacén está vacío o no hay referencias registradas.');
    return;
  }

  let msg = '📊 *STOCK ACTUAL DEL ALMACÉN*\n\n';
  rows.forEach(r => {
    const stock = Number(r[4]);
    const emoji = stock > 10 ? '🟢' : stock > 0 ? '🟡' : '🔴';
    msg += `${emoji} *${r[0]}* — ${r[1]}: ${stock} pal.\n`;
  });

  sendMessage(chatId, msg, { parse_mode: 'Markdown' });
}

function sendFullStockToAdmin() {
  const config = getConfig();
  const admins = (config.ALLOWED_USERS || '').split(',').map(id => id.trim()).filter(Boolean);
  admins.forEach(chatId => sendFullStockToUser(chatId));
  SpreadsheetApp.getUi().alert('✅ Stock enviado por Telegram a los usuarios autorizados.');
}

function _sumMovements(sheet, referencia, cliente) {
  const rows = sheet.getDataRange().getValues();
  // Columna E (índice 4) = referencia, columna D (índice 3) = cliente, columna G (índice 6) = palets
  return rows
    .filter(r => String(r[4]) === String(referencia) && String(r[3]) === String(cliente))
    .reduce((sum, r) => sum + (Number(r[6]) || 0), 0);
}

// ─── Telegram API ────────────────────────────────────────────

function sendMessage(chatId, text, opts = {}) {
  const config  = getConfig();
  const payload = Object.assign({ chat_id: chatId, text }, opts);
  UrlFetchApp.fetch(
    `https://api.telegram.org/bot${config.TELEGRAM_BOT_TOKEN}/sendMessage`,
    { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true }
  );
}

function editMessage(chatId, messageId, text, opts = {}) {
  const config  = getConfig();
  const payload = Object.assign({ chat_id: chatId, message_id: messageId, text }, opts);
  UrlFetchApp.fetch(
    `https://api.telegram.org/bot${config.TELEGRAM_BOT_TOKEN}/editMessageText`,
    { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true }
  );
}

function answerCallbackQuery(id, text) {
  const config  = getConfig();
  const payload = { callback_query_id: id };
  if (text) payload.text = text;
  UrlFetchApp.fetch(
    `https://api.telegram.org/bot${config.TELEGRAM_BOT_TOKEN}/answerCallbackQuery`,
    { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true }
  );
}
