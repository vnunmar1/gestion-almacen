// ============================================================
// SHEETS.GS — Escritura en Google Sheets, config y Telegram API
// ============================================================

// ─── Configuración ───────────────────────────────────────────

function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG');
  const rows  = sheet.getDataRange().getValues();
  const config = {};
  rows.forEach(r => { if (r[0] && r[0] !== 'CLAVE') config[String(r[0]).trim()] = String(r[1]).trim(); });
  return config;
}

// ─── Escritura de movimiento ──────────────────────────────────

function writeMovement(data, chatId) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = data.tipo === 'ENTRADA' ? 'ENTRADAS' : 'SALIDAS';
  const sheet     = ss.getSheetByName(sheetName);

  sheet.appendRow([
    new Date(),                    // A: Timestamp de confirmación
    data.fecha        || '',       // B: Fecha del albarán
    data.referencia   || '',       // C: Referencia
    data.cliente      || '',       // D: Cliente
    data.palets       || 0,        // E: Número de palets
    data.matricula    || '',       // F: Matrícula
    data.observaciones|| '',       // G: Observaciones
    String(chatId),                // H: ID Telegram del operario
  ]);

  // Formato de la nueva fila (última)
  const lastRow   = sheet.getLastRow();
  const rowRange  = sheet.getRange(lastRow, 1, 1, 8);
  const isEven    = (lastRow % 2 === 0);
  rowRange.setBackground(isEven ? '#F9F9F9' : '#FFFFFF');
  sheet.getRange(lastRow, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  sheet.getRange(lastRow, 2).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(lastRow, 5).setNumberFormat('#,##0');

  // Añadir al stock si es una referencia nueva
  addStockRowIfMissing(ss, data.referencia, data.cliente);

  Logger.log(`Movimiento ${data.tipo} registrado: ${data.referencia} / ${data.cliente} / ${data.palets} palets`);
}

// ─── Consultas de stock ──────────────────────────────────────

function sendStockForRef(chatId, referencia, cliente) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ent = _sumMovements(ss.getSheetByName('ENTRADAS'), referencia, cliente);
  const sal = _sumMovements(ss.getSheetByName('SALIDAS'),  referencia, cliente);
  const stock = ent - sal;

  const emoji = stock > 10 ? '🟢' : stock > 0 ? '🟡' : '🔴';
  sendMessage(chatId,
    `${emoji} *Stock actual de ${referencia}*\n` +
    `Cliente: ${cliente}\n` +
    `Entradas totales: ${ent} palets\n` +
    `Salidas totales:  ${sal} palets\n` +
    `*Stock actual:    ${stock} palets*`,
    { parse_mode: 'Markdown' }
  );
}

function sendFullStockToUser(chatId) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName('STOCK');
  const data     = stockSheet.getDataRange().getValues();

  // Filas de datos empiezan en índice 3 (fila 4 de la hoja)
  const rows = data.slice(3).filter(r => r[0] && r[1]);

  if (!rows.length) {
    sendMessage(chatId, '📦 El almacén está vacío.');
    return;
  }

  let msg = '📊 *STOCK ACTUAL DEL ALMACÉN*\n\n';
  rows.forEach(r => {
    const stock = Number(r[4]);
    const emoji = stock > 10 ? '🟢' : stock > 0 ? '🟡' : '🔴';
    msg += `${emoji} *${r[0]}* (${r[1]}): ${stock} palets\n`;
  });

  sendMessage(chatId, msg, { parse_mode: 'Markdown' });
}

// Para llamada desde menú de Sheets
function sendFullStockToAdmin() {
  const config = getConfig();
  const admins = (config.ALLOWED_USERS || '').split(',').map(id => id.trim()).filter(Boolean);
  admins.forEach(chatId => sendFullStockToUser(chatId));
  SpreadsheetApp.getUi().alert('✅ Stock enviado a los usuarios autorizados por Telegram.');
}

function _sumMovements(sheet, referencia, cliente) {
  const rows = sheet.getDataRange().getValues();
  return rows
    .filter(r => String(r[2]) === String(referencia) && String(r[3]) === String(cliente))
    .reduce((sum, r) => sum + (Number(r[4]) || 0), 0);
}

// ─── Telegram API helpers ────────────────────────────────────

function sendMessage(chatId, text, opts = {}) {
  const config = getConfig();
  const payload = Object.assign({ chat_id: chatId, text }, opts);
  UrlFetchApp.fetch(
    `https://api.telegram.org/bot${config.TELEGRAM_BOT_TOKEN}/sendMessage`,
    { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true }
  );
}

function editMessage(chatId, messageId, text, opts = {}) {
  const config = getConfig();
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
