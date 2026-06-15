// ============================================================
// WEBHOOK.GS — Punto de entrada de todos los eventos Telegram
// ============================================================

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.message) {
      const msg = body.message;
      if (!isAuthorized(msg.chat.id)) {
        sendMessage(msg.chat.id, '⛔ No tienes autorización para usar este sistema.');
        return _ok();
      }
      if (msg.photo)    handlePhoto(msg);
      else if (msg.text) handleText(msg);
    }

    if (body.callback_query) handleCallback(body.callback_query);

  } catch (err) {
    Logger.log('Error en doPost: ' + err + '\n' + err.stack);
  }
  return _ok();
}

function isAuthorized(chatId) {
  const config = getConfig();
  if (!config.ALLOWED_USERS) return false;
  const allowed = config.ALLOWED_USERS.split(',').map(id => id.trim());
  return allowed.includes(String(chatId));
}

function handleText(msg) {
  const chatId = msg.chat.id;
  const text   = (msg.text || '').trim();
  const state  = getState(chatId);

  const commands = {
    '/start': () => {
      setState(chatId, { mode: 'IDLE' });
      const config = getConfig();
      const nombre = config.ALMACEN_NOMBRE || 'Almacén';
      sendMessage(chatId,
        `👋 Bienvenido al sistema de control de *${nombre}*.\n\n` +
        `📷 Envía una foto del albarán o carta de porte para registrar una *entrada* o *salida*.\n\n` +
        `Comandos disponibles:\n` +
        `/stock — Ver stock actual\n` +
        `/ayuda — Ver esta ayuda`,
        { parse_mode: 'Markdown' }
      );
    },
    '/stock': () => sendFullStockToUser(chatId),
    '/ayuda': () => {
      sendMessage(chatId,
        '📖 *Ayuda del sistema de almacén*\n\n' +
        '1️⃣ Haz una foto clara al albarán o carta de porte\n' +
        '2️⃣ Envíala a este chat\n' +
        '3️⃣ Revisa los datos extraídos\n' +
        '4️⃣ Confirma o corrige si hay errores\n\n' +
        '⚠️ Asegúrate de que la imagen tenga buena iluminación y los datos sean legibles.',
        { parse_mode: 'Markdown' }
      );
    },
  };

  if (commands[text]) {
    commands[text]();
    return;
  }

  // Estado: esperando valor corregido para un campo
  if (state && state.mode === 'PENDING_EDIT_FIELD') {
    applyFieldCorrection(chatId, state.field, text);
    return;
  }

  sendMessage(chatId, '❓ No entiendo ese comando. Envía /ayuda para ver las opciones disponibles.');
}

function handlePhoto(msg) {
  const chatId = msg.chat.id;
  const processingMsg = sendMessage(chatId, '⏳ Procesando imagen, un momento...');

  try {
    const fileId = msg.photo[msg.photo.length - 1].file_id;
    const data   = extractDataFromPhoto(fileId);

    if (!data || !data.referencia) {
      sendMessage(chatId,
        '❌ No pude extraer los datos del documento.\n\n' +
        '💡 *Consejos:*\n' +
        '• Asegúrate de que el albarán esté bien iluminado\n' +
        '• Acércate más para que el texto sea legible\n' +
        '• Evita sombras y reflejos',
        { parse_mode: 'Markdown' }
      );
      return;
    }

    savePendingData(chatId, data);
    setState(chatId, { mode: 'PENDING_CONFIRM' });
    sendConfirmationMessage(chatId, data);

  } catch (err) {
    Logger.log('Error en handlePhoto: ' + err);
    sendMessage(chatId, '❌ Error al procesar la imagen. Inténtalo de nuevo.');
  }
}

function handleCallback(cbq) {
  const chatId   = cbq.message.chat.id;
  const msgId    = cbq.message.message_id;
  const data     = cbq.data;
  answerCallbackQuery(cbq.id);

  if (data === 'CONFIRM') {
    const pending = getPendingData(chatId);
    if (!pending) { sendMessage(chatId, '⚠️ Los datos han expirado. Envía la foto de nuevo.'); return; }
    writeMovement(pending, chatId);
    clearPendingData(chatId);
    setState(chatId, { mode: 'IDLE' });
    editMessage(chatId, msgId, '✅ *Registro guardado correctamente.*', {});
    sendStockForRef(chatId, pending.referencia, pending.cliente);
    notifyGroup(pending, chatId);
    return;
  }

  if (data === 'CANCEL') {
    clearPendingData(chatId);
    setState(chatId, { mode: 'IDLE' });
    editMessage(chatId, msgId, '❌ Registro cancelado.', {});
    return;
  }

  if (data === 'EDIT_MENU') {
    sendEditMenu(chatId);
    return;
  }

  if (data.startsWith('EDIT_')) {
    const field = data.replace('EDIT_', '').toLowerCase();
    setState(chatId, { mode: 'PENDING_EDIT_FIELD', field });
    const fieldLabels = {
      tipo: 'tipo (ENTRADA o SALIDA)',
      referencia: 'referencia del producto',
      cliente: 'nombre del cliente',
      palets: 'número de palets',
      fecha: 'fecha (DD/MM/YYYY)',
      matricula: 'matrícula del vehículo',
      observaciones: 'observaciones',
    };
    sendMessage(chatId,
      `✏️ Escribe el valor correcto para *${fieldLabels[field] || field}*:`,
      { parse_mode: 'Markdown' }
    );
  }
}

function sendEditMenu(chatId) {
  const keyboard = { inline_keyboard: [
    [
      { text: '🔄 Tipo',        callback_data: 'EDIT_tipo' },
      { text: '🏷️ Referencia',  callback_data: 'EDIT_referencia' },
    ],
    [
      { text: '👤 Cliente',     callback_data: 'EDIT_cliente' },
      { text: '📦 Palets',      callback_data: 'EDIT_palets' },
    ],
    [
      { text: '📅 Fecha',       callback_data: 'EDIT_fecha' },
      { text: '🚛 Matrícula',   callback_data: 'EDIT_matricula' },
    ],
    [
      { text: '📝 Observaciones', callback_data: 'EDIT_observaciones' },
    ],
    [
      { text: '« Volver a confirmar', callback_data: 'SHOW_CONFIRM' },
    ],
  ]};
  sendMessage(chatId, '¿Qué campo quieres corregir?', { reply_markup: JSON.stringify(keyboard) });
}

function sendConfirmationMessage(chatId, data) {
  const tipo_emoji = data.tipo === 'ENTRADA' ? '📥' : '📤';
  const text =
    `${tipo_emoji} *Datos extraídos del documento:*\n\n` +
    `🔄 Tipo:        *${data.tipo || '?'}*\n` +
    `🏷️ Referencia: *${data.referencia || '?'}*\n` +
    `👤 Cliente:    *${data.cliente || '?'}*\n` +
    `📦 Palets:     *${data.palets || '?'}*\n` +
    `📅 Fecha:      *${data.fecha || '?'}*\n` +
    `🚛 Matrícula:  *${data.matricula || 'No detectada'}*\n` +
    `📝 Notas:      *${data.observaciones || '—'}*\n\n` +
    `¿Son correctos estos datos?`;

  const keyboard = { inline_keyboard: [[
    { text: '✅ Confirmar y guardar', callback_data: 'CONFIRM' },
    { text: '✏️ Corregir',           callback_data: 'EDIT_MENU' },
    { text: '❌ Cancelar',           callback_data: 'CANCEL' },
  ]]};

  sendMessage(chatId, text, { parse_mode: 'Markdown', reply_markup: JSON.stringify(keyboard) });
}

function notifyGroup(data, operarioChatId) {
  const config = getConfig();
  if (!config.NOTIF_GROUP_ID) return;
  const tipo_emoji = data.tipo === 'ENTRADA' ? '📥' : '📤';
  const msg =
    `${tipo_emoji} *Nuevo movimiento registrado*\n` +
    `Ref: ${data.referencia} | Cliente: ${data.cliente}\n` +
    `Palets: ${data.palets} | Fecha: ${data.fecha}\n` +
    `Operario: ${operarioChatId}`;
  sendMessage(config.NOTIF_GROUP_ID, msg, { parse_mode: 'Markdown' });
}

// ─── Registrar webhook (ejecutar desde menú Sheets) ──────────
function registerWebhook() {
  const config  = getConfig();
  const token   = config.TELEGRAM_BOT_TOKEN;
  const url     = config.WEBHOOK_URL;

  if (!token || !url) {
    SpreadsheetApp.getUi().alert('❌ Faltan TELEGRAM_BOT_TOKEN o WEBHOOK_URL en CONFIG.');
    return;
  }

  const res = UrlFetchApp.fetch(
    `https://api.telegram.org/bot${token}/setWebhook?url=${encodeURIComponent(url)}`
  );
  const json = JSON.parse(res.getContentText());
  SpreadsheetApp.getUi().alert(
    json.ok ? '✅ Webhook registrado' : '❌ Error al registrar webhook',
    JSON.stringify(json, null, 2),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function _ok() {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
