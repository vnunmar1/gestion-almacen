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
      if (msg.photo)     handlePhoto(msg);
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
  return config.ALLOWED_USERS.split(',').map(id => id.trim()).includes(String(chatId));
}

// ─── Texto / Comandos ─────────────────────────────────────────

function handleText(msg) {
  const chatId = msg.chat.id;
  const text   = (msg.text || '').trim();
  const state  = getState(chatId);

  // Comandos
  if (text === '/start') {
    setState(chatId, { mode: 'IDLE' });
    clearAllPendingData(chatId);
    const config = getConfig();
    sendMessage(chatId,
      `👋 Bienvenido al sistema de control de *${config.ALMACEN_NOMBRE || 'Almacén'}*.\n\n` +
      `📋 *Proceso de registro en 2 pasos:*\n` +
      `1️⃣ Envía foto del *albarán* (referencias, palets)\n` +
      `2️⃣ Envía foto de la *carta de porte* (matrícula, orden de carga)\n\n` +
      `Comandos:\n` +
      `/stock — Ver stock actual\n` +
      `/cancelar — Cancelar registro en curso\n` +
      `/ayuda — Esta ayuda`,
      { parse_mode: 'Markdown' }
    );
    return;
  }

  if (text === '/cancelar') {
    clearAllPendingData(chatId);
    setState(chatId, { mode: 'IDLE' });
    sendMessage(chatId, '❌ Registro cancelado. Envía el albarán cuando quieras empezar de nuevo.');
    return;
  }

  if (text === '/stock') {
    sendFullStockToUser(chatId);
    return;
  }

  if (text === '/ayuda') {
    sendHelp(chatId);
    return;
  }

  // Corrección de campo
  if (state && state.mode === 'PENDING_EDIT_FIELD') {
    applyFieldCorrection(chatId, state.field, text);
    return;
  }

  sendMessage(chatId, '❓ Envía una foto del albarán para comenzar, o /ayuda para ver los comandos.');
}

// ─── Fotos ────────────────────────────────────────────────────

function handlePhoto(msg) {
  const chatId = msg.chat.id;
  const state  = getState(chatId);
  const mode   = state ? state.mode : 'IDLE';

  // Foto más grande (mejor resolución)
  const fileId = msg.photo[msg.photo.length - 1].file_id;

  // Paso 1: albarán
  if (mode === 'IDLE' || mode === 'PENDING_EDIT_FIELD') {
    sendMessage(chatId, '⏳ Procesando el *albarán*...', { parse_mode: 'Markdown' });
    const data = extractDataFromPhoto(fileId, 'albaran');

    if (!data || !Array.isArray(data.lineas) || data.lineas.length === 0) {
      sendMessage(chatId,
        '❌ No pude leer el albarán correctamente.\n\n' +
        '💡 *Consejos:*\n• Buena iluminación, sin sombras\n• El documento completo en el encuadre\n• Acércate para que el texto sea legible',
        { parse_mode: 'Markdown' }
      );
      return;
    }

    saveAlbaranData(chatId, data);
    setState(chatId, { mode: 'WAITING_PORTE' });

    // Confirmar lo leído del albarán
    const lineasTxt = data.lineas.map((l, i) =>
      `  ${i+1}. Ref *${l.referencia || '?'}* — ${l.descripcion || ''} — *${l.palets} palets*`
    ).join('\n');

    sendMessage(chatId,
      `✅ *Albarán leído* (${data.tipo || 'ENTRADA'})\n` +
      `📅 Fecha: ${data.fecha || '?'}\n` +
      `👤 Cliente: ${data.cliente || '?'}\n\n` +
      `📦 *Líneas de producto:*\n${lineasTxt}\n\n` +
      `📄 Ahora envía la foto de la *carta de porte*.`,
      { parse_mode: 'Markdown' }
    );
    return;
  }

  // Paso 2: carta de porte
  if (mode === 'WAITING_PORTE') {
    sendMessage(chatId, '⏳ Procesando la *carta de porte*...', { parse_mode: 'Markdown' });
    const porteData   = extractDataFromPhoto(fileId, 'porte');
    const albaranData = getAlbaranData(chatId);

    if (!porteData || !porteData.orden_carga) {
      sendMessage(chatId,
        '❌ No pude leer la carta de porte.\n\n' +
        '💡 Intenta de nuevo con mejor iluminación, o usa /cancelar para empezar de nuevo.',
        { parse_mode: 'Markdown' }
      );
      return;
    }

    savePorteData(chatId, porteData);
    const merged = mergeDocuments(albaranData, porteData);
    saveMergedData(chatId, merged);
    setState(chatId, { mode: 'PENDING_CONFIRM' });
    sendConfirmationMessage(chatId, merged);
    return;
  }

  // Si ya hay datos pendientes de confirmar
  if (mode === 'PENDING_CONFIRM') {
    sendMessage(chatId,
      '⚠️ Hay un registro pendiente de confirmar.\n' +
      'Usa los botones anteriores para *confirmar*, *corregir* o *cancelar*.',
      { parse_mode: 'Markdown' }
    );
  }
}

// ─── Callbacks de botones inline ─────────────────────────────

function handleCallback(cbq) {
  const chatId = cbq.message.chat.id;
  const msgId  = cbq.message.message_id;
  const data   = cbq.data;
  answerCallbackQuery(cbq.id);

  if (data === 'CONFIRM') {
    const pending = getMergedData(chatId);
    if (!pending) {
      sendMessage(chatId, '⚠️ Los datos han expirado. Envía el albarán de nuevo.');
      return;
    }
    writeMovement(pending, chatId);
    clearAllPendingData(chatId);
    setState(chatId, { mode: 'IDLE' });
    editMessage(chatId, msgId, '✅ *Registro guardado correctamente.*', { parse_mode: 'Markdown' });
    sendStockSummaryAfterMovement(chatId, pending);
    notifyGroup(pending, chatId);
    return;
  }

  if (data === 'CANCEL') {
    clearAllPendingData(chatId);
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
    const labels = {
      tipo:         'tipo (ENTRADA o SALIDA)',
      fecha:        'fecha (DD/MM/YYYY)',
      cliente:      'nombre del cliente',
      orden_carga:  'número de orden de carga',
      matricula:    'matrícula del vehículo',
      total_palets: 'total de palets',
      observaciones:'observaciones',
    };
    sendMessage(chatId,
      `✏️ Escribe el valor correcto para *${labels[field] || field}*:`,
      { parse_mode: 'Markdown' }
    );
  }
}

// ─── Mensajes de confirmación y edición ──────────────────────

function sendConfirmationMessage(chatId, data) {
  const tipoEmoji = data.tipo === 'ENTRADA' ? '📥' : '📤';
  const lineasTxt = (data.lineas || []).map((l, i) =>
    `  ${i+1}. *${l.referencia || '?'}* — ${l.descripcion || ''} — *${l.palets} palets*`
  ).join('\n');

  const text =
    `${tipoEmoji} *Resumen del movimiento:*\n\n` +
    `🔄 Tipo:          *${data.tipo || '?'}*\n` +
    `📅 Fecha:         *${data.fecha || '?'}*\n` +
    `👤 Cliente:       *${data.cliente || '?'}*\n\n` +
    `📦 *Líneas de producto:*\n${lineasTxt || '  (ninguna)'}\n` +
    `📦 Total palets:  *${data.total_palets || 0}*\n\n` +
    `🚛 Matrícula:     *${data.matricula || '?'}*\n` +
    `📋 Orden de carga:*${data.orden_carga || '?'}*\n` +
    `📝 Observaciones: *${data.observaciones || '—'}*\n\n` +
    `¿Son correctos estos datos?`;

  const keyboard = { inline_keyboard: [[
    { text: '✅ Confirmar y guardar', callback_data: 'CONFIRM'   },
    { text: '✏️ Corregir campo',      callback_data: 'EDIT_MENU' },
    { text: '❌ Cancelar',            callback_data: 'CANCEL'    },
  ]]};

  sendMessage(chatId, text, { parse_mode: 'Markdown', reply_markup: JSON.stringify(keyboard) });
}

function sendEditMenu(chatId) {
  const keyboard = { inline_keyboard: [
    [
      { text: '🔄 Tipo',          callback_data: 'EDIT_tipo'         },
      { text: '📅 Fecha',         callback_data: 'EDIT_fecha'        },
    ],
    [
      { text: '👤 Cliente',       callback_data: 'EDIT_cliente'      },
      { text: '📦 Total palets',  callback_data: 'EDIT_total_palets' },
    ],
    [
      { text: '📋 Orden de carga',callback_data: 'EDIT_orden_carga'  },
      { text: '🚛 Matrícula',     callback_data: 'EDIT_matricula'    },
    ],
    [
      { text: '📝 Observaciones', callback_data: 'EDIT_observaciones'},
    ],
  ]};
  sendMessage(chatId, '¿Qué campo quieres corregir?', { reply_markup: JSON.stringify(keyboard) });
}

function sendHelp(chatId) {
  sendMessage(chatId,
    '📖 *Ayuda — Sistema de Almacén*\n\n' +
    '*Registro en 2 pasos:*\n' +
    '1️⃣ Foto del *albarán* → referencias y cantidad de palets\n' +
    '2️⃣ Foto de la *carta de porte* → orden de carga y matrícula\n\n' +
    '*Comandos:*\n' +
    '/start — Reiniciar\n' +
    '/stock — Ver inventario actual\n' +
    '/cancelar — Cancelar registro en curso\n' +
    '/ayuda — Esta ayuda\n\n' +
    '*Consejos para mejores fotos:*\n' +
    '• Buena iluminación, sin sombras ni reflejos\n' +
    '• El documento completo visible\n' +
    '• Texto legible, sin movimiento',
    { parse_mode: 'Markdown' }
  );
}

function notifyGroup(data, operarioChatId) {
  const config = getConfig();
  if (!config.NOTIF_GROUP_ID) return;
  const tipoEmoji = data.tipo === 'ENTRADA' ? '📥' : '📤';
  sendMessage(config.NOTIF_GROUP_ID,
    `${tipoEmoji} *Nuevo movimiento* — ${data.tipo}\n` +
    `Orden: ${data.orden_carga} | Matrícula: ${data.matricula}\n` +
    `Cliente: ${data.cliente} | Total: ${data.total_palets} palets\n` +
    `Operario: ${operarioChatId}`,
    { parse_mode: 'Markdown' }
  );
}

// ─── Registrar webhook (ejecutar desde menú Sheets) ──────────

function registerWebhook() {
  const config = getConfig();
  const token  = config.TELEGRAM_BOT_TOKEN;
  const url    = config.WEBHOOK_URL;
  if (!token || !url) {
    SpreadsheetApp.getUi().alert('❌ Faltan TELEGRAM_BOT_TOKEN o WEBHOOK_URL en CONFIG.');
    return;
  }
  const res  = UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/setWebhook?url=${encodeURIComponent(url)}`);
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
