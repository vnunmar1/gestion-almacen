// ============================================================
// SETUP.GS — Creación e inicialización de todas las hojas
// Ejecutar UNA SOLA VEZ tras vincular el script al Sheet
// ============================================================

/**
 * Función principal de setup. Crea las 4 hojas con sus cabeceras,
 * datos de ejemplo y protecciones básicas.
 */
function setupAlmacen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    Logger.log('Iniciando setup del sistema de almacén...');

    _setupSheetConfig(ss);
    _setupSheetEntradas(ss);
    _setupSheetSalidas(ss);
    _setupSheetStock(ss);
    _setupNamedRanges(ss);

    // Reordenar hojas
    const order = ['CONFIG', 'STOCK', 'ENTRADAS', 'SALIDAS'];
    order.reverse().forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (sheet) ss.setActiveSheet(sheet), ss.moveActiveSheet(1);
    });

    ss.setActiveSheet(ss.getSheetByName('STOCK'));

    ui.alert(
      '✅ Setup completado',
      'Sistema de almacén inicializado correctamente.\n\n' +
      'Próximos pasos:\n' +
      '1. Completa los valores de CONFIG (Token Telegram, API Key Gemini, IDs autorizados)\n' +
      '2. Publica el script como Web App\n' +
      '3. Copia la URL del Web App en CONFIG > WEBHOOK_URL\n' +
      '4. Registra el webhook con Telegram (ver README)',
      ui.ButtonSet.OK
    );

    Logger.log('Setup completado correctamente.');
  } catch (e) {
    Logger.log('Error en setup: ' + e);
    ui.alert('❌ Error en setup', e.toString(), ui.ButtonSet.OK);
  }
}

// ─────────────────────────────────────────────────────────────
// HOJA: CONFIG
// ─────────────────────────────────────────────────────────────
function _setupSheetConfig(ss) {
  let sheet = ss.getSheetByName('CONFIG');
  if (!sheet) sheet = ss.insertSheet('CONFIG');
  sheet.clear();

  // Estilo general
  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 420);
  sheet.setColumnWidth(3, 380);

  // Título
  const title = sheet.getRange('A1:C1');
  title.merge();
  title.setValue('⚙️ CONFIGURACIÓN DEL SISTEMA');
  title.setBackground('#1A56A0');
  title.setFontColor('#FFFFFF');
  title.setFontWeight('bold');
  title.setFontSize(13);
  title.setHorizontalAlignment('center');
  sheet.setRowHeight(1, 36);

  // Cabeceras
  const headers = [['CLAVE', 'VALOR', 'DESCRIPCIÓN']];
  const headerRange = sheet.getRange('A2:C2');
  headerRange.setValues(headers);
  headerRange.setBackground('#2E75B6');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  sheet.setRowHeight(2, 28);

  // Datos de configuración
  const configData = [
    ['TELEGRAM_BOT_TOKEN', '', 'Token del bot obtenido desde @BotFather en Telegram'],
    ['GEMINI_API_KEY',     '', 'API Key de Google AI Studio (aistudio.google.com)'],
    ['ALLOWED_USERS',      '', 'IDs de Telegram autorizados separados por coma (ej: 123456,789012)'],
    ['WEBHOOK_URL',        '', 'URL del Web App publicado en Apps Script (se rellena en el paso 3)'],
    ['ALMACEN_NOMBRE',     'Almacén Principal', 'Nombre del almacén (aparece en mensajes de Telegram)'],
    ['STOCK_MINIMO_ALERTA','0', 'Avisar por Telegram cuando el stock baje de este valor (0 = desactivado)'],
    ['NOTIF_GROUP_ID',     '', 'Chat ID de grupo de supervisores para notificaciones (opcional)'],
  ];

  const dataRange = sheet.getRange(3, 1, configData.length, 3);
  dataRange.setValues(configData);

  // Estilo filas alternas
  configData.forEach((_, i) => {
    const row = sheet.getRange(3 + i, 1, 1, 3);
    row.setBackground(i % 2 === 0 ? '#EEF4FB' : '#FFFFFF');
  });

  // Columna CLAVE en negrita
  sheet.getRange(3, 1, configData.length, 1).setFontWeight('bold').setFontColor('#1A56A0');

  // Columna VALOR — resaltar los que están vacíos y son obligatorios
  sheet.getRange('B3').setBackground('#FFF3CD'); // Token Telegram
  sheet.getRange('B4').setBackground('#FFF3CD'); // Gemini API Key
  sheet.getRange('B5').setBackground('#FFF3CD'); // Allowed users

  // Bordes
  sheet.getRange(1, 1, configData.length + 2, 3)
    .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // Nota informativa
  const noteRow = configData.length + 4;
  sheet.getRange(noteRow, 1, 1, 3).merge()
    .setValue('⚠️ Los campos marcados en amarillo son obligatorios para que el sistema funcione.')
    .setBackground('#FFF3CD')
    .setFontStyle('italic')
    .setFontSize(10);

  Logger.log('Hoja CONFIG creada.');
}

// ─────────────────────────────────────────────────────────────
// HOJA: ENTRADAS
// ─────────────────────────────────────────────────────────────
function _setupSheetEntradas(ss) {
  let sheet = ss.getSheetByName('ENTRADAS');
  if (!sheet) sheet = ss.insertSheet('ENTRADAS');
  sheet.clear();

  const cols = {
    TIMESTAMP:      { width: 160, header: '🕐 Timestamp confirmación' },
    FECHA_ALBARAN:  { width: 110, header: '📅 Fecha albarán' },
    REFERENCIA:     { width: 140, header: '🏷️ Referencia' },
    CLIENTE:        { width: 180, header: '👤 Cliente' },
    PALETS:         { width:  80, header: '📦 Palets' },
    MATRICULA:      { width: 110, header: '🚛 Matrícula' },
    OBSERVACIONES:  { width: 240, header: '📝 Observaciones' },
    TELEGRAM_USER:  { width: 130, header: '👷 Operario (TG ID)' },
  };

  _applySheetStyle(sheet, cols, '#1A7A1A', '#E8F4E8', '📥 REGISTRO DE ENTRADAS');

  // Datos de ejemplo
  const exampleData = [
    [new Date('2025-06-01T08:32:00'), '01/06/2025', 'REF-001-A', 'CLIENTE EJEMPLO S.L.', 10, '1234-ABC', '', '123456789'],
    [new Date('2025-06-02T11:15:00'), '02/06/2025', 'REF-002-B', 'DISTRIBUCIONES XYZ',   5, '5678-DEF', 'Urgente', '123456789'],
    [new Date('2025-06-03T09:00:00'), '03/06/2025', 'REF-001-A', 'CLIENTE EJEMPLO S.L.', 8, '9012-GHI', '', '987654321'],
  ];

  _insertExampleData(sheet, exampleData);
  _applyColumnFormats(sheet, { dateTimeCols: [1], dateCols: [2], numberCols: [5] });

  Logger.log('Hoja ENTRADAS creada.');
}

// ─────────────────────────────────────────────────────────────
// HOJA: SALIDAS
// ─────────────────────────────────────────────────────────────
function _setupSheetSalidas(ss) {
  let sheet = ss.getSheetByName('SALIDAS');
  if (!sheet) sheet = ss.insertSheet('SALIDAS');
  sheet.clear();

  const cols = {
    TIMESTAMP:      { width: 160, header: '🕐 Timestamp confirmación' },
    FECHA_ALBARAN:  { width: 110, header: '📅 Fecha albarán' },
    REFERENCIA:     { width: 140, header: '🏷️ Referencia' },
    CLIENTE:        { width: 180, header: '👤 Cliente' },
    PALETS:         { width:  80, header: '📦 Palets' },
    MATRICULA:      { width: 110, header: '🚛 Matrícula' },
    OBSERVACIONES:  { width: 240, header: '📝 Observaciones' },
    TELEGRAM_USER:  { width: 130, header: '👷 Operario (TG ID)' },
  };

  _applySheetStyle(sheet, cols, '#B22222', '#FBE8E8', '📤 REGISTRO DE SALIDAS');

  const exampleData = [
    [new Date('2025-06-02T14:20:00'), '02/06/2025', 'REF-001-A', 'CLIENTE EJEMPLO S.L.', 4, '1111-AAA', '', '123456789'],
    [new Date('2025-06-03T16:45:00'), '03/06/2025', 'REF-002-B', 'DISTRIBUCIONES XYZ',   3, '2222-BBB', 'Parcial', '987654321'],
  ];

  _insertExampleData(sheet, exampleData);
  _applyColumnFormats(sheet, { dateTimeCols: [1], dateCols: [2], numberCols: [5] });

  Logger.log('Hoja SALIDAS creada.');
}

// ─────────────────────────────────────────────────────────────
// HOJA: STOCK
// ─────────────────────────────────────────────────────────────
function _setupSheetStock(ss) {
  let sheet = ss.getSheetByName('STOCK');
  if (!sheet) sheet = ss.insertSheet('STOCK');
  sheet.clear();

  // Título
  const title = sheet.getRange('A1:G1');
  title.merge();
  title.setValue('📊 STOCK ACTUAL DEL ALMACÉN');
  title.setBackground('#1A56A0');
  title.setFontColor('#FFFFFF');
  title.setFontWeight('bold');
  title.setFontSize(13);
  title.setHorizontalAlignment('center');
  sheet.setRowHeight(1, 36);

  // Nota dinámica
  const noteRange = sheet.getRange('A2:G2');
  noteRange.merge();
  noteRange.setFormula('="🔄 Actualizado automáticamente — Último movimiento: "&TEXT(MAX(ENTRADAS!A:A,SALIDAS!A:A),"DD/MM/YYYY HH:MM")');
  noteRange.setBackground('#EEF4FB');
  noteRange.setFontStyle('italic');
  noteRange.setFontSize(10);
  noteRange.setHorizontalAlignment('center');

  // Cabeceras
  const headers = [['🏷️ Referencia', '👤 Cliente', '📥 Total Entradas', '📤 Total Salidas', '📦 Stock Actual', '🟢 Estado', '📅 Últ. Movimiento']];
  const headerRange = sheet.getRange('A3:G3');
  headerRange.setValues(headers);
  headerRange.setBackground('#2E75B6');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setRowHeight(3, 30);

  // Anchos de columna
  [160, 200, 140, 140, 120, 100, 160].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Datos con fórmulas dinámicas para las 2 referencias de ejemplo
  const stockRows = [
    ['REF-001-A', 'CLIENTE EJEMPLO S.L.'],
    ['REF-002-B', 'DISTRIBUCIONES XYZ'],
  ];

  stockRows.forEach((row, i) => {
    const r = 4 + i;
    const ref = row[0];
    const cli = row[1];
    sheet.getRange(r, 1).setValue(ref);
    sheet.getRange(r, 2).setValue(cli);
    // Total entradas
    sheet.getRange(r, 3).setFormula(
      `=SUMPRODUCT((ENTRADAS!C:C="${ref}")*(ENTRADAS!D:D="${cli}")*ENTRADAS!E:E)`
    );
    // Total salidas
    sheet.getRange(r, 4).setFormula(
      `=SUMPRODUCT((SALIDAS!C:C="${ref}")*(SALIDAS!D:D="${cli}")*SALIDAS!E:E)`
    );
    // Stock actual
    sheet.getRange(r, 5).setFormula(`=C${r}-D${r}`);
    // Estado con formato condicional vía fórmula
    sheet.getRange(r, 6).setFormula(`=IF(E${r}>10,"🟢 OK",IF(E${r}>0,"🟡 BAJO","🔴 AGOTADO"))`);
    // Último movimiento
    sheet.getRange(r, 7).setFormula(
      `=TEXT(MAX(MAXIFS(ENTRADAS!A:A,ENTRADAS!C:C,"${ref}",ENTRADAS!D:D,"${cli}"),` +
      `MAXIFS(SALIDAS!A:A,SALIDAS!C:C,"${ref}",SALIDAS!D:D,"${cli}")),"DD/MM/YYYY HH:MM")`
    );

    // Estilo fila alterna
    sheet.getRange(r, 1, 1, 7).setBackground(i % 2 === 0 ? '#EEF4FB' : '#FFFFFF');
    sheet.getRange(r, 5).setFontWeight('bold');
  });

  // Formato condicional para columna Stock
  const stockRange = sheet.getRange('E4:E100');
  const rules = sheet.getConditionalFormatRules();

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(1)
    .setBackground('#F4CCCC').setFontColor('#CC0000')
    .setRanges([stockRange]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, 10)
    .setBackground('#FCF8C8').setFontColor('#7D6608')
    .setRanges([stockRange]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(10)
    .setBackground('#D9EAD3').setFontColor('#274E13')
    .setRanges([stockRange]).build());

  sheet.setConditionalFormatRules(rules);

  // Bordes generales
  sheet.getRange(3, 1, stockRows.length + 1, 7)
    .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // Inmovilizar filas de cabecera
  sheet.setFrozenRows(3);

  Logger.log('Hoja STOCK creada.');
}

// ─────────────────────────────────────────────────────────────
// FUNCIÓN HELPER: Añadir nueva referencia al stock dinámicamente
// (llamada desde sheets.gs tras cada writeMovement)
// ─────────────────────────────────────────────────────────────
function addStockRowIfMissing(ss, referencia, cliente) {
  const sheet = ss.getSheetByName('STOCK');
  const data = sheet.getDataRange().getValues();

  // Buscar si ya existe la combinación ref+cliente (desde fila 4)
  for (let i = 3; i < data.length; i++) {
    if (data[i][0] === referencia && data[i][1] === cliente) return; // ya existe
  }

  // Añadir nueva fila
  const r = data.length + 1;
  sheet.getRange(r, 1).setValue(referencia);
  sheet.getRange(r, 2).setValue(cliente);
  sheet.getRange(r, 3).setFormula(
    `=SUMPRODUCT((ENTRADAS!C:C="${referencia}")*(ENTRADAS!D:D="${cliente}")*ENTRADAS!E:E)`
  );
  sheet.getRange(r, 4).setFormula(
    `=SUMPRODUCT((SALIDAS!C:C="${referencia}")*(SALIDAS!D:D="${cliente}")*SALIDAS!E:E)`
  );
  sheet.getRange(r, 5).setFormula(`=C${r}-D${r}`);
  sheet.getRange(r, 6).setFormula(`=IF(E${r}>10,"🟢 OK",IF(E${r}>0,"🟡 BAJO","🔴 AGOTADO"))`);
  sheet.getRange(r, 7).setFormula(
    `=TEXT(MAX(MAXIFS(ENTRADAS!A:A,ENTRADAS!C:C,"${referencia}",ENTRADAS!D:D,"${cliente}"),` +
    `MAXIFS(SALIDAS!A:A,SALIDAS!C:C,"${referencia}",SALIDAS!D:D,"${cliente}")),"DD/MM/YYYY HH:MM")`
  );

  const rowRange = sheet.getRange(r, 1, 1, 7);
  rowRange.setBackground((r % 2 === 0) ? '#EEF4FB' : '#FFFFFF');
  sheet.getRange(r, 5).setFontWeight('bold');
  rowRange.setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
}

// ─────────────────────────────────────────────────────────────
// HELPERS INTERNOS
// ─────────────────────────────────────────────────────────────
function _applySheetStyle(sheet, cols, titleBg, rowBg, titleText) {
  const colNames = Object.keys(cols);
  const colDefs = Object.values(cols);

  // Anchos
  colDefs.forEach((c, i) => sheet.setColumnWidth(i + 1, c.width));

  // Título
  const title = sheet.getRange(1, 1, 1, colNames.length);
  title.merge();
  title.setValue(titleText);
  title.setBackground(titleBg);
  title.setFontColor('#FFFFFF');
  title.setFontWeight('bold');
  title.setFontSize(13);
  title.setHorizontalAlignment('center');
  sheet.setRowHeight(1, 36);

  // Cabeceras
  const headerRange = sheet.getRange(2, 1, 1, colNames.length);
  headerRange.setValues([colDefs.map(c => c.header)]);
  headerRange.setBackground('#404040');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setRowHeight(2, 28);

  // Inmovilizar filas
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);
}

function _insertExampleData(sheet, data) {
  if (!data || !data.length) return;
  const startRow = 3;
  sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);

  data.forEach((_, i) => {
    sheet.getRange(startRow + i, 1, 1, data[0].length)
      .setBackground(i % 2 === 0 ? '#F9F9F9' : '#FFFFFF');
  });
}

function _applyColumnFormats(sheet, { dateTimeCols = [], dateCols = [], numberCols = [] }) {
  dateTimeCols.forEach(c => {
    sheet.getRange(3, c, sheet.getMaxRows() - 2)
      .setNumberFormat('dd/mm/yyyy hh:mm');
  });
  dateCols.forEach(c => {
    sheet.getRange(3, c, sheet.getMaxRows() - 2)
      .setNumberFormat('dd/mm/yyyy');
  });
  numberCols.forEach(c => {
    sheet.getRange(3, c, sheet.getMaxRows() - 2)
      .setNumberFormat('#,##0');
  });
}

function _setupNamedRanges(ss) {
  // Rangos nombrados para facilitar referencias en fórmulas
  const namedRanges = [
    { name: 'Config_BotToken',   sheet: 'CONFIG',   range: 'B3' },
    { name: 'Config_GeminiKey',  sheet: 'CONFIG',   range: 'B4' },
    { name: 'Config_AllowedUsers', sheet: 'CONFIG', range: 'B5' },
    { name: 'Config_WebhookUrl', sheet: 'CONFIG',   range: 'B6' },
  ];

  namedRanges.forEach(nr => {
    try {
      const sheet = ss.getSheetByName(nr.sheet);
      const range = sheet.getRange(nr.range);
      ss.setNamedRange(nr.name, range);
    } catch(e) {
      Logger.log('No se pudo crear el rango nombrado ' + nr.name + ': ' + e);
    }
  });
}

// ─────────────────────────────────────────────────────────────
// MENÚ PERSONALIZADO en Google Sheets
// ─────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏭 Almacén')
    .addItem('⚙️ Inicializar sistema (setup)', 'setupAlmacen')
    .addSeparator()
    .addItem('📊 Ver stock completo (Telegram)', 'sendFullStockToAdmin')
    .addItem('🔄 Registrar webhook de Telegram', 'registerWebhook')
    .addSeparator()
    .addItem('🧪 Test: Simular entrada', 'testSimulateEntry')
    .addItem('📋 Ver logs de ejecución', 'openLogs')
    .addToUi();
}

function openLogs() {
  const url = 'https://script.google.com/home/executions';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(`<script>window.open('${url}','_blank');google.script.host.close();</script>`),
    'Abriendo logs...'
  );
}

function testSimulateEntry() {
  const testData = {
    tipo: 'ENTRADA',
    referencia: 'REF-TEST-001',
    cliente: 'CLIENTE TEST',
    palets: 5,
    fecha: Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy'),
    matricula: '0000-TST',
    observaciones: 'Registro de prueba desde setup'
  };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writeMovement(testData, 'SISTEMA_TEST');
  SpreadsheetApp.getUi().alert('✅ Entrada de prueba registrada', JSON.stringify(testData, null, 2), SpreadsheetApp.getUi().ButtonSet.OK);
}
