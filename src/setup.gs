// ============================================================
// SETUP.GS — Creación e inicialización de todas las hojas
// Ejecutar UNA SOLA VEZ tras vincular el script al Sheet
// ============================================================

function setupAlmacen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  try {
    _setupSheetConfig(ss);
    _setupSheetEntradas(ss);
    _setupSheetSalidas(ss);
    _setupSheetStock(ss);
    _setupNamedRanges(ss);

    ['CONFIG','STOCK','ENTRADAS','SALIDAS'].reverse().forEach(name => {
      const s = ss.getSheetByName(name);
      if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(1); }
    });
    ss.setActiveSheet(ss.getSheetByName('STOCK'));

    ui.alert('✅ Setup completado',
      'Sistema inicializado.\n\nPróximos pasos:\n' +
      '1. Completa CONFIG (Token Telegram, Gemini Key, IDs)\n' +
      '2. Publica como Web App\n' +
      '3. Pega la URL en CONFIG > WEBHOOK_URL\n' +
      '4. Menú → Registrar webhook de Telegram',
      ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('❌ Error en setup', e.toString(), ui.ButtonSet.OK);
  }
}

// ─── CONFIG ──────────────────────────────────────────────────

function _setupSheetConfig(ss) {
  let sheet = ss.getSheetByName('CONFIG') || ss.insertSheet('CONFIG');
  sheet.clear();
  sheet.setColumnWidths(1, 3, 0);
  sheet.setColumnWidth(1, 240); sheet.setColumnWidth(2, 400); sheet.setColumnWidth(3, 380);

  const title = sheet.getRange('A1:C1');
  title.merge().setValue('⚙️ CONFIGURACIÓN DEL SISTEMA')
    .setBackground('#1A56A0').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sheet.setRowHeight(1, 36);

  sheet.getRange('A2:C2').setValues([['CLAVE','VALOR','DESCRIPCIÓN']])
    .setBackground('#2E75B6').setFontColor('#FFFFFF').setFontWeight('bold');

  const rows = [
    ['TELEGRAM_BOT_TOKEN', '', 'Token del bot — @BotFather en Telegram'],
    ['GEMINI_API_KEY',     '', 'API Key de Google AI Studio (aistudio.google.com)'],
    ['ALLOWED_USERS',      '', 'IDs de Telegram autorizados, separados por coma'],
    ['WEBHOOK_URL',        '', 'URL del Web App publicado (se rellena en paso 3)'],
    ['ALMACEN_NOMBRE',     'Almacén Principal', 'Nombre del almacén'],
    ['STOCK_MINIMO_ALERTA','0', 'Avisar cuando stock baje de este valor (0=desactivado)'],
    ['NOTIF_GROUP_ID',     '', 'Chat ID de grupo de supervisores (opcional)'],
  ];

  sheet.getRange(3, 1, rows.length, 3).setValues(rows);
  rows.forEach((_, i) => {
    sheet.getRange(3+i, 1, 1, 3).setBackground(i%2===0?'#EEF4FB':'#FFFFFF');
  });
  sheet.getRange(3,1,rows.length,1).setFontWeight('bold').setFontColor('#1A56A0');
  sheet.getRange('B3').setBackground('#FFF3CD');
  sheet.getRange('B4').setBackground('#FFF3CD');
  sheet.getRange('B5').setBackground('#FFF3CD');
  sheet.getRange(1,1,rows.length+2,3)
    .setBorder(true,true,true,true,true,true,'#CCCCCC',SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(rows.length+4, 1, 1, 3).merge()
    .setValue('⚠️ Campos en amarillo son obligatorios para que el sistema funcione.')
    .setBackground('#FFF3CD').setFontStyle('italic').setFontSize(10);
}

// ─── ENTRADAS ────────────────────────────────────────────────

function _setupSheetEntradas(ss) {
  let sheet = ss.getSheetByName('ENTRADAS') || ss.insertSheet('ENTRADAS');
  sheet.clear();

  const headers = [
    '🕐 Timestamp',       // A
    '📅 Fecha albarán',   // B
    '📋 Orden de carga',  // C
    '👤 Cliente',         // D
    '🏷️ Referencia',     // E
    '📝 Descripción',     // F
    '📦 Palets',          // G
    '🚛 Matrícula',       // H
    '🚛 Matrícula remolque', // I
    '🏢 Transportista',   // J
    '📄 Observaciones',   // K
    '👷 Operario (TG)',   // L
  ];
  const widths = [155,110,140,180,140,200,80,110,130,160,200,130];

  _buildMovSheet(sheet, headers, widths, '#1A7A1A', '📥 REGISTRO DE ENTRADAS');

  const ex = [
    [new Date('2025-06-01T08:32:00'), '01/06/2025', 'OC-2025-0847', 'CLIENTE EJEMPLO S.L.',
     'REF-001-A', 'Palet europeo tipo A', 10, '1234-ABC', '5678-DEF', 'TRANSPORTES XYZ', '', '123456789'],
    [new Date('2025-06-01T08:32:00'), '01/06/2025', 'OC-2025-0847', 'CLIENTE EJEMPLO S.L.',
     'REF-002-B', 'Caja cartón reforzada', 5, '1234-ABC', '5678-DEF', 'TRANSPORTES XYZ', 'Urgente', '123456789'],
  ];
  sheet.getRange(3,1,ex.length,ex[0].length).setValues(ex);
  ex.forEach((_,i) => sheet.getRange(3+i,1,1,ex[0].length).setBackground(i%2===0?'#F9F9F9':'#FFFFFF'));
  sheet.getRange(3,1,ex.length,1).setNumberFormat('dd/mm/yyyy hh:mm');
  sheet.getRange(3,2,ex.length,1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(3,7,ex.length,1).setNumberFormat('#,##0');
}

// ─── SALIDAS ─────────────────────────────────────────────────

function _setupSheetSalidas(ss) {
  let sheet = ss.getSheetByName('SALIDAS') || ss.insertSheet('SALIDAS');
  sheet.clear();

  const headers = [
    '🕐 Timestamp','📅 Fecha albarán','📋 Orden de carga','👤 Cliente',
    '🏷️ Referencia','📝 Descripción','📦 Palets','🚛 Matrícula',
    '🚛 Matrícula remolque','🏢 Transportista','📄 Observaciones','👷 Operario (TG)',
  ];
  const widths = [155,110,140,180,140,200,80,110,130,160,200,130];

  _buildMovSheet(sheet, headers, widths, '#B22222', '📤 REGISTRO DE SALIDAS');

  const ex = [
    [new Date('2025-06-02T14:20:00'), '02/06/2025', 'OC-2025-0901', 'CLIENTE EJEMPLO S.L.',
     'REF-001-A', 'Palet europeo tipo A', 4, '1111-AAA', '', 'LOGÍSTICA SUR', '', '123456789'],
  ];
  sheet.getRange(3,1,ex.length,ex[0].length).setValues(ex);
  sheet.getRange(3,1,1,ex[0].length).setBackground('#F9F9F9');
  sheet.getRange(3,1,1).setNumberFormat('dd/mm/yyyy hh:mm');
  sheet.getRange(3,2,1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(3,7,1).setNumberFormat('#,##0');
}

// ─── STOCK ───────────────────────────────────────────────────

function _setupSheetStock(ss) {
  let sheet = ss.getSheetByName('STOCK') || ss.insertSheet('STOCK');
  sheet.clear();

  sheet.getRange('A1:G1').merge().setValue('📊 STOCK ACTUAL DEL ALMACÉN')
    .setBackground('#1A56A0').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(13).setHorizontalAlignment('center');
  sheet.setRowHeight(1,36);

  sheet.getRange('A2:G2').merge()
    .setFormula('="🔄 Actualizado automáticamente — Último mov.: "' +
      '&TEXT(MAX(ENTRADAS!A:A,SALIDAS!A:A),"DD/MM/YYYY HH:MM")')
    .setBackground('#EEF4FB').setFontStyle('italic').setFontSize(10).setHorizontalAlignment('center');

  const hdrs = [['🏷️ Referencia','👤 Cliente','📥 Entradas','📤 Salidas','📦 Stock','🚦 Estado','📅 Últ. mov.']];
  sheet.getRange('A3:G3').setValues(hdrs)
    .setBackground('#2E75B6').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(3,30);
  [160,200,120,120,100,100,160].forEach((w,i)=>sheet.setColumnWidth(i+1,w));

  // Filas de ejemplo con fórmulas SUMPRODUCT (columna E=referencia, D=cliente, G=palets)
  const refs = [
    ['REF-001-A','CLIENTE EJEMPLO S.L.'],
    ['REF-002-B','CLIENTE EJEMPLO S.L.'],
  ];
  refs.forEach((r,i) => {
    const row = 4+i;
    const ref = r[0]; const cli = r[1];
    sheet.getRange(row,1).setValue(ref);
    sheet.getRange(row,2).setValue(cli);
    sheet.getRange(row,3).setFormula(`=SUMPRODUCT((ENTRADAS!E:E="${ref}")*(ENTRADAS!D:D="${cli}")*ENTRADAS!G:G)`);
    sheet.getRange(row,4).setFormula(`=SUMPRODUCT((SALIDAS!E:E="${ref}")*(SALIDAS!D:D="${cli}")*SALIDAS!G:G)`);
    sheet.getRange(row,5).setFormula(`=C${row}-D${row}`);
    sheet.getRange(row,6).setFormula(`=IF(E${row}>10,"🟢 OK",IF(E${row}>0,"🟡 BAJO","🔴 AGOTADO"))`);
    sheet.getRange(row,7).setFormula(
      `=TEXT(MAX(MAXIFS(ENTRADAS!A:A,ENTRADAS!E:E,"${ref}",ENTRADAS!D:D,"${cli}"),`+
      `MAXIFS(SALIDAS!A:A,SALIDAS!E:E,"${ref}",SALIDAS!D:D,"${cli}")),"DD/MM/YYYY HH:MM")`
    );
    sheet.getRange(row,1,1,7).setBackground(i%2===0?'#EEF4FB':'#FFFFFF');
    sheet.getRange(row,5).setFontWeight('bold');
  });

  // Formato condicional stock
  const stockRange = sheet.getRange('E4:E100');
  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(1).setBackground('#F4CCCC').setFontColor('#CC0000')
    .setRanges([stockRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1,10).setBackground('#FCF8C8').setFontColor('#7D6608')
    .setRanges([stockRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(10).setBackground('#D9EAD3').setFontColor('#274E13')
    .setRanges([stockRange]).build());
  sheet.setConditionalFormatRules(rules);

  sheet.getRange(3,1,refs.length+1,7)
    .setBorder(true,true,true,true,true,true,'#CCCCCC',SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenRows(3);
}

// ─── addStockRowIfMissing (llamado desde sheets.gs) ──────────

function addStockRowIfMissing(ss, referencia, cliente) {
  const sheet = ss.getSheetByName('STOCK');
  const data  = sheet.getDataRange().getValues();
  for (let i=3; i<data.length; i++) {
    if (data[i][0] === referencia && data[i][1] === cliente) return;
  }
  const row = data.length + 1;
  sheet.getRange(row,1).setValue(referencia);
  sheet.getRange(row,2).setValue(cliente);
  sheet.getRange(row,3).setFormula(`=SUMPRODUCT((ENTRADAS!E:E="${referencia}")*(ENTRADAS!D:D="${cliente}")*ENTRADAS!G:G)`);
  sheet.getRange(row,4).setFormula(`=SUMPRODUCT((SALIDAS!E:E="${referencia}")*(SALIDAS!D:D="${cliente}")*SALIDAS!G:G)`);
  sheet.getRange(row,5).setFormula(`=C${row}-D${row}`);
  sheet.getRange(row,6).setFormula(`=IF(E${row}>10,"🟢 OK",IF(E${row}>0,"🟡 BAJO","🔴 AGOTADO"))`);
  sheet.getRange(row,7).setFormula(
    `=TEXT(MAX(MAXIFS(ENTRADAS!A:A,ENTRADAS!E:E,"${referencia}",ENTRADAS!D:D,"${cliente}"),`+
    `MAXIFS(SALIDAS!A:A,SALIDAS!E:E,"${referencia}",SALIDAS!D:D,"${cliente}")),"DD/MM/YYYY HH:MM")`
  );
  sheet.getRange(row,1,1,7).setBackground(row%2===0?'#EEF4FB':'#FFFFFF');
  sheet.getRange(row,5).setFontWeight('bold');
  sheet.getRange(row,1,1,7)
    .setBorder(true,true,true,true,true,true,'#CCCCCC',SpreadsheetApp.BorderStyle.SOLID);
}

// ─── Rangos nombrados ─────────────────────────────────────────

function _setupNamedRanges(ss) {
  [
    ['Config_BotToken',    'CONFIG','B3'],
    ['Config_GeminiKey',   'CONFIG','B4'],
    ['Config_AllowedUsers','CONFIG','B5'],
    ['Config_WebhookUrl',  'CONFIG','B6'],
  ].forEach(([name, sheetName, range]) => {
    try { ss.setNamedRange(name, ss.getSheetByName(sheetName).getRange(range)); }
    catch(e) { Logger.log('NamedRange error: '+e); }
  });
}

// ─── Helper construcción de hoja de movimientos ──────────────

function _buildMovSheet(sheet, headers, widths, titleColor, titleText) {
  headers.forEach((_,i) => sheet.setColumnWidth(i+1, widths[i]||120));
  sheet.getRange(1,1,1,headers.length).merge()
    .setValue(titleText).setBackground(titleColor)
    .setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sheet.setRowHeight(1,36);
  sheet.getRange(2,1,1,headers.length).setValues([headers])
    .setBackground('#404040').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(2,28);
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);
}

// ─── Menú personalizado ───────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏭 Almacén')
    .addItem('⚙️ Inicializar sistema (setup)', 'setupAlmacen')
    .addSeparator()
    .addItem('📊 Enviar stock por Telegram',     'sendFullStockToAdmin')
    .addItem('🔄 Registrar webhook de Telegram', 'registerWebhook')
    .addSeparator()
    .addItem('🧪 Test: Simular entrada',  'testSimulateEntry')
    .addItem('📋 Ver logs de ejecución',  'openLogs')
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
    fecha: Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy'),
    orden_carga: 'OC-TEST-001',
    cliente: 'CLIENTE TEST',
    lineas: [
      { referencia: 'REF-TEST-001', descripcion: 'Producto de prueba A', palets: 3 },
      { referencia: 'REF-TEST-002', descripcion: 'Producto de prueba B', palets: 2 },
    ],
    total_palets: 5,
    matricula: '0000-TST',
    matricula_remolque: null,
    transportista: 'TRANSPORTE TEST',
    observaciones: 'Registro de prueba desde setup',
  };
  writeMovement(testData, 'SISTEMA_TEST');
  SpreadsheetApp.getUi().alert('✅ Entrada de prueba registrada',
    '2 líneas registradas en hoja ENTRADAS:\n' +
    '• REF-TEST-001: 3 palets\n• REF-TEST-002: 2 palets',
    SpreadsheetApp.getUi().ButtonSet.OK);
}
