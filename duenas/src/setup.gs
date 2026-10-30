// ============================================================
// SETUP.GS — Inicialización: hojas, triggers y menú
// ============================================================
// El formulario es un Web App de Apps Script (index.html).
// Para obtener la URL: Implementar → Nueva implementación → Aplicación web
//   · Ejecutar como: Yo
//   · Quién tiene acceso: Cualquier usuario
// ============================================================

const SPREADSHEET_ID = '1QDXElIotoCZEErCAbabJgdYaYmO-eUXI8eD18Es-O2M';

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// Ejecutar UNA VEZ desde Apps Script para crear/resetear las hojas y los triggers.
function setup() {
  const ss = getSS();

  ['FORMULARIO'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) ss.deleteSheet(s);
  });

  _setupEntradas(ss);
  _setupSalidas(ss);
  _setupStock(ss);
  _setupHistorico(ss);

  ['HISTÓRICO', 'SALIDAS', 'ENTRADAS', 'STOCK'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(1); }
  });
  ss.setActiveSheet(ss.getSheetByName('STOCK'));

  setupTriggers(ss);

  const webAppUrl = _getWebAppUrl();
  if (webAppUrl) {
    PropertiesService.getScriptProperties().setProperty('WEBAPP_URL', webAppUrl);
  }

  Logger.log('✅ Setup completado.');
  if (webAppUrl) Logger.log('📝 Formulario: ' + webAppUrl);

  SpreadsheetApp.getUi().alert(
    '✅ Setup completado',
    webAppUrl
      ? '📝 URL del formulario:\n\n' + webAppUrl + '\n\nCompártela con los operarios.'
      : '⚠️ El formulario web aún no está desplegado.\n\nVe a: Implementar → Nueva implementación → Aplicación web\n  · Ejecutar como: Yo\n  · Acceso: Cualquier usuario\n\nDespués ejecuta setup() de nuevo para guardar la URL.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function setupTriggers(ss) {
  if (!ss) ss = getSS();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('addSpreadsheetMenu').forSpreadsheet(ss).onOpen().create();
  Logger.log('✅ Triggers instalados.');
}

function _getWebAppUrl() {
  try {
    const url = ScriptApp.getService().getUrl();
    return url || null;
  } catch (_) {
    return null;
  }
}

// ─── Menú ────────────────────────────────────────────────────

function addSpreadsheetMenu(e) {
  const ss = (e && e.source) ? e.source : getSS();
  ss.addMenu('📦 Almacén', [
    { name: '🔄 Actualizar stock ahora',                  functionName: 'actualizarStockDesdeMenu' },
    { name: '📊 Exportar informe para proveedor (.xlsx)', functionName: 'exportarInformeProveedor' },
  ]);
}

function actualizarStockDesdeMenu() {
  const ss = getSS();
  actualizarStock(ss); // re-ordena ENTRADAS/SALIDAS y recalcula STOCK (no toca HISTÓRICO)
  ss.toast('✅ Stock actualizado.', '📊 Almacén', 4);
}

// ─── Creación de hojas ───────────────────────────────────────

// ENTRADAS: Nº Albarán | Carta Porte | Fecha | Referencia | Piezas/palet | Nº Palets | Total piezas | Cliente  (8 cols)
function _setupEntradas(ss) {
  let sheet = ss.getSheetByName('ENTRADAS');
  if (!sheet) sheet = ss.insertSheet('ENTRADAS');
  sheet.clear();

  const headers = ['🔖 Nº Albarán', '🚢 Carta Porte', '📅 Fecha', '📋 Referencia', '🔢 Piezas/palet',
                   '📦 Nº Palets', '➕ Total piezas', '👤 Cliente'];
  const widths  = [110, 110, 100, 130, 110, 100, 115, 170];

  sheet.getRange(1, 1, 1, headers.length).merge()
    .setValue('📥 ENTRADAS DE ALMACÉN · Julio Núñez S.L.')
    .setBackground('#1B5E20').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sheet.setRowHeight(1, 36);
  sheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setBackground('#2E7D32').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(2, 28);
  sheet.setFrozenRows(2);
  headers.forEach((_, i) => sheet.setColumnWidth(i + 1, widths[i]));
}

// SALIDAS: Fecha | Referencia | Piezas/palet | Nº Palets | Total piezas | Cliente  (6 cols)
function _setupSalidas(ss) {
  let sheet = ss.getSheetByName('SALIDAS');
  if (!sheet) sheet = ss.insertSheet('SALIDAS');
  sheet.clear();

  const headers = ['📅 Fecha', '📋 Referencia', '🔢 Piezas/palet',
                   '📦 Nº Palets', '➕ Total piezas', '👤 Cliente'];
  const widths  = [100, 130, 110, 100, 115, 170];

  sheet.getRange(1, 1, 1, headers.length).merge()
    .setValue('📤 SALIDAS DE ALMACÉN · Julio Núñez S.L.')
    .setBackground('#B71C1C').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sheet.setRowHeight(1, 36);
  sheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setBackground('#C62828').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(2, 28);
  sheet.setFrozenRows(2);
  headers.forEach((_, i) => sheet.setColumnWidth(i + 1, widths[i]));
}

// HISTÓRICO: 11 cols (K=Origen: APP|MANUAL)
function _setupHistorico(ss) {
  let sheet = ss.getSheetByName('HISTÓRICO');
  if (!sheet) sheet = ss.insertSheet('HISTÓRICO');
  const headers = ['📅 Fecha', '🔄 Tipo', '🔖 Nº Albarán', '🚢 Carta Porte',
                   '📋 Referencia', '🔢 Piezas/palet', '📦 Nº Palets', '➕ Total piezas',
                   '👤 Cliente', '📝 Observaciones', '📍 Origen'];
  sheet.getRange(1, 1, 1, headers.length).merge()
    .setValue('📋 HISTÓRICO DE MOVIMIENTOS')
    .setBackground('#37474F').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sheet.setRowHeight(1, 36);
  sheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setBackground('#546E7A').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(2, 28);
  sheet.setFrozenRows(2);
  sheet.setColumnWidth(11, 90);
}

// STOCK: Cliente | Referencia | Piezas/pal | Stock Pal. | Stock Piezas | Estado  (6 cols)
function _setupStock(ss) {
  let sheet = ss.getSheetByName('STOCK');
  if (!sheet) sheet = ss.insertSheet('STOCK');
  sheet.clear();

  const headers = ['👤 Cliente', '📋 Referencia', '🔢 Piezas/palet',
                   '📦 Stock palets', '➕ Stock piezas', '🚦 Estado'];
  const widths  = [200, 130, 110, 120, 120, 110];

  sheet.getRange(1, 1, 1, headers.length).merge()
    .setValue('📊 STOCK ACTUAL DEL ALMACÉN')
    .setBackground('#1A56A0').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sheet.setRowHeight(1, 36);

  sheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setBackground('#2E75B6').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(2, 28);
  sheet.setFrozenRows(2);
  headers.forEach((_, i) => sheet.setColumnWidth(i + 1, widths[i]));

  sheet.getRange(3, 1, 1, headers.length).merge()
    .setValue('Sin movimientos todavía. Rellena el formulario para añadir entradas/salidas.')
    .setFontStyle('italic').setFontColor('#888888').setHorizontalAlignment('center');
}
