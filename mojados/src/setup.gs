// ============================================================
// SETUP.GS — Inicialización: hojas, triggers y menú
// ============================================================
// El formulario es un Web App de Apps Script (index.html).
// Para obtener la URL: Implementar → Nueva implementación → Aplicación web
//   · Ejecutar como: Yo
//   · Quién tiene acceso: Cualquier usuario
// ============================================================

const SPREADSHEET_ID = '11GwVh2za1o0WZKJvapVanujr7HSJdX6LI1W9J9blXsQ';

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
    { name: '🔄 Actualizar stock ahora',                  functionName: 'actualizarStockDesdeMenu'   },
    { name: '📊 Exportar informe para proveedor (.xlsx)', functionName: 'exportarInformeProveedor'   },
    null,
    { name: '📋 Importar inventario físico...',           functionName: 'importarInventario'         },
  ]);
}

// ── Diálogo: importar una hoja de inventario como nueva base de stock ─────────
function importarInventario() {
  var hoy = Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd');
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,sans-serif;padding:16px;font-size:13px;color:#333}' +
    'h3{margin:0 0 12px;font-size:14px}' +
    'label{display:block;margin:10px 0 3px;font-weight:bold;font-size:12px}' +
    'input{width:100%;padding:6px 8px;box-sizing:border-box;border:1px solid #ccc;border-radius:3px;font-size:13px}' +
    '.nota{font-size:11px;color:#888;margin:4px 0 0}' +
    '.btn{margin-top:14px;padding:8px 20px;background:#2E75B6;color:#fff;border:none;border-radius:4px;font-size:13px;cursor:pointer;font-weight:bold}' +
    '.btn:hover{background:#1A5C99}.btn:disabled{background:#aaa;cursor:default}' +
    '.msg{margin-top:10px;font-size:12px;min-height:16px}' +
    '</style></head><body>' +
    '<h3>📋 Importar inventario físico</h3>' +
    '<label>Nombre de la hoja con el inventario</label>' +
    '<input type="text" id="hoja" placeholder="ej: stock20260731">' +
    '<p class="nota">Formato esperado: Cliente | Referencia | Piezas/pal | Palets</p>' +
    '<label>Fecha del inventario</label>' +
    '<input type="date" id="fecha" value="' + hoy + '">' +
    '<br><button class="btn" onclick="doImport()">📥 Importar como base</button>' +
    '<p class="msg" id="msg"></p>' +
    '<script>function doImport(){' +
    'var h=document.getElementById("hoja").value.trim();' +
    'var f=document.getElementById("fecha").value;' +
    'if(!h){document.getElementById("msg").textContent="⚠️ Indica el nombre de la hoja";return;}' +
    'if(!f){document.getElementById("msg").textContent="⚠️ Indica la fecha del inventario";return;}' +
    'document.getElementById("msg").textContent="⏳ Importando...";' +
    'document.querySelector(".btn").disabled=true;' +
    'google.script.run' +
    '.withSuccessHandler(function(msg){document.getElementById("msg").textContent="✅ "+msg;setTimeout(function(){google.script.host.close();},2500);})' +
    '.withFailureHandler(function(e){document.getElementById("msg").textContent="❌ "+e.message;document.querySelector(".btn").disabled=false;})' +
    '._ejecutarImportarInventario(h,f);}' +
    '<\/script></body></html>'
  ).setWidth(380).setHeight(290);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Importar inventario');
}

// ── Lógica de importación (llamada desde el diálogo HTML) ─────────────────────
function _ejecutarImportarInventario(sheetName, fechaStr) {
  var ss = getSS();
  var srcSheet = ss.getSheetByName(sheetName);
  if (!srcSheet) throw new Error('No se encontró la hoja "' + sheetName + '"');

  // Leer todas las filas y filtrar solo las que tienen datos válidos
  // (ignora automáticamente filas de título y cabeceras)
  var allRaw = srcSheet.getRange(1, 1, srcSheet.getLastRow(), 4).getValues();
  var inventario = allRaw.filter(function(r) {
    var pzs = Number(r[2]);
    var pal = Number(r[3]);
    return String(r[0] || '').trim() && String(r[1] || '').trim() && pzs > 0 && pal > 0;
  });
  if (!inventario.length)
    throw new Error('La hoja "' + sheetName + '" no tiene datos válidos (se necesitan: cliente, referencia, piezas/pal > 0 y palets > 0)');

  // Parsear fecha
  var m = String(fechaStr || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  var fechaCorte = m ? new Date(+m[1], +m[2] - 1, +m[3]) : new Date();

  // Crear / limpiar hoja STOCK BASE
  var baseSheet = ss.getSheetByName('STOCK BASE');
  if (!baseSheet) {
    baseSheet = ss.insertSheet('STOCK BASE');
    // mover al final de las pestañas
    ss.moveActiveSheet(ss.getSheets().length);
  }
  baseSheet.clear();

  // Cabecera
  var nCols = 5;
  var headers = ['👤 Cliente', '📋 Referencia', '🔢 Piezas/palet', '📦 Stock Palets', '📅 Fecha Inventario'];
  var widths  = [200, 130, 110, 120, 130];
  baseSheet.getRange(1, 1, 1, nCols).merge()
    .setValue('📋 STOCK BASE · Inventario ' + Utilities.formatDate(fechaCorte, 'Europe/Madrid', 'dd/MM/yyyy'))
    .setBackground('#37474F').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  baseSheet.setRowHeight(1, 36);
  baseSheet.getRange(2, 1, 1, nCols).setValues([headers])
    .setBackground('#546E7A').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');
  baseSheet.setRowHeight(2, 28);
  baseSheet.setFrozenRows(2);
  headers.forEach(function(_, i) { baseSheet.setColumnWidth(i + 1, widths[i]); });

  // Datos
  var data = inventario.map(function(r) {
    return [String(r[0]).trim().toUpperCase(), String(r[1]).trim(),
            Number(r[2]), Number(r[3]), fechaCorte];
  });
  var rng = baseSheet.getRange(3, 1, data.length, nCols);
  rng.setValues(data);
  rng.setHorizontalAlignment('center').setVerticalAlignment('middle');
  baseSheet.getRange(3, 3, data.length, 2).setNumberFormat('#,##0');
  baseSheet.getRange(3, 5, data.length, 1).setNumberFormat('dd/mm/yyyy');

  // Colorear filas alternas
  var bgGrid = data.map(function(_, i) {
    var bg = i % 2 === 0 ? '#F5F5F5' : '#FFFFFF';
    return [bg, bg, bg, bg, bg];
  });
  rng.setBackgrounds(bgGrid);

  // Recalcular stock desde la nueva base
  actualizarStock(ss);

  return data.length + ' líneas importadas desde "' + sheetName + '" · Base: ' +
         Utilities.formatDate(fechaCorte, 'Europe/Madrid', 'dd/MM/yyyy');
}

function mostrarUrlFormulario() {
  const url = PropertiesService.getScriptProperties().getProperty('WEBAPP_URL')
    || _getWebAppUrl()
    || '⚠️ Sin desplegar — ve a Implementar → Nueva implementación → Aplicación web.';
  getSS().toast(url, '📝 Formulario', 12);
  Logger.log('URL formulario: ' + url);
}

function actualizarStockDesdeMenu() {
  const ss = getSS();
  actualizarStock(ss); // re-ordena ENTRADAS/SALIDAS y recalcula STOCK (no toca HISTÓRICO)
  ss.toast('✅ Stock actualizado.', '📊 Almacén', 4);
}

function reiniciarHojaStock() {
  const ss = getSS();
  _setupStock(ss);
  actualizarStock(ss);
  ss.toast('✅ Hoja STOCK reiniciada y stock actualizado.', '📊 Almacén', 5);
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
  if (!sheet) { _crearHojaHistorico(ss); return; }
  const headers = ['📅 Fecha', '🔄 Tipo', '🔖 Nº Albarán', '🚢 Carta Porte',
                   '📋 Referencia', '🔢 Piezas/palet', '📦 Nº Palets', '➕ Total piezas',
                   '👤 Cliente', '📝 Observaciones', '📍 Origen'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setBackground('#546E7A').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');
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
