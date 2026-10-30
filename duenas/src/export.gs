// ============================================================
// EXPORT.GS — Exportar informe .xlsx copiando las pestañas tal cual
// ============================================================

function exportarInformeProveedor() {
  const ss = getSS();
  const ui = SpreadsheetApp.getUi();

  ss.toast('⏳ Generando informe Excel...', '📊 Exportar', 30);

  let tempId = null;
  let fileId = null;
  try {
    const tempSS = SpreadsheetApp.create('_tmp_almacen_export');
    tempId = tempSS.getId();

    const hojas = [
      ['STOCK',            'Stock Actual'],
      ['ENTRADAS',         'Entradas'],
      ['SALIDAS',          'Salidas'],
      ['STOCK MAQUINARIA', 'Maquinaria'],
    ];

    var copied = 0;
    hojas.forEach(function(par) {
      var srcSheet = ss.getSheetByName(par[0]);
      if (!srcSheet || srcSheet.getLastRow() < 2) return;
      var copia = srcSheet.copyTo(tempSS);
      copia.setName(par[1]);

      if (par[0] === 'STOCK') _exportStock(copia);
      // ENTRADAS y SALIDAS se exportan tal cual (hoja entera sin filtrar)

      copied++;
    });

    if (!copied) {
      DriveApp.getFileById(tempId).setTrashed(true);
      ui.alert('Sin datos', 'No hay pestañas con datos para exportar.', ui.ButtonSet.OK);
      return;
    }

    var nombresValidos = ['Stock Actual', 'Entradas', 'Salidas', 'Maquinaria'];
    tempSS.getSheets().forEach(function(s) {
      if (nombresValidos.indexOf(s.getName()) < 0 && tempSS.getSheets().length > 1)
        tempSS.deleteSheet(s);
    });

    SpreadsheetApp.flush();
    Utilities.sleep(800);

    const token = ScriptApp.getOAuthToken();
    const blob  = UrlFetchApp.fetch(
      'https://docs.google.com/spreadsheets/d/' + tempId + '/export?format=xlsx',
      { headers: { Authorization: 'Bearer ' + token } }
    ).getBlob();

    DriveApp.getFileById(tempId).setTrashed(true);
    tempId = null;

    const fileName = 'Informe_Almacen_JulioNunez_Duenas_' +
      Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyyMMdd') + '.xlsx';
    blob.setName(fileName);
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileId = file.getId();

    const html = HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
      'body{font-family:Arial,sans-serif;padding:18px;text-align:center;color:#333}p{margin:6px 0}' +
      '.btn{display:inline-block;margin-top:12px;padding:10px 22px;background:#2E75B6;color:#fff;' +
      'text-decoration:none;border-radius:5px;font-size:14px;font-weight:bold;cursor:pointer;border:none}' +
      '.btn:hover{background:#1A5C99}.note{font-size:11px;color:#888;margin-top:10px}' +
      '</style></head><body>' +
      '<p style="font-size:22px">📊</p>' +
      '<p><strong>' + fileName + '</strong></p>' +
      '<button class="btn" onclick="descargar()">📥 Descargar</button>' +
      '<p class="note" id="msg">Haz clic para descargar el archivo a tu equipo.</p>' +
      '<script>var fid="' + fileId + '";' +
      'function descargar(){' +
      'window.open("https://drive.google.com/uc?export=download&id="+fid);' +
      'document.getElementById("msg").textContent="Descargando...";' +
      'setTimeout(function(){google.script.run' +
      '.withSuccessHandler(function(){google.script.host.close();})' +
      '.withFailureHandler(function(){google.script.host.close();})' +
      '.borrarArchivoExport(fid);},4000);}' +
      '<\/script></body></html>'
    ).setWidth(320).setHeight(180);

    ui.showModalDialog(html, '📊 Exportar para proveedor');

  } catch (err) {
    Logger.log('exportarInformeProveedor: ' + err + '\n' + err.stack);
    if (tempId) try { DriveApp.getFileById(tempId).setTrashed(true); } catch (_) {}
    if (fileId)  try { DriveApp.getFileById(fileId).setTrashed(true);  } catch (_) {}
    ui.alert('Error', 'No se pudo generar el informe:\n' + err.message, ui.ButtonSet.OK);
  }
}

// ─── STOCK: solo positivos, sin columna Estado ────────────────────────────────
// Estructura STOCK: Cliente(0) | Ref(1) | Pzs(2) | StockPal(3) | StockPiezas(4) | Estado(5)
function _exportStock(sheet) {
  if (sheet.getLastRow() < 3) return;
  var nCols = sheet.getLastColumn();
  var nRows = sheet.getLastRow() - 2;

  sheet.getRange(3, 1, nRows, nCols).breakApart();
  var raw = sheet.getRange(3, 1, nRows, nCols).getValues();

  // Fill-forward + filtrar stock > 0 + quitar Estado (col índice 5)
  var lCli = '', lRef = '';
  var data = [];
  raw.forEach(function(r) {
    if (r[0]) lCli = r[0]; else r[0] = lCli;
    if (r[1]) lRef = r[1]; else r[1] = lRef;
    if (Number(r[3]) > 0) data.push(r.slice(0, 5)); // Cliente|Ref|Pzs|StockPal|StockPiezas
  });

  sheet.getRange(3, 1, Math.max(nRows + 50, 300), nCols).clearContent().clearFormat();
  if (sheet.getLastColumn() >= 6) sheet.deleteColumn(6); // eliminar columna Estado

  if (!data.length) return;

  var n = data.length;
  var rng = sheet.getRange(3, 1, n, 5);
  rng.setValues(data);
  rng.setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(3, 3, n, 3).setNumberFormat('#,##0'); // pzs/pal, stock pal, stock piezas

  var bgGrid = [], fgGrid = [], fwGrid = [];
  data.forEach(function(r, i) {
    var sp    = Number(r[3]);
    var rowBg = i % 2 === 0 ? '#EEF4FB' : '#FFFFFF';
    var spBg  = sp > 0 ? '#D9EAD3' : (sp === 0 ? '#FFF9C4' : '#F4CCCC');
    var spFg  = sp > 0 ? '#274E13' : (sp === 0 ? '#7D6608' : '#CC0000');
    bgGrid.push([rowBg, rowBg, rowBg, spBg, spBg]);
    fgGrid.push(['#000000','#000000','#000000', spFg, spFg]);
    fwGrid.push(['normal', 'normal', 'normal', 'bold', 'normal']);
  });
  rng.setBackgrounds(bgGrid);
  rng.setFontColors(fgGrid);
  rng.setFontWeights(fwGrid);

  _aplicarBordes(sheet, data, 2, n + 1, 5, function(r) { return r[0]; });
  _mergeGroups(sheet, data, 3, function(r) { return r[0]; },               [1]);
  _mergeGroups(sheet, data, 3, function(r) { return r[0] + '||' + r[1]; }, [2]);
}

function borrarArchivoExport(fileId) {
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (_) {}
}
