// ============================================================
// STOCK.GS — Re-ordenar ENTRADAS/SALIDAS y calcular STOCK
// ============================================================
// ENTRADAS cols: Nº Albarán | Carta Porte | Fecha | Referencia | Piezas/pal | NºPalets | TotPiezas | Cliente  (8)
// SALIDAS cols:  Fecha | Referencia | Piezas/pal | NºPalets | TotPiezas | Cliente  (6)
// STOCK cols:    Cliente | Referencia | Piezas/pal | Stock Pal. | Stock Piezas | Estado  (6)

function _cleanAlbaran(v) {
  const s = String(v || '').trim();
  return /^\d{1,3}(\.\d{3})+$/.test(s) ? s.replace(/\./g, '') : s;
}

// ── Lectura de filas desde ENTRADAS con fill-forward ─────────────────────────
// Detecta formato:
//   antigua   (7 cols): alb|fecha|ref|pzs|pal|tot|cli
//   nueva     (8 cols): alb|cmr|fecha|ref|pzs|pal|tot|cli  ← actual
function _leerFilasEntradas(sheet) {
  var result = [];
  if (!sheet || sheet.getLastRow() < 3) return result;
  var lastCol = Math.min(sheet.getLastColumn(), 8);
  var allRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  var nueva = String(allRows[0][1] || '').toLowerCase().indexOf('carta') >= 0;
  var raw = allRows.slice(1);
  var lastAlb = '', lastCmr = '', lastFecha = '', lastRef = '', lastCli = '', lastPzs = 0;
  raw.forEach(function(r) {
    var alb  = _cleanAlbaran(String(r[0] || '').trim());
    var cmr  = nueva ? String(r[1] || '').trim() : '';
    var fecha = nueva ? r[2] : r[1];
    var ref  = String((nueva ? r[3] : r[2]) || '').trim();
    var pzs  = Number(nueva ? r[4] : r[3]) || 0;
    var pal  = Number(nueva ? r[5] : r[4]) || 0;
    var tot  = Number(nueva ? r[6] : r[5]) || 0;
    var cli  = String((nueva ? r[7] : r[6]) || '').trim().toUpperCase();
    if (alb && alb !== lastAlb) {
      lastAlb = alb; lastCmr = cmr; lastCli = cli; lastPzs = pzs;
    } else {
      if (cmr) lastCmr = cmr;
      if (cli) lastCli = cli;
      if (pzs) lastPzs = pzs;
    }
    if (fecha instanceof Date || (fecha && String(fecha).trim())) lastFecha = fecha;
    if (ref) lastRef = ref;
    var pzsEfectivo = pzs || lastPzs;
    if (lastRef && pal && lastCli && pzsEfectivo) {
      result.push({ albaran: lastAlb, cmr: lastCmr, fecha: lastFecha, referencia: lastRef,
                    piezasPal: pzsEfectivo, palets: pal,
                    totPiezas: tot || pzsEfectivo * pal, cliente: lastCli });
    }
  });
  return result;
}

// ── Lectura de filas desde SALIDAS con fill-forward ───────────────────────────
// Detecta formato:
//   antigua (8 cols): alb|cmr|fecha|ref|pzs|pal|tot|cli
//   nueva   (6 cols): fecha|ref|pzs|pal|tot|cli  ← actual
function _leerFilasSalidas(sheet) {
  var result = [];
  if (!sheet || sheet.getLastRow() < 3) return result;
  var lastCol = Math.min(sheet.getLastColumn(), 8);
  var allRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  var antigua = String(allRows[0][0] || '').toLowerCase().indexOf('albar') >= 0;
  var raw = allRows.slice(1);
  var lastFecha = '', lastRef = '', lastCli = '';
  raw.forEach(function(r) {
    var fecha, ref, pzs, pal, tot, cli;
    if (antigua) {
      fecha = r[2]; ref = String(r[3] || '').trim();
      pzs = Number(r[4]) || 0; pal = Number(r[5]) || 0;
      tot = Number(r[6]) || 0; cli = String(r[7] || '').trim().toUpperCase();
    } else {
      fecha = r[0]; ref = String(r[1] || '').trim();
      pzs = Number(r[2]) || 0; pal = Number(r[3]) || 0;
      tot = Number(r[4]) || 0; cli = String(r[5] || '').trim().toUpperCase();
    }
    if (fecha instanceof Date || (fecha && String(fecha).trim())) lastFecha = fecha;
    if (ref) lastRef = ref;
    if (cli) lastCli = cli;
    if (lastRef && pal && lastCli) {
      result.push({ albaran: '', cmr: '',
                    fecha: lastFecha, referencia: lastRef, piezasPal: pzs,
                    palets: pal, totPiezas: tot || pzs * pal, cliente: lastCli });
    }
  });
  return result;
}

// ── Consolida filas con misma clave sumando palets y totPiezas ────────────────
function _consolidarFilas(rows, keyFn) {
  var seen = {}, order = [];
  rows.forEach(function(p) {
    var k = keyFn(p);
    if (seen[k]) {
      seen[k].palets    += p.palets;
      seen[k].totPiezas += p.totPiezas;
    } else {
      seen[k] = { albaran: p.albaran, cmr: p.cmr || '', fecha: p.fecha,
                  referencia: p.referencia, piezasPal: p.piezasPal,
                  palets: p.palets, totPiezas: p.totPiezas, cliente: p.cliente };
      order.push(k);
    }
  });
  return order.map(function(k) { return seen[k]; });
}

// ── Añade nuevas filas a ENTRADAS/SALIDAS en una sola escritura batch ─────────
function _appendToMovimientos(ss, tipo, newRows) {
  if (!newRows || !newRows.length) return;
  const sheetName = tipo === 'ENTRADA' ? 'ENTRADAS' : 'SALIDAS';
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const lastRow = Math.max(2, sheet.getLastRow());

  const keyEntrada = function(p) { return p.albaran + '||' + p.referencia + '||' + p.piezasPal + '||' + p.cliente; };
  const keySalida  = function(p) { return _dateToMs(p.fecha) + '||' + p.referencia + '||' + p.piezasPal + '||' + p.cliente; };
  const filas = _consolidarFilas(newRows, tipo === 'ENTRADA' ? keyEntrada : keySalida);

  if (tipo === 'ENTRADA') {
    const data = filas.map(p =>
      [p.albaran, p.cmr || '', _toDate(p.fecha), p.referencia, p.piezasPal, p.palets, p.totPiezas, p.cliente]);
    const rng = sheet.getRange(lastRow + 1, 1, data.length, 8);
    rng.setValues(data);
    sheet.getRange(lastRow + 1, 3, data.length, 1).setNumberFormat('dd/mm/yyyy');
  } else {
    const data = filas.map(p =>
      [_toDate(p.fecha), p.referencia, p.piezasPal, p.palets, p.totPiezas, p.cliente]);
    const rng = sheet.getRange(lastRow + 1, 1, data.length, 6);
    rng.setValues(data);
    sheet.getRange(lastRow + 1, 1, data.length, 1).setNumberFormat('dd/mm/yyyy');
  }
}

// ── Lee la hoja STOCK BASE (inventario físico puntual) ────────────────────────
// Estructura: Cliente | Referencia | Piezas/pal | Palets | Fecha  (5 cols)
function _leerStockBase(ss) {
  var sheet = ss.getSheetByName('STOCK BASE');
  if (!sheet || sheet.getLastRow() < 3) return { baseMap: {}, fechaCorteMs: 0 };
  var raw = sheet.getRange(3, 1, sheet.getLastRow() - 2, 5).getValues();
  var baseMap = {};
  var fechaCorteMs = Infinity;
  raw.forEach(function(r) {
    var cli = String(r[0] || '').trim().toUpperCase();
    var ref = String(r[1] || '').trim();
    var pzs = Number(r[2]) || 0;
    var pal = Number(r[3]) || 0;
    if (!cli || !ref) return;
    var key = cli + '||' + ref + '||' + pzs;
    baseMap[key] = { cli: cli, ref: ref, piezasPal: pzs, stockPal: pal };
    var fMs = _dateToMs(r[4]);
    if (fMs && fMs < fechaCorteMs) fechaCorteMs = fMs;
  });
  if (fechaCorteMs === Infinity) fechaCorteMs = 0;
  return { baseMap: baseMap, fechaCorteMs: fechaCorteMs };
}

// ── Recalcula STOCK leyendo ENTRADAS/SALIDAS ──────────────────────────────────
function actualizarStock(ss, sincronizar) {
  if (!ss) ss = getSS();
  if (sincronizar === undefined) sincronizar = true;

  const stockSheet = ss.getSheetByName('STOCK');
  if (!stockSheet) { Logger.log('Hoja STOCK no encontrada.'); return; }

  const entSheet = ss.getSheetByName('ENTRADAS');
  const salSheet = ss.getSheetByName('SALIDAS');

  function _restaurarFiltros() {
    var toRestaurar = sincronizar ? ['STOCK', 'ENTRADAS', 'SALIDAS'] : ['STOCK'];
    toRestaurar.forEach(function(nombre) {
      var sh = ss.getSheetByName(nombre);
      if (!sh || sh.getLastRow() < 2) return;
      if (sh.getFilter()) sh.getFilter().remove();
      sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), sh.getLastColumn()).createFilter();
    });
  }

  var entRows, salRows;
  if (sincronizar) {
    ['ENTRADAS', 'SALIDAS', 'STOCK'].forEach(function(nombre) {
      var sh = ss.getSheetByName(nombre);
      if (sh && sh.getFilter()) sh.getFilter().remove();
    });
    entRows = entSheet ? _rebuildEntradas(entSheet, _leerFilasEntradas(entSheet)) : [];
    salRows = salSheet ? _rebuildSalidas(salSheet, _leerFilasSalidas(salSheet)) : [];
  } else {
    var sf = stockSheet.getFilter();
    if (sf) sf.remove();
    entRows = entSheet ? _leerFilasEntradas(entSheet) : [];
    salRows = salSheet ? _leerFilasSalidas(salSheet) : [];
  }

  const stBase = _leerStockBase(ss);
  const baseMap = stBase.baseMap;
  const fechaCorteMs = stBase.fechaCorteMs;

  const map = {};
  Object.keys(baseMap).forEach(function(key) {
    var b = baseMap[key];
    map[key] = { cli: b.cli, ref: b.ref, piezasPal: b.piezasPal,
                 palEnt: b.stockPal, palSal: 0,
                 piezEnt: b.stockPal * b.piezasPal, piezSal: 0 };
  });

  entRows.forEach(function(p) {
    const key = p.cliente + '||' + p.referencia + '||' + p.piezasPal;
    // Solo filtrar por fecha de corte si la referencia está en STOCK BASE
    // (para las que no están, contamos todos sus movimientos desde el principio)
    if (fechaCorteMs && baseMap[key] && _dateToMs(p.fecha) < fechaCorteMs) return;
    if (!map[key]) map[key] = { cli: p.cliente, ref: p.referencia, piezasPal: p.piezasPal, palEnt: 0, palSal: 0, piezEnt: 0, piezSal: 0 };
    map[key].palEnt  += p.palets;
    map[key].piezEnt += p.totPiezas || p.piezasPal * p.palets;
  });
  salRows.forEach(function(p) {
    const key = p.cliente + '||' + p.referencia + '||' + p.piezasPal;
    if (fechaCorteMs && baseMap[key] && _dateToMs(p.fecha) < fechaCorteMs) return;
    if (!map[key]) map[key] = { cli: p.cliente, ref: p.referencia, piezasPal: p.piezasPal, palEnt: 0, palSal: 0, piezEnt: 0, piezSal: 0 };
    map[key].palSal  += p.palets;
    map[key].piezSal += p.totPiezas || p.piezasPal * p.palets;
  });

  const rows = Object.values(map)
    .sort((a, b) => a.cli.localeCompare(b.cli) || a.ref.localeCompare(b.ref))
    .map(function(d) {
      const sp = d.palEnt - d.palSal;
      return [d.cli, d.ref, d.piezasPal, sp, d.piezEnt - d.piezSal,
              sp > 0 ? '🟢 OK' : (sp === 0 ? '🟡 AGOTADO' : '🔴 NEGATIVO')];
    });

  const nOldCols = Math.max(6, stockSheet.getLastColumn());
  const lastRow = stockSheet.getLastRow();
  if (lastRow >= 3) stockSheet.getRange(3, 1, lastRow - 2, nOldCols).breakApart().clear();

  stockSheet.getRange(2, 1, 1, 6)
    .setValues([['👤 Cliente','📋 Referencia','🔢 Piezas/palet','📦 Stock palets','➕ Stock piezas','🚦 Estado']])
    .setBackground('#2E75B6').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  if (nOldCols > 6) stockSheet.getRange(2, 7, 1, nOldCols - 6).clearContent().clearFormat();

  if (!rows.length) {
    stockSheet.getRange(3, 1, 1, 6).merge()
      .setValue('Sin movimientos. Rellena el formulario para añadir entradas/salidas.')
      .setFontStyle('italic').setFontColor('#888888').setHorizontalAlignment('center');
    _restaurarFiltros();
    return;
  }

  const n = rows.length;
  const stockRng = stockSheet.getRange(3, 1, n, 6);
  stockRng.setValues(rows);
  stockRng.setHorizontalAlignment('center').setVerticalAlignment('middle');
  stockSheet.getRange(3, 3, n, 3).setNumberFormat('#,##0');

  const bgGrid = [], fgGrid = [], fwGrid = [];
  rows.forEach(function(r, i) {
    const sp    = r[3];
    const rowBg = i % 2 === 0 ? '#EEF4FB' : '#FFFFFF';
    var spBg, spFg;
    if      (sp < 0)  { spBg = '#F4CCCC'; spFg = '#CC0000'; }
    else if (sp === 0){ spBg = '#FFF9C4'; spFg = '#7D6608'; }
    else               { spBg = '#D9EAD3'; spFg = '#274E13'; }
    bgGrid.push([rowBg, rowBg, rowBg, spBg, spBg, spBg]);
    fgGrid.push(['#000000','#000000','#000000', spFg, spFg, '#000000']);
    fwGrid.push(['normal', 'normal', 'normal', 'bold', 'normal', 'normal']);
  });
  stockRng.setBackgrounds(bgGrid);
  stockRng.setFontColors(fgGrid);
  stockRng.setFontWeights(fwGrid);

  _aplicarBordes(stockSheet, rows, 2, n + 1, 6, function(r) { return r[0]; });
  _mergeGroups(stockSheet, rows, 3, function(r) { return r[0]; },              [1]);
  _mergeGroups(stockSheet, rows, 3, function(r) { return r[0] + '||' + r[1]; }, [2]);

  _restaurarFiltros();
  Logger.log('Stock actualizado: ' + n + ' líneas.');
}

// ── ENTRADAS: rebuild, devuelve las filas procesadas ─────────────────────────
// Estructura: alb(0)|cmr(1)|fecha(2)|ref(3)|pzs(4)|pal(5)|tot(6)|cli(7)
function _rebuildEntradas(sheet, rows) {
  if (!sheet) return [];
  const nOldCols = Math.max(8, sheet.getLastColumn());
  const lastRow = sheet.getLastRow();
  if (lastRow >= 3) sheet.getRange(3, 1, lastRow - 2, nOldCols).breakApart().clear();

  sheet.getRange(2, 1, 1, 8)
    .setValues([['🔖 Nº Albarán','🚢 Carta Porte','📅 Fecha','📋 Referencia','🔢 Piezas/palet','📦 Nº Palets','➕ Total piezas','👤 Cliente']])
    .setBackground('#2E7D32').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  if (nOldCols > 8) sheet.getRange(2, 9, 1, nOldCols - 8).clearContent().clearFormat();

  if (!rows.length) return [];

  rows = _consolidarFilas(rows, function(p) {
    return p.albaran + '||' + p.referencia + '||' + p.piezasPal + '||' + p.cliente;
  });

  rows.sort(function(a, b) {
    var fd = _dateToMs(a.fecha) - _dateToMs(b.fecha);
    if (fd) return fd;
    var na = parseInt(a.albaran, 10), nb = parseInt(b.albaran, 10);
    var ad = (!isNaN(na) && !isNaN(nb)) ? na - nb : String(a.albaran).localeCompare(String(b.albaran));
    if (ad) return ad;
    return String(a.cliente).localeCompare(String(b.cliente));
  });

  const data = rows.map(function(p) {
    return [p.albaran, p.cmr || '', _toDate(p.fecha), p.referencia, p.piezasPal, p.palets, p.totPiezas, p.cliente];
  });
  const n = data.length;
  const rng = sheet.getRange(3, 1, n, 8);

  rng.setValues(data);
  rng.setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRangeList([_a1(3,1,n,2), _a1(3,4,n,1), _a1(3,8,n,1)]).setNumberFormat('@');
  sheet.getRange(3, 3, n, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(3, 5, n, 3).setNumberFormat('#,##0');

  var grupIdx = -1, prevAlb = null;
  const bgGrid = data.map(function(r) {
    if (r[0] !== prevAlb) { prevAlb = r[0]; grupIdx++; }
    var bg = grupIdx % 2 === 0 ? '#F5F5F5' : '#FFFFFF';
    return [bg, bg, bg, bg, bg, bg, bg, bg];
  });
  rng.setBackgrounds(bgGrid);

  _aplicarBordes(sheet, data, 2, n + 1, 8, function(r) { return r[0]; });
  _mergeGroups(sheet, data, 3, function(r) { return r[0]; },               [1, 2, 3]);
  _mergeGroups(sheet, data, 3, function(r) { return r[0] + '||' + r[3]; }, [4]);
  _mergeGroups(sheet, data, 3, function(r) { return r[0] + '||' + r[7]; }, [8]);

  return rows;
}

// ── SALIDAS: rebuild, devuelve las filas procesadas ──────────────────────────
// Estructura: fecha(0)|ref(1)|pzs(2)|pal(3)|tot(4)|cli(5)
function _rebuildSalidas(sheet, rows) {
  if (!sheet) return [];
  const nOldCols = Math.max(6, sheet.getLastColumn());
  const lastRow = sheet.getLastRow();
  if (lastRow >= 3) sheet.getRange(3, 1, lastRow - 2, nOldCols).breakApart().clear();

  sheet.getRange(2, 1, 1, 6)
    .setValues([['📅 Fecha','📋 Referencia','🔢 Piezas/palet','📦 Nº Palets','➕ Total piezas','👤 Cliente']])
    .setBackground('#C62828').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  if (nOldCols > 6) sheet.getRange(2, 7, 1, nOldCols - 6).clearContent().clearFormat();

  if (!rows.length) return [];

  rows = _consolidarFilas(rows, function(p) {
    return _dateToMs(p.fecha) + '||' + p.referencia + '||' + p.piezasPal + '||' + p.cliente;
  });

  rows.sort(function(a, b) {
    return _dateToMs(a.fecha) - _dateToMs(b.fecha)
        || a.cliente.localeCompare(b.cliente)
        || String(a.referencia).localeCompare(String(b.referencia));
  });

  const data = rows.map(function(p) {
    return [_toDate(p.fecha), p.referencia, p.piezasPal, p.palets, p.totPiezas, p.cliente];
  });
  const n = data.length;
  const rng = sheet.getRange(3, 1, n, 6);

  rng.setValues(data);
  rng.setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRangeList([_a1(3,2,n,1), _a1(3,6,n,1)]).setNumberFormat('@');
  sheet.getRange(3, 1, n, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(3, 3, n, 3).setNumberFormat('#,##0');

  var grupIdx = -1, prevKey = null;
  const bgGrid = data.map(function(r) {
    var k = _dateToMs(r[0]) + '||' + r[5];
    if (k !== prevKey) { prevKey = k; grupIdx++; }
    var bg = grupIdx % 2 === 0 ? '#F5F5F5' : '#FFFFFF';
    return [bg, bg, bg, bg, bg, bg];
  });
  rng.setBackgrounds(bgGrid);

  _aplicarBordes(sheet, data, 2, n + 1, 6, function(r) { return _dateToMs(r[0]) + '||' + r[5]; });
  _mergeGroups(sheet, data, 3, function(r) { return _dateToMs(r[0]); },                          [1]);
  _mergeGroups(sheet, data, 3, function(r) { return _dateToMs(r[0]) + '||' + r[5] + '||' + r[1]; }, [2]);
  _mergeGroups(sheet, data, 3, function(r) { return _dateToMs(r[0]) + '||' + r[5]; },             [6]);

  return rows;
}

// ─── Helper: bordes de tabla con separadores de grupo ────────────────────────
function _aplicarBordes(sheet, data, headerRow, nRows, nCols, groupKeyFn) {
  var rng = sheet.getRange(headerRow, 1, nRows, nCols);
  rng.setBorder(true, true, true, true, true, true,  '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  rng.setBorder(true, true, true, true, null, null,  '#555555', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  if (groupKeyFn && data && data.length > 1) {
    var dataStart = headerRow + 1;
    var a1s = [];
    for (var i = 0; i < data.length - 1; i++) {
      if (groupKeyFn(data[i]) !== groupKeyFn(data[i + 1])) {
        a1s.push(_a1(dataStart + i, 1, 1, nCols));
      }
    }
    if (a1s.length) {
      sheet.getRangeList(a1s)
        .setBorder(null, null, true, null, null, null, '#555555', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }
}

// ─── Helper: combina celdas consecutivas con mismo valor de clave ─────────────
function _mergeGroups(sheet, data, startRow, keyFn, cols) {
  if (!data.length) return;
  var a1s = [];
  var gs = 0;
  for (var i = 1; i <= data.length; i++) {
    var key  = keyFn(data[gs]);
    var same = i < data.length && key && keyFn(data[i]) === key;
    if (!same) {
      if (i - gs > 1) {
        cols.forEach(function(c) {
          var a1 = _a1(startRow + gs, c, i - gs, 1);
          a1s.push(a1);
          sheet.getRange(a1).mergeVertically();
        });
      }
      gs = i;
    }
  }
  if (a1s.length) {
    sheet.getRangeList(a1s).setVerticalAlignment('middle').setHorizontalAlignment('center');
  }
}

// ─── Helpers de fecha ────────────────────────────────────────────────────────
var _MMAP = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};

function _dateToMs(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate()).getTime();
  const s = String(v || '').trim();
  if (!s) return 0;
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]).getTime();
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]).getTime();
  var dm = s.match(/([A-Z][a-z]{2})\s+(\d{1,2})\s+(\d{4})/);
  if (dm && _MMAP[dm[1]] !== undefined) return new Date(+dm[3], _MMAP[dm[1]], +dm[2]).getTime();
  return 0;
}

function _toDate(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const s = String(v || '').trim();
  if (!s) return v;
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  var dm = s.match(/([A-Z][a-z]{2})\s+(\d{1,2})\s+(\d{4})/);
  if (dm && _MMAP[dm[1]] !== undefined) return new Date(+dm[3], _MMAP[dm[1]], +dm[2]);
  return v;
}

// ─── Helper: construye notación A1 sin llamadas a la API ─────────────────────
function _a1(r, c, nr, nc) {
  function col(n) { var s=''; while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s; }
  return col(c)+r+':'+col(c+nc-1)+(r+nr-1);
}
