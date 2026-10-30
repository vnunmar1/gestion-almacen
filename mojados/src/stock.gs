// ============================================================
// STOCK.GS — Re-ordenar ENTRADAS/SALIDAS y calcular STOCK
// ============================================================
// ENTRADAS cols: Nº Albarán | Fecha | Referencia | Piezas/pal | NºPalets | TotPiezas | Cliente  (7)
// SALIDAS cols:  Nº Albarán | Carta Porte | Fecha | Referencia | Piezas/pal | NºPalets | TotPiezas | Cliente  (8)
// STOCK cols:    Cliente | Referencia | Piezas/pal | Stock Pal. | Stock Piezas | Estado  (6)

function _cleanAlbaran(v) {
  const s = String(v || '').trim();
  return /^\d{1,3}(\.\d{3})+$/.test(s) ? s.replace(/\./g, '') : s;
}

// ── Lectura de filas desde ENTRADAS con fill-forward ─────────────────────────
// Detecta formato por cabecera: col2 "Carta Porte" = nuevo (8 cols), otro = antiguo (7 cols)
function _leerFilasEntradas(sheet) {
  var result = [];
  if (!sheet || sheet.getLastRow() < 3) return result;
  var lastCol = Math.min(sheet.getLastColumn(), 8);
  // Cabecera + datos en una sola llamada API
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
// Detecta formato por cabecera: col1 "Albarán" = antiguo (8 cols), "Fecha" = nuevo (6 cols)
function _leerFilasSalidas(sheet) {
  var result = [];
  if (!sheet || sheet.getLastRow() < 3) return result;
  var lastCol = Math.min(sheet.getLastColumn(), 8);
  // Cabecera + datos en una sola llamada API
  var allRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  var antigua = String(allRows[0][0] || '').toLowerCase().indexOf('albar') >= 0;
  var raw = allRows.slice(1);
  var lastFecha = '', lastRef = '', lastCli = '';
  raw.forEach(function(r) {
    var fecha, ref, pzs, pal, tot, cli;
    if (antigua) {
      // antigua: alb(0)|cmr(1)|fecha(2)|ref(3)|pzs(4)|pal(5)|tot(6)|cli(7)
      fecha = r[2]; ref = String(r[3] || '').trim();
      pzs = Number(r[4]) || 0; pal = Number(r[5]) || 0;
      tot = Number(r[6]) || 0; cli = String(r[7] || '').trim().toUpperCase();
    } else {
      // nueva: fecha(0)|ref(1)|pzs(2)|pal(3)|tot(4)|cli(5)
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

// ── Actualización rápida del STOCK: aplica solo el delta del movimiento nuevo ──
// Usado desde la app. NO lee el historial completo de ENTRADAS/SALIDAS.
// El menú "Actualizar stock ahora" hace el recálculo completo desde cero.
function _aplicarDeltaStock(ss, tipo, nuevasFilas) {
  var stockSheet = ss.getSheetByName('STOCK');
  if (!stockSheet || stockSheet.getLastRow() < 3) {
    actualizarStock(ss, true);
    return;
  }

  var sf = stockSheet.getFilter();
  if (sf) sf.remove();

  var nRows = stockSheet.getLastRow() - 2;
  var nCols = 6;
  stockSheet.getRange(3, 1, nRows, nCols).breakApart();
  var data = stockSheet.getRange(3, 1, nRows, nCols).getValues();

  // Fill-forward Cliente y Referencia (pueden estar combinadas)
  var lastCli = '', lastRef = '';
  data.forEach(function(r) {
    if (r[0]) lastCli = String(r[0]).trim().toUpperCase(); else r[0] = lastCli;
    if (r[1]) lastRef = String(r[1]).trim();               else r[1] = lastRef;
  });

  // Índice: clave → posición en el array
  var idx = {};
  data.forEach(function(r, i) {
    idx[r[0] + '||' + r[1] + '||' + Number(r[2])] = i;
  });

  var sign = (tipo === 'ENTRADA') ? 1 : -1;
  nuevasFilas.forEach(function(p) {
    var cli = String(p.cliente    || '').trim().toUpperCase();
    var ref = String(p.referencia || '').trim();
    var pzs = Number(p.piezasPal) || 0;
    var key = cli + '||' + ref + '||' + pzs;
    var deltaPal = sign * (Number(p.palets)    || 0);
    var deltaPzs = sign * (Number(p.totPiezas) || pzs * Number(p.palets));

    if (idx[key] !== undefined) {
      data[idx[key]][3] = Number(data[idx[key]][3]) + deltaPal;
      data[idx[key]][4] = Number(data[idx[key]][4]) + deltaPzs;
    } else {
      data.push([cli, ref, pzs, deltaPal, deltaPzs, '']);
      idx[key] = data.length - 1;
    }
  });

  // Actualizar Estado y ordenar
  data.forEach(function(r) {
    var sp = Number(r[3]);
    r[5] = sp > 0 ? '🟢 OK' : (sp === 0 ? '🟡 AGOTADO' : '🔴 NEGATIVO');
  });
  data.sort(function(a, b) {
    return String(a[0]).localeCompare(String(b[0])) || String(a[1]).localeCompare(String(b[1]));
  });

  // Limpiar y escribir
  stockSheet.getRange(3, 1, Math.max(nRows + 10, 50), nCols).clearContent().clearFormat();
  var n = data.length;
  if (!n) return;

  var rng = stockSheet.getRange(3, 1, n, nCols);
  rng.setValues(data);
  rng.setHorizontalAlignment('center').setVerticalAlignment('middle');
  stockSheet.getRange(3, 3, n, 3).setNumberFormat('#,##0');

  var bgGrid = [], fgGrid = [], fwGrid = [];
  data.forEach(function(r, i) {
    var sp    = Number(r[3]);
    var rowBg = i % 2 === 0 ? '#EEF4FB' : '#FFFFFF';
    var spBg  = sp > 0 ? '#D9EAD3' : (sp === 0 ? '#FFF9C4' : '#F4CCCC');
    var spFg  = sp > 0 ? '#274E13' : (sp === 0 ? '#7D6608' : '#CC0000');
    bgGrid.push([rowBg, rowBg, rowBg, spBg, spBg, spBg]);
    fgGrid.push(['#000000','#000000','#000000', spFg, spFg, '#000000']);
    fwGrid.push(['normal', 'normal', 'normal', 'bold', 'normal', 'normal']);
  });
  rng.setBackgrounds(bgGrid);
  rng.setFontColors(fgGrid);
  rng.setFontWeights(fwGrid);

  _aplicarBordes(stockSheet, data, 2, n + 1, nCols, function(r) { return r[0]; });
  _mergeGroups(stockSheet, data, 3, function(r) { return r[0]; },              [1]);
  _mergeGroups(stockSheet, data, 3, function(r) { return r[0] + '||' + r[1]; }, [2]);

  if (stockSheet.getLastRow() >= 2)
    stockSheet.getRange(2, 1, Math.max(1, stockSheet.getLastRow() - 1), nCols).createFilter();
}

// ── Añade nuevas filas a ENTRADAS/SALIDAS en una sola escritura batch ─────────
// _aplicarDeltaStock se encargará de actualizar el STOCK.
function _appendToMovimientos(ss, tipo, newRows) {
  if (!newRows || !newRows.length) return;
  const sheetName = tipo === 'ENTRADA' ? 'ENTRADAS' : 'SALIDAS';
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const lastRow = Math.max(2, sheet.getLastRow());
  if (tipo === 'ENTRADA') {
    // ENTRADAS: alb | cmr | fecha | ref | pzs | pal | tot | cli  (8 cols)
    const data = newRows.map(p =>
      [p.albaran, p.cmr || '', _toDate(p.fecha), p.referencia, p.piezasPal, p.palets, p.totPiezas, p.cliente]);
    const rng = sheet.getRange(lastRow + 1, 1, data.length, 8);
    rng.setValues(data);
    sheet.getRange(lastRow + 1, 3, data.length, 1).setNumberFormat('dd/mm/yyyy'); // fecha col 3
  } else {
    // SALIDAS: fecha | ref | pzs | pal | tot | cli  (6 cols, sin alb ni cmr)
    const data = newRows.map(p =>
      [_toDate(p.fecha), p.referencia, p.piezasPal, p.palets, p.totPiezas, p.cliente]);
    const rng = sheet.getRange(lastRow + 1, 1, data.length, 6);
    rng.setValues(data);
    sheet.getRange(lastRow + 1, 1, data.length, 1).setNumberFormat('dd/mm/yyyy'); // fecha col 1
  }
}

// ── Lee la hoja STOCK BASE (inventario físico puntual) ────────────────────────
// Devuelve: { baseMap: { key → {cli,ref,piezasPal,stockPal} }, fechaCorteMs }
// Si no hay hoja STOCK BASE, devuelve baseMap vacío y fechaCorteMs=0 (sin filtro).
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
// sincronizar=true (default, desde menú): re-ordena y reformatea ENTRADAS/SALIDAS
// sincronizar=false (desde app): solo recalcula STOCK sin tocar las hojas de movimientos
function actualizarStock(ss, sincronizar) {
  if (!ss) ss = getSS();
  if (sincronizar === undefined) sincronizar = true;

  const stockSheet = ss.getSheetByName('STOCK');
  if (!stockSheet) { Logger.log('Hoja STOCK no encontrada.'); return; }

  const entSheet = ss.getSheetByName('ENTRADAS');
  const salSheet = ss.getSheetByName('SALIDAS');

  var filtros = {};
  function _restaurarFiltros() {
    var toRestaurar = sincronizar ? ['STOCK', 'ENTRADAS', 'SALIDAS'] : ['STOCK'];
    toRestaurar.forEach(function(nombre) {
      var sh = ss.getSheetByName(nombre);
      if (!sh || sh.getLastRow() < 2) return;
      if (sh.getFilter()) sh.getFilter().remove();
      sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), sh.getLastColumn()).createFilter();
    });
    if (sincronizar) {
      ss.getSheets().forEach(function(sh) {
        var n = sh.getName();
        if (filtros[n] && ['STOCK','ENTRADAS','SALIDAS'].indexOf(n) < 0
            && sh.getLastRow() >= 2 && !sh.getFilter()) {
          sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), sh.getLastColumn()).createFilter();
        }
      });
    }
  }

  var entRows, salRows;
  if (sincronizar) {
    // Quitar filtros de todas las hojas (necesario para mergear celdas en rebuild)
    ss.getSheets().forEach(function(sh) {
      if (sh.getFilter()) { filtros[sh.getName()] = true; sh.getFilter().remove(); }
    });
    // Rebuild re-ordena y reformatea las hojas de movimientos
    entRows = entSheet ? _rebuildEntradas(entSheet, _leerFilasEntradas(entSheet)) : [];
    salRows = salSheet ? _rebuildSalidas(salSheet, _leerFilasSalidas(salSheet)) : [];
  } else {
    // Solo quitar filtro de STOCK (necesario para mergear celdas en STOCK)
    var sf = stockSheet.getFilter();
    if (sf) sf.remove();
    // Leer datos directamente sin modificar las hojas de movimientos
    entRows = entSheet ? _leerFilasEntradas(entSheet) : [];
    salRows = salSheet ? _leerFilasSalidas(salSheet) : [];
  }

  // Leer base de inventario (si existe)
  const stBase = _leerStockBase(ss);
  const baseMap = stBase.baseMap;
  const fechaCorteMs = stBase.fechaCorteMs;

  // Inicializar mapa desde la base de inventario (palEnt = stock base, palSal = 0)
  const map = {};
  Object.keys(baseMap).forEach(function(key) {
    var b = baseMap[key];
    map[key] = { cli: b.cli, ref: b.ref, piezasPal: b.piezasPal,
                 palEnt: b.stockPal, palSal: 0,
                 piezEnt: b.stockPal * b.piezasPal, piezSal: 0 };
  });

  // Acumular ENTRADAS (solo >= fechaCorte si hay base)
  entRows.forEach(function(p) {
    if (fechaCorteMs && _dateToMs(p.fecha) < fechaCorteMs) return;
    const key = p.cliente + '||' + p.referencia + '||' + p.piezasPal;
    if (!map[key]) map[key] = { cli: p.cliente, ref: p.referencia, piezasPal: p.piezasPal, palEnt: 0, palSal: 0, piezEnt: 0, piezSal: 0 };
    map[key].palEnt  += p.palets;
    map[key].piezEnt += p.totPiezas || p.piezasPal * p.palets;
  });
  // Acumular SALIDAS (solo >= fechaCorte si hay base)
  salRows.forEach(function(p) {
    if (fechaCorteMs && _dateToMs(p.fecha) < fechaCorteMs) return;
    const key = p.cliente + '||' + p.referencia + '||' + p.piezasPal;
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

  // Limpiar filas antiguas (compatible con hoja de 8 cols de versión anterior)
  const nOldCols = Math.max(6, stockSheet.getLastColumn());
  const lastRow = stockSheet.getLastRow();
  if (lastRow >= 3) stockSheet.getRange(3, 1, lastRow - 2, nOldCols).breakApart().clear();

  // Actualizar cabecera a 6 cols (quita Pal.Ent/Pal.Sal de versiones anteriores)
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
  stockSheet.getRange(3, 3, n, 3).setNumberFormat('#,##0'); // pzs/pal, stock pal, stock piezas

  // Formateo por fila → arrays 2D → 3 llamadas API en lugar de N×4
  const bgGrid = [], fgGrid = [], fwGrid = [];
  rows.forEach(function(r, i) {
    const sp    = r[3]; // índice 3 = Stock Palets
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
// Estructura: alb(0) | cmr(1) | fecha(2) | ref(3) | pzs(4) | pal(5) | tot(6) | cli(7)
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
  sheet.getRange(3, 1, n, 2).setNumberFormat('@');           // alb, cmr — texto
  sheet.getRange(3, 3, n, 1).setNumberFormat('dd/mm/yyyy'); // col 3 = Fecha
  sheet.getRange(3, 4, n, 1).setNumberFormat('@');           // ref — texto
  sheet.getRange(3, 5, n, 3).setNumberFormat('#,##0');       // cols 5-7 = pzs, pal, tot
  sheet.getRange(3, 8, n, 1).setNumberFormat('@');           // cli — texto

  var grupIdx = -1, prevAlb = null;
  const bgGrid = data.map(function(r) {
    if (r[0] !== prevAlb) { prevAlb = r[0]; grupIdx++; }
    var bg = grupIdx % 2 === 0 ? '#F5F5F5' : '#FFFFFF';
    return [bg, bg, bg, bg, bg, bg, bg, bg];
  });
  rng.setBackgrounds(bgGrid);

  _aplicarBordes(sheet, data, 2, n + 1, 8, function(r) { return r[0]; });
  _mergeGroups(sheet, data, 3, function(r) { return r[0]; },               [1, 2, 3]);    // alb, cmr, fecha
  _mergeGroups(sheet, data, 3, function(r) { return r[0] + '||' + r[3]; }, [4]);          // ref
  _mergeGroups(sheet, data, 3, function(r) { return r[0] + '||' + r[7]; }, [8]);          // cli: fusiona solo si mismo alb Y mismo cliente

  return rows;
}

// ── SALIDAS: rebuild, devuelve las filas procesadas ──────────────────────────
// Estructura: fecha(0) | ref(1) | pzs(2) | pal(3) | tot(4) | cli(5)
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
  sheet.getRange(3, 1, n, 1).setNumberFormat('dd/mm/yyyy'); // col 1 = Fecha
  sheet.getRange(3, 2, n, 1).setNumberFormat('@');           // ref — texto
  sheet.getRange(3, 3, n, 3).setNumberFormat('#,##0');       // cols 3-5 = pzs, pal, tot
  sheet.getRange(3, 6, n, 1).setNumberFormat('@');           // cli — texto

  var grupIdx = -1, prevKey = null;
  const bgGrid = data.map(function(r) {
    var k = _dateToMs(r[0]) + '||' + r[5];
    if (k !== prevKey) { prevKey = k; grupIdx++; }
    var bg = grupIdx % 2 === 0 ? '#F5F5F5' : '#FFFFFF';
    return [bg, bg, bg, bg, bg, bg];
  });
  rng.setBackgrounds(bgGrid);

  _aplicarBordes(sheet, data, 2, n + 1, 6, function(r) { return _dateToMs(r[0]) + '||' + r[5]; });
  _mergeGroups(sheet, data, 3, function(r) { return _dateToMs(r[0]); },                          [1]); // fecha
  _mergeGroups(sheet, data, 3, function(r) { return _dateToMs(r[0]) + '||' + r[5] + '||' + r[1]; }, [2]); // ref
  _mergeGroups(sheet, data, 3, function(r) { return _dateToMs(r[0]) + '||' + r[5]; },             [6]); // cli

  return rows;
}

// ─── Helper: bordes de tabla con separadores de grupo ────────────────────────
function _aplicarBordes(sheet, data, headerRow, nRows, nCols, groupKeyFn) {
  var rng = sheet.getRange(headerRow, 1, nRows, nCols);
  rng.setBorder(true, true, true, true, true, true,  '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  rng.setBorder(true, true, true, true, null, null,  '#555555', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  if (groupKeyFn && data && data.length > 1) {
    var dataStart = headerRow + 1;
    for (var i = 0; i < data.length - 1; i++) {
      if (groupKeyFn(data[i]) !== groupKeyFn(data[i + 1])) {
        sheet.getRange(dataStart + i, 1, 1, nCols)
          .setBorder(null, null, true, null, null, null, '#555555', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
    }
  }
}

// ─── Helper: combina celdas consecutivas con mismo valor de clave ─────────────
function _mergeGroups(sheet, data, startRow, keyFn, cols) {
  if (!data.length) return;
  var gs = 0;
  for (var i = 1; i <= data.length; i++) {
    var key  = keyFn(data[gs]);
    var same = i < data.length && key && keyFn(data[i]) === key;
    if (!same) {
      if (i - gs > 1) {
        cols.forEach(function(c) {
          sheet.getRange(startRow + gs, c, i - gs, 1)
            .mergeVertically().setVerticalAlignment('middle').setHorizontalAlignment('center');
        });
      }
      gs = i;
    }
  }
}

// ─── Helpers de fecha ────────────────────────────────────────────────────────

// Mapa de meses en inglés (formato Date.toString() de V8/GAS)
var _MMAP = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};

function _dateToMs(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate()).getTime();
  const s = String(v || '').trim();
  if (!s) return 0;
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]).getTime();
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]).getTime();
  // "Wed Jun 02 2026 00:00:00 GMT+0200 (hora de Europa central de verano)"
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
  // "Wed Jun 02 2026 00:00:00 GMT+0200 (hora de Europa central de verano)"
  var dm = s.match(/([A-Z][a-z]{2})\s+(\d{1,2})\s+(\d{4})/);
  if (dm && _MMAP[dm[1]] !== undefined) return new Date(+dm[3], _MMAP[dm[1]], +dm[2]);
  return v;
}
