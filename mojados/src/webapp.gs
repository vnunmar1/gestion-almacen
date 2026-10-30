// ============================================================
// WEBAPP.GS — Web App: sirve el formulario HTML y procesa envíos
// ============================================================

function doGet(e) {
  if (e && e.parameter && e.parameter.manifest === '1') {
    return _serveManifest();
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('JNSL Almacén')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _serveManifest() {
  const url  = ScriptApp.getService().getUrl();
  const icon = "data:image/svg+xml," + encodeURIComponent(
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 192 192">' +
    '<rect width="192" height="192" rx="40" fill="#1A56A0"/>' +
    '<text x="96" y="126" font-family="sans-serif" font-size="86"' +
    ' font-weight="bold" text-anchor="middle" fill="white">JN</text>' +
    '</svg>'
  );
  const manifest = {
    name:             'JNSL Almacén',
    short_name:       'Almacén',
    description:      'Control de entradas y salidas · Julio Núñez S.L.',
    start_url:        url,
    scope:            url,
    display:          'standalone',
    orientation:      'portrait',
    background_color: '#ffffff',
    theme_color:      '#1A56A0',
    icons: [
      { src: icon, sizes: '192x192', type: 'image/svg+xml' },
      { src: icon, sizes: '512x512', type: 'image/svg+xml' }
    ]
  };
  return ContentService.createTextOutput(JSON.stringify(manifest))
    .setMimeType(ContentService.MimeType.JSON);
}

// Llamado desde el HTML via google.script.run.procesarMovimiento(data).
// data = { fecha, tipo, albaran, cmr, obs,
//          lineas: [{referencia, piezas, palets, cliente}, ...] }
function procesarMovimiento(data) {
  try {
    const ss = getSS();
    const { fecha, tipo, albaran, cmr, obs, lineas } = data;

    if (!lineas || !lineas.length) throw new Error('No hay líneas de referencia.');

    const nuevasFilas = [];
    for (const linea of lineas) {
      const totalPiezas = Number(linea.piezas) * Number(linea.palets);
      nuevasFilas.push({
        albaran:    String(albaran          || ''),
        cmr:        String(cmr              || ''),
        fecha:      String(fecha),
        referencia: String(linea.referencia || ''),
        piezasPal:  Number(linea.piezas),
        palets:     Number(linea.palets),
        totPiezas:  totalPiezas,
        cliente:    String(linea.cliente   || '').trim().toUpperCase()
      });
    }

    try {
      // Añade las nuevas filas a ENTRADAS/SALIDAS y recalcula solo el STOCK.
      // El rebuild completo (reordenar, reformatear) lo hace "Actualizar stock ahora" del menú.
      _appendToMovimientos(ss, String(tipo), nuevasFilas);
      actualizarStock(ss, false);
    } catch (stockErr) {
      Logger.log('Aviso stock: ' + stockErr + '\n' + stockErr.stack);
      return { ok: true, warn: 'Entrada guardada. Recalcula el stock manualmente.' };
    }

    const totalPalets = lineas.reduce((s, l) => s + Number(l.palets), 0);
    Logger.log('✅ ' + tipo + ': ' + albaran + '/' + cmr + ' · ' + lineas.length + ' líneas · ' + totalPalets + ' pal.');
    return { ok: true };

  } catch (err) {
    Logger.log('❌ procesarMovimiento: ' + err + '\n' + err.stack);
    return { ok: false, error: err.toString() };
  }
}
