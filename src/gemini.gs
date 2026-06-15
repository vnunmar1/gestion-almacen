// ============================================================
// GEMINI.GS — Extracción de datos con Gemini Vision
// Soporta dos documentos distintos: albarán y carta de porte
// ============================================================

// ─── Prompt por tipo de documento ────────────────────────────

const PROMPTS = {
  albaran:
    'Eres un asistente de control de almacén logístico. ' +
    'Analiza esta imagen de un ALBARÁN DE ENTREGA y extrae los datos. ' +
    'Responde ÚNICAMENTE con JSON válido, sin texto adicional ni bloques markdown. ' +
    'Estructura exacta:\n' +
    '{\n' +
    '  "tipo": "ENTRADA" o "SALIDA" (infiere por contexto),\n' +
    '  "fecha": "DD/MM/YYYY",\n' +
    '  "cliente": "nombre completo del cliente o empresa",\n' +
    '  "lineas": [\n' +
    '    { "referencia": "código del producto", "descripcion": "descripción breve", "palets": número entero }\n' +
    '  ],\n' +
    '  "observaciones": "notas relevantes o null"\n' +
    '}\n' +
    'El campo "lineas" debe tener UNA entrada por cada línea de producto del albarán. ' +
    'Si un campo no es legible escribe null. Para palets, si aparece "bultos" o "unidades" úsalo.',

  porte:
    'Eres un asistente de control de almacén logístico. ' +
    'Analiza esta imagen de una CARTA DE PORTE y extrae los datos. ' +
    'Responde ÚNICAMENTE con JSON válido, sin texto adicional ni bloques markdown. ' +
    'Estructura exacta:\n' +
    '{\n' +
    '  "orden_carga": "número de orden de carga o expedición",\n' +
    '  "matricula": "matrícula del vehículo (cabeza tractora)",\n' +
    '  "matricula_remolque": "matrícula del remolque o null",\n' +
    '  "transportista": "nombre de la empresa transportista o null",\n' +
    '  "fecha": "DD/MM/YYYY o null",\n' +
    '  "cliente": "destinatario o remitente si aparece, o null",\n' +
    '  "observaciones": "notas relevantes o null"\n' +
    '}\n' +
    'Si un campo no es legible o no aparece escribe null.',
};

// ─── Función principal de extracción ─────────────────────────

function extractDataFromPhoto(fileId, docType) {
  // docType: 'albaran' | 'porte'
  try {
    const config = getConfig();
    const prompt = PROMPTS[docType] || PROMPTS.albaran;

    // 1. Obtener URL de descarga desde Telegram
    const fileRes = UrlFetchApp.fetch(
      `https://api.telegram.org/bot${config.TELEGRAM_BOT_TOKEN}/getFile?file_id=${fileId}`
    );
    const fileMeta = JSON.parse(fileRes.getContentText());
    if (!fileMeta.ok) throw new Error('getFile falló: ' + JSON.stringify(fileMeta));
    const filePath = fileMeta.result.file_path;
    const imgUrl   = `https://api.telegram.org/file/bot${config.TELEGRAM_BOT_TOKEN}/${filePath}`;

    // 2. Descargar imagen y convertir a base64
    const imgResponse = UrlFetchApp.fetch(imgUrl);
    const imgBlob     = imgResponse.getBlob();
    const base64      = Utilities.base64Encode(imgBlob.getBytes());
    const mimeType    = imgBlob.getContentType() || 'image/jpeg';

    // 3. Llamar a Gemini 1.5 Flash Vision
    const geminiUrl =
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${config.GEMINI_API_KEY}`;

    const payload = {
      contents: [{
        parts: [
          { inline_data: { mime_type: mimeType, data: base64 } },
          { text: prompt }
        ]
      }],
      generationConfig: {
        temperature: 0.1,
        maxOutputTokens: 1024,
      }
    };

    const res = UrlFetchApp.fetch(geminiUrl, {
      method:      'post',
      contentType: 'application/json',
      payload:     JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    if (res.getResponseCode() !== 200) {
      Logger.log(`Error Gemini API (${res.getResponseCode()}): ` + res.getContentText());
      return null;
    }

    const raw  = JSON.parse(res.getContentText());
    const text = raw.candidates[0].content.parts[0].text;
    const clean = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();

    const parsed = JSON.parse(clean);
    Logger.log(`[Gemini ${docType}] extraído: ` + JSON.stringify(parsed));

    // Normalizar tipo en albarán
    if (docType === 'albaran' && parsed.tipo) {
      parsed.tipo = parsed.tipo.toUpperCase().includes('SALIDA') ? 'SALIDA' : 'ENTRADA';
    }

    // Normalizar matrículas a mayúsculas
    if (parsed.matricula)         parsed.matricula         = String(parsed.matricula).toUpperCase();
    if (parsed.matricula_remolque) parsed.matricula_remolque = String(parsed.matricula_remolque).toUpperCase();

    // Normalizar palets en líneas a entero
    if (docType === 'albaran' && Array.isArray(parsed.lineas)) {
      parsed.lineas = parsed.lineas.map(l => ({
        ...l,
        palets: parseInt(l.palets) || 0,
      }));
    }

    return parsed;

  } catch (err) {
    Logger.log(`Error en extractDataFromPhoto (${docType}): ` + err + '\n' + err.stack);
    return null;
  }
}
