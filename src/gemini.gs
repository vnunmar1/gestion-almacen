// ============================================================
// GEMINI.GS — Extracción de datos de imagen con Gemini Vision
// ============================================================

function extractDataFromPhoto(fileId) {
  try {
    const config = getConfig();

    // 1. Obtener ruta del archivo desde Telegram
    const fileRes = UrlFetchApp.fetch(
      `https://api.telegram.org/bot${config.TELEGRAM_BOT_TOKEN}/getFile?file_id=${fileId}`
    );
    const filePath = JSON.parse(fileRes.getContentText()).result.file_path;
    const imgUrl   = `https://api.telegram.org/file/bot${config.TELEGRAM_BOT_TOKEN}/${filePath}`;

    // 2. Descargar imagen y convertir a base64
    const imgBlob  = UrlFetchApp.fetch(imgUrl).getBlob();
    const base64   = Utilities.base64Encode(imgBlob.getBytes());
    const mimeType = imgBlob.getContentType() || 'image/jpeg';

    // 3. Prompt estructurado para Gemini Vision
    const prompt =
      'Eres un asistente de control de almacén logístico. ' +
      'Analiza esta imagen de un albarán de entrega o carta de porte y extrae los datos clave. ' +
      'Responde ÚNICAMENTE con un objeto JSON válido, sin texto adicional, sin bloques de código markdown. ' +
      'El JSON debe tener exactamente esta estructura:\n' +
      '{\n' +
      '  "tipo": "ENTRADA" o "SALIDA" (infiere por el contexto del documento),\n' +
      '  "referencia": "código alfanumérico del producto/artículo",\n' +
      '  "cliente": "nombre completo del cliente o empresa",\n' +
      '  "palets": número entero de palets o bultos,\n' +
      '  "fecha": "fecha en formato DD/MM/YYYY",\n' +
      '  "matricula": "matrícula del vehículo o null si no aparece",\n' +
      '  "observaciones": "cualquier nota relevante o null"\n' +
      '}\n' +
      'Si un campo no es legible o no aparece en el documento, usa null. ' +
      'Para el campo palets, si aparece una cantidad de bultos, cajas o unidades de transporte, usa ese número.';

    // 4. Llamada a Gemini 1.5 Flash
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
        temperature: 0.1,      // Baja temperatura para mayor precisión
        maxOutputTokens: 512,
      }
    };

    const res = UrlFetchApp.fetch(geminiUrl, {
      method:      'post',
      contentType: 'application/json',
      payload:     JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    if (res.getResponseCode() !== 200) {
      Logger.log('Error Gemini API: ' + res.getContentText());
      return null;
    }

    const raw  = JSON.parse(res.getContentText());
    const text = raw.candidates[0].content.parts[0].text;

    // Limpiar posibles bloques markdown
    const clean = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();

    const parsed = JSON.parse(clean);

    // Normalizar tipo
    if (parsed.tipo) {
      parsed.tipo = parsed.tipo.toUpperCase().includes('SALIDA') ? 'SALIDA' : 'ENTRADA';
    }

    // Normalizar palets a número entero
    if (parsed.palets !== null && parsed.palets !== undefined) {
      parsed.palets = parseInt(parsed.palets) || null;
    }

    Logger.log('Datos extraídos por Gemini: ' + JSON.stringify(parsed));
    return parsed;

  } catch (err) {
    Logger.log('Error en extractDataFromPhoto: ' + err + '\n' + err.stack);
    return null;
  }
}
