/**
 * GmailConnector.gs — Lectura de correos académicos (Capa 2 opcional).
 * Soporta modo 'label' (correos reenviados) y 'institutional' (query académica directa),
 * con precedencia booleana explícita para dominios EdTech frente a -category:promotions
 * y filtrado previo contra IDs_Procesados.
 */

var MAX_GMAIL_THREADS_PER_RUN = 25;

/**
 * Construye la query de búsqueda de Gmail según la configuración.
 * Protege explícitamente dominios académicos/EdTech frente a -category:promotions.
 * @param {Object} config
 * @return {string}
 */
function buildGmailSearchQuery(config) {
  var cfg = config || {};
  var mode = String(cfg.GMAIL_MODE || 'label').toLowerCase();
  if (mode === 'off') return '';

  var edtechTerms = [];
  var domains = (typeof APP_CONSTANTS !== 'undefined' && APP_CONSTANTS.EDTECH_DOMAINS)
    ? APP_CONSTANTS.EDTECH_DOMAINS
    : ['eltngl.com', 'ngl.cengage.com', 'moodle', 'canvas', 'blackboard', 'brightspace'];

  for (var i = 0; i < domains.length; i++) {
    var d = String(domains[i]).trim();
    if (!d) continue;
    edtechTerms.push(d.indexOf('.') !== -1 ? 'from:' + d : d);
  }

  // Precedencia booleana explícita: los correos EdTech se evalúan en OR con (-category:promotions)
  var edtechProtectionGroup = '((' + edtechTerms.join(' OR ') + ') OR (-category:promotions))';

  if (mode === 'label') {
    var label = String(cfg.GMAIL_LABEL || 'Universidad').trim() || 'Universidad';
    var safeLabel = label.replace(/"/g, '');
    return 'label:"' + safeLabel + '" newer_than:14d ' + edtechProtectionGroup;
  }

  // Modo institutional: búsqueda académica directa en el buzón
  var academicKeywords = '(tarea OR taller OR parcial OR examen OR quiz OR entrega OR proyecto OR laboratorio OR assignment OR homework OR deadline OR moodle OR canvas OR blackboard OR brightspace)';
  return academicKeywords + ' newer_than:14d ' + edtechProtectionGroup;
}

/**
 * Lee los IDs ya procesados desde la hoja IDs_Procesados.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} optSpreadsheet
 * @return {Object<string, boolean>}
 */
function getProcessedIdMap(optSpreadsheet) {
  var sheets = ensureSheetStructure(optSpreadsheet);
  var sheet = sheets[SHEET_SCHEMAS.IDs_Procesados.name];
  var lastRow = sheet.getLastRow();
  var map = {};
  if (lastRow < 2) return map;

  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    var id = String(values[i][0] || '').trim();
    if (id) map[id] = true;
  }
  return map;
}

/**
 * Registra nuevos IDs procesados en IDs_Procesados sin duplicar los existentes.
 * @param {Array<{id: string, origen: string, fechaIso: string, resultado: string}>} entries
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} optSpreadsheet
 * @return {number}
 */
function recordProcessedIds(entries, optSpreadsheet) {
  if (!entries || entries.length === 0) return 0;

  var sheets = ensureSheetStructure(optSpreadsheet);
  var sheet = sheets[SHEET_SCHEMAS.IDs_Procesados.name];
  var headers = SHEET_SCHEMAS.IDs_Procesados.headers;
  var existing = getProcessedIdMap(optSpreadsheet);

  var rows = [];
  for (var i = 0; i < entries.length; i++) {
    var e = entries[i];
    var id = String((e && (e.id || e.ID_Referencia)) || '').trim();
    if (!id || existing[id]) continue;
    existing[id] = true;
    rows.push([
      id,
      String(e.origen || e.Origen || 'GMAIL'),
      String(e.fechaIso || e.Fecha_ISO || ''),
      String(e.resultado || e.Resultado || 'OK')
    ]);
  }

  if (rows.length > 0) {
    var startRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
    ensureDateColumnsTextFormat_(sheet, headers, SHEET_SCHEMAS.IDs_Procesados.dateTextColumns);
  }
  return rows.length;
}

/**
 * Pre-filtra mensajes determinísticamente antes de gastar llamadas a Gemini.
 * @param {{id: string, subject: string, body: string}} msg
 * @return {boolean} true si vale la pena analizarlo con IA.
 */
function shouldAnalyzeEmailDeterministically(msg) {
  if (!msg || !msg.id) return false;
  var combined = (String(msg.subject || '') + '\n' + String(msg.body || '')).trim();
  if (combined.length < 20) return false;

  // Descartar notificaciones automáticas puramente de seguridad / login sin contenido académico
  if (/c[oó]digo de verificaci[oó]n|verification code|password reset|restablecer contrase[nñ]a/i.test(combined) &&
      !/tarea|taller|parcial|examen|entrega|quiz|proyecto|assignment/i.test(combined)) {
    return false;
  }
  return true;
}

/**
 * Busca y extrae correos académicos nuevos listos para analizar con Gemini.
 * @param {Object=} optConfig
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} optSpreadsheet
 * @return {{messages: Object[], skippedProcessed: number, error: string|null}}
 */
function fetchPendingAcademicEmails(optConfig, optSpreadsheet) {
  var config = optConfig || getUserConfig();
  if (config.GMAIL_MODE === 'off') {
    return { messages: [], skippedProcessed: 0, error: null };
  }

  var query = buildGmailSearchQuery(config);
  if (!query) {
    return { messages: [], skippedProcessed: 0, error: null };
  }

  var processedMap = getProcessedIdMap(optSpreadsheet);
  var messages = [];
  var skippedProcessed = 0;

  try {
    var threads = GmailApp.search(query, 0, MAX_GMAIL_THREADS_PER_RUN);
    for (var i = 0; i < threads.length; i++) {
      var threadMsgs = threads[i].getMessages();
      for (var j = 0; j < threadMsgs.length; j++) {
        var m = threadMsgs[j];
        var msgId = 'gmail:' + m.getId();
        if (processedMap[msgId]) {
          skippedProcessed++;
          continue;
        }

        var subject = String(m.getSubject() || '').trim();
        var body = String(m.getPlainBody() || '').slice(0, 6000);
        if (!shouldAnalyzeEmailDeterministically({ id: msgId, subject: subject, body: body })) {
          continue;
        }

        var msgDate = m.getDate();
        var dateIso = formatIsoInTimezone(msgDate, config.TIMEZONE);
        var urls = extractUrlsFromText_(body);

        messages.push({
          id: msgId,
          subject: subject,
          from: String(m.getFrom() || ''),
          dateIso: dateIso,
          body: body,
          urls: urls
        });
      }
    }
  } catch (err) {
    return {
      messages: messages,
      skippedProcessed: skippedProcessed,
      error: 'Error leyendo Gmail: ' + err.message
    };
  }

  return {
    messages: messages,
    skippedProcessed: skippedProcessed,
    error: null
  };
}
