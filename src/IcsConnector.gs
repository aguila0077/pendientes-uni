/**
 * IcsConnector.gs — Conector .ics (Capa 1, 0 tokens, sin API Key).
 * Descarga y parsea feeds iCalendar (Moodle, Canvas, Blackboard) usando ICAL.parse(),
 * normaliza fechas a CONFIG.TIMEZONE, filtra ventana [-7, +120] días y autodetecta materias.
 */

var ICS_WINDOW_PAST_DAYS = 7;
var ICS_WINDOW_FUTURE_DAYS = 120;

/**
 * Descarga el feed .ics, lo parsea, registra materias nuevas y retorna ítems normalizados.
 * @param {Object=} optConfig
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} optSpreadsheet
 * @return {{items: Object[], detectedSubjects: Object[], error: string|null}}
 */
function fetchAndParseIcs(optConfig, optSpreadsheet) {
  var config = optConfig || getUserConfig();
  if (!config.ICS_URL) {
    return { items: [], detectedSubjects: [], error: null };
  }

  var response;
  try {
    response = UrlFetchApp.fetch(config.ICS_URL, {
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true
    });
  } catch (err) {
    return { items: [], detectedSubjects: [], error: 'Error de red descargando .ics: ' + err.message };
  }

  var status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    return { items: [], detectedSubjects: [], error: 'HTTP ' + status + ' al descargar .ics' };
  }

  var icsText = response.getContentText('UTF-8');
  var parsed = parseIcsContent(icsText, config, new Date());

  if (parsed.detectedSubjects.length > 0 && typeof SpreadsheetApp !== 'undefined') {
    ensureSubjectsInConfig(parsed.detectedSubjects, optSpreadsheet);
  }

  return {
    items: parsed.items,
    detectedSubjects: parsed.detectedSubjects,
    error: null
  };
}

/**
 * Parsea texto .ics de forma pura (testeable offline en Node.js y Apps Script).
 * @param {string} icsText
 * @param {Object=} optConfig
 * @param {Date=} optNow
 * @return {{items: Object[], detectedSubjects: Object[]}}
 */
function parseIcsContent(icsText, optConfig, optNow) {
  if (!icsText || typeof icsText !== 'string' || icsText.indexOf('BEGIN:VCALENDAR') === -1) {
    return { items: [], detectedSubjects: [] };
  }

  var config = optConfig || { TIMEZONE: 'UTC' };
  var tz = config.TIMEZONE || 'UTC';
  var now = optNow instanceof Date ? optNow : new Date();
  var minMs = now.getTime() - ICS_WINDOW_PAST_DAYS * 24 * 60 * 60 * 1000;
  var maxMs = now.getTime() + ICS_WINDOW_FUTURE_DAYS * 24 * 60 * 60 * 1000;

  var jcalData = ICAL.parse(icsText);
  var comp = new ICAL.Component(jcalData);
  registerVTimezones_(comp);

  var vevents = comp.getAllSubcomponents('vevent');
  var items = [];
  var subjectMap = {};

  for (var i = 0; i < vevents.length; i++) {
    var vevent = vevents[i];
    var event = new ICAL.Event(vevent);
    if (!event.startDate) continue;

    var isAllDay = Boolean(event.startDate.isDate);
    var startDateJs = icalTimeToDate_(event.startDate, tz);
    if (!startDateJs) continue;

    var endDateJs = event.endDate ? icalTimeToDate_(event.endDate, tz) : new Date(startDateJs.getTime());
    if (!endDateJs) endDateJs = new Date(startDateJs.getTime());

    // Para eventos de entrega (Moodle usa DTSTART/DTEND iguales o cortos; usar DTEND como límite si existe)
    var dueJs = endDateJs.getTime() >= startDateJs.getTime() ? endDateJs : startDateJs;

    if (dueJs.getTime() < minMs || dueJs.getTime() > maxMs) {
      continue;
    }

    var categories = extractCategories_(vevent);
    var materia = categories.length > 0 ? normalizeSubjectName_(categories[0]) : 'General';
    if (categories.length > 0) {
      var codeKey = categories[0].trim();
      if (codeKey && !subjectMap[codeKey]) {
        subjectMap[codeKey] = {
          Codigo_o_Alias: codeKey,
          Nombre_Materia: materia,
          Activa: true,
          Origen_Deteccion: 'ICS_CATEGORIES'
        };
      }
    }

    var summary = cleanIcsText_(event.summary || 'Actividad sin título');
    var description = cleanIcsText_(event.description || '');
    var urls = extractUrlsFromText_(description);
    var urlProp = vevent.getFirstPropertyValue('url');
    if (urlProp && typeof urlProp === 'string' && urls.indexOf(urlProp.trim()) === -1) {
      urls.unshift(urlProp.trim());
    }

    var lastModProp = vevent.getFirstPropertyValue('last-modified') || vevent.getFirstPropertyValue('dtstamp');
    var lastModJs = lastModProp ? icalTimeToDate_(lastModProp, tz) : now;
    if (!lastModJs) lastModJs = now;

    var fechaInicioIso = isAllDay
      ? formatAllDayDate_(startDateJs, tz)
      : formatIsoInTimezone(startDateJs, tz);
    var fechaLimiteIso = isAllDay
      ? formatAllDayDate_(dueJs, tz)
      : formatIsoInTimezone(dueJs, tz);

    items.push({
      uid: String(event.uid || '').trim(),
      ID_Referencia: String(event.uid || '').trim(),
      Materia: materia,
      Actividad: summary,
      Descripcion: description,
      Fecha_Inicio_ISO: fechaInicioIso,
      Fecha_Limite: fechaLimiteIso,
      hasTime: !isAllDay,
      Origen: 'ICS',
      URLs: urls.join(' | '),
      Ultima_Actualizacion_ISO: formatIsoInTimezone(lastModJs, tz),
      accion: 'crear_o_actualizar',
      es_pendiente: true
    });
  }

  var detectedSubjects = [];
  var subjectKeys = Object.keys(subjectMap);
  for (var s = 0; s < subjectKeys.length; s++) {
    detectedSubjects.push(subjectMap[subjectKeys[s]]);
  }

  return {
    items: items,
    detectedSubjects: detectedSubjects
  };
}

/**
 * Registra VTIMEZONE en ICAL.TimezoneService si el calendario los incluye.
 * @private
 */
function registerVTimezones_(rootComp) {
  if (!ICAL.TimezoneService) return;
  var vtzs = rootComp.getAllSubcomponents('vtimezone');
  for (var i = 0; i < vtzs.length; i++) {
    try {
      var tzid = vtzs[i].getFirstPropertyValue('tzid');
      if (tzid && !ICAL.TimezoneService.has(tzid)) {
        ICAL.TimezoneService.register(tzid, new ICAL.Timezone(vtzs[i]));
      }
    } catch (e) {
      // Ignorar VTIMEZONE malformado
    }
  }
}

/**
 * Convierte ICAL.Time a Date JS respetando UTC, TZID o fallback a tz de usuario.
 * @private
 */
function icalTimeToDate_(icalTime, fallbackTz) {
  if (!icalTime) return null;
  try {
    if (icalTime.isDate) {
      return buildDateInTimezone_(icalTime.year, icalTime.month - 1, icalTime.day, 0, 0, 0, fallbackTz);
    }
    if (typeof icalTime.toJSDate === 'function') {
      var d = icalTime.toJSDate();
      if (d && !isNaN(d.getTime())) return d;
    }
  } catch (e) {
    // Fallback manual
  }
  return buildDateInTimezone_(
    icalTime.year || 1970,
    (icalTime.month || 1) - 1,
    icalTime.day || 1,
    icalTime.hour || 0,
    icalTime.minute || 0,
    icalTime.second || 0,
    fallbackTz
  );
}

/**
 * Formatea fecha de día completo como YYYY-MM-DD en la zona horaria indicada.
 * @private
 */
function formatAllDayDate_(dateVal, tz) {
  if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.formatDate === 'function') {
    return Utilities.formatDate(dateVal, tz || 'UTC', 'yyyy-MM-dd');
  }
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: tz || 'UTC',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).format(dateVal);
}

/**
 * Extrae categorías de un VEVENT manejando múltiples propiedades o valores separados por coma.
 * @private
 */
function extractCategories_(vevent) {
  var props = vevent.getAllProperties('categories');
  var out = [];
  for (var i = 0; i < props.length; i++) {
    var vals = props[i].getValues();
    for (var j = 0; j < vals.length; j++) {
      var raw = String(vals[j] || '').trim();
      if (!raw) continue;
      var parts = raw.split(',');
      for (var k = 0; k < parts.length; k++) {
        var clean = parts[k].trim();
        if (clean && out.indexOf(clean) === -1) {
          out.push(clean);
        }
      }
    }
  }
  return out;
}

/**
 * Normaliza un nombre de materia sin cortar información útil.
 * @private
 */
function normalizeSubjectName_(raw) {
  return String(raw || '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Limpia saltos de línea escapados de iCalendar.
 * @private
 */
function cleanIcsText_(txt) {
  return String(txt || '')
    .replace(/\\n/g, '\n')
    .replace(/\\,/g, ',')
    .replace(/\\;/g, ';')
    .trim();
}

/**
 * Extrae URLs únicas de un bloque de texto.
 * @param {string} text
 * @return {string[]}
 */
function extractUrlsFromText_(text) {
  if (!text) return [];
  var matches = String(text).match(/https?:\/\/[^\s<>"')]+/gi);
  if (!matches) return [];
  var unique = [];
  for (var i = 0; i < matches.length; i++) {
    var u = matches[i].replace(/[.,;]+$/, '');
    if (unique.indexOf(u) === -1) {
      unique.push(u);
    }
  }
  return unique;
}

/**
 * Inserta materias autodetectadas en Config_Materias sin duplicar Codigo_o_Alias.
 * @param {Object[]} detectedSubjects
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} optSpreadsheet
 * @return {number} Cantidad de materias nuevas insertadas.
 */
function ensureSubjectsInConfig(detectedSubjects, optSpreadsheet) {
  if (!detectedSubjects || detectedSubjects.length === 0) return 0;

  var sheets = ensureSheetStructure(optSpreadsheet);
  var sheet = sheets[SHEET_SCHEMAS.Config_Materias.name];
  var headers = SHEET_SCHEMAS.Config_Materias.headers;
  var lastRow = sheet.getLastRow();

  var existingCodes = {};
  if (lastRow >= 2) {
    var existingVals = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (var i = 0; i < existingVals.length; i++) {
      var code = String(existingVals[i][0] || '').trim().toLowerCase();
      if (code) existingCodes[code] = true;
    }
  }

  var toInsert = [];
  for (var j = 0; j < detectedSubjects.length; j++) {
    var subj = detectedSubjects[j];
    var alias = String(subj.Codigo_o_Alias || '').trim();
    if (!alias) continue;
    var key = alias.toLowerCase();
    if (existingCodes[key]) continue;

    existingCodes[key] = true;
    toInsert.push([
      alias,
      String(subj.Nombre_Materia || alias),
      subj.Activa !== false,
      String(subj.Origen_Deteccion || 'ICS_CATEGORIES')
    ]);
  }

  if (toInsert.length > 0) {
    var startRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(startRow, 1, toInsert.length, headers.length).setValues(toInsert);
  }

  return toInsert.length;
}
