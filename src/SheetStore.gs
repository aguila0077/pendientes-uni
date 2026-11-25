/**
 * SheetStore.gs — Capa de persistencia en Google Sheets.
 * Inicialización idempotente, formato texto (@) en fechas, lectura segura Date|string,
 * escritura quirúrgica preservando Completada y Notas_Personales, poda >60d y latido A1.
 */

var SHEET_SCHEMAS = {
  Resumen_Tareas: {
    name: 'Resumen_Tareas',
    headers: [
      'Clave_Semantica',
      'Materia',
      'Actividad',
      'Fecha_Limite',
      'Fecha_Inicio_ISO',
      'Estado',
      'Completada',
      'Origen',
      'ID_Referencia',
      'URLs',
      'Ultima_Actualizacion_ISO',
      'Calendar_Event_ID',
      'Tasks_ID',
      'Notas_Personales'
    ],
    dateTextColumns: ['Fecha_Limite', 'Fecha_Inicio_ISO', 'Ultima_Actualizacion_ISO']
  },
  Config_Materias: {
    name: 'Config_Materias',
    headers: [
      'Codigo_o_Alias',
      'Nombre_Materia',
      'Activa',
      'Origen_Deteccion'
    ],
    dateTextColumns: []
  },
  IDs_Procesados: {
    name: 'IDs_Procesados',
    headers: [
      'ID_Referencia',
      'Origen',
      'Fecha_ISO',
      'Resultado'
    ],
    dateTextColumns: ['Fecha_ISO']
  }
};

var WELCOME_SHEET_NAME = '👋 Empieza_Aquí';

/**
 * Garantiza que la pestaña de bienvenida y las 3 pestañas de datos existan.
 * Es idempotente: nunca borra ni sobrescribe datos existentes.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} optSpreadsheet
 * @return {Object} Mapa de hojas inicializadas.
 */
function ensureSheetStructure(optSpreadsheet) {
  var ss = optSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No hay Spreadsheet activo disponible.');
  }

  var result = {};
  var welcomeSheet = ensureWelcomeSheet_(ss);
  if (welcomeSheet) {
    result[welcomeSheet.getName ? welcomeSheet.getName() : WELCOME_SHEET_NAME] = welcomeSheet;
  }

  var schemaKeys = Object.keys(SHEET_SCHEMAS);

  for (var i = 0; i < schemaKeys.length; i++) {
    var schema = SHEET_SCHEMAS[schemaKeys[i]];
    var sheet = ss.getSheetByName(schema.name);
    if (!sheet) {
      sheet = ss.insertSheet(schema.name);
    }

    ensureHeaders_(sheet, schema.headers);
    ensureDateColumnsTextFormat_(sheet, schema.headers, schema.dateTextColumns);
    result[schema.name] = sheet;
  }

  return result;
}

/**
 * Crea o conserva la primera pestaña visual 👋 Empieza_Aquí con 4 pasos simples.
 * Nunca la borra ni sobrescribe su contenido si ya tiene datos.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 * @private
 */
function ensureWelcomeSheet_(ss) {
  var welcomeName = (typeof APP_CONSTANTS !== 'undefined' && APP_CONSTANTS.SHEETS && APP_CONSTANTS.SHEETS.WELCOME)
    ? APP_CONSTANTS.SHEETS.WELCOME
    : WELCOME_SHEET_NAME;

  var sheet = ss.getSheetByName(welcomeName);
  if (!sheet) {
    sheet = ss.insertSheet(welcomeName, 0);
  }
  if (!sheet) return null;

  if (typeof sheet.getLastRow === 'function' && sheet.getLastRow() > 0) {
    return sheet;
  }

  var rows = [
    ['🎓 ¡Bienvenido a Pendientes Uni!'],
    ['Conecta tus entregas de la universidad con Google Calendar y Google Tasks en tu celular en 3 minutos (sin saber programar).'],
    [''],
    ['PASO 1', 'Haz clic arriba en el menú 🎓 Pendientes Uni → 🚀 Abrir asistente (Empezar aquí)'],
    ['PASO 2', "Si Google te muestra una alerta que dice 'Google no ha verificado esta app', es normal porque es tu propia copia: haz clic en Opciones avanzadas → Ir a Pendientes Uni → Permitir"],
    ['PASO 3', 'Sigue los 3 pasos del panel que se abre a la derecha'],
    ['PASO 4', '¡Listo! Ya puedes cerrar esta página y ver tus tareas en Google Calendar y Google Tasks en tu celular'],
    [''],
    ['💡 Tip: Las demás pestañas (Resumen_Tareas, Config_Materias, IDs_Procesados) se llenan solas automáticamente. No tienes que tocarlas.']
  ];

  sheet.getRange(1, 1, rows.length, 1);
  sheet.getRange(1, 1, 1, 2).setValues([[rows[0][0], '']]);
  sheet.getRange(2, 1, 1, 2).setValues([[rows[1][0], '']]);
  sheet.getRange(4, 1, 4, 2).setValues(rows.slice(3, 7));
  sheet.getRange(9, 1, 1, 2).setValues([[rows[8][0], '']]);

  try {
    if (typeof sheet.setHiddenGridlines === 'function') sheet.setHiddenGridlines(true);
    if (typeof sheet.setColumnWidth === 'function') {
      sheet.setColumnWidth(1, 130);
      sheet.setColumnWidth(2, 760);
    }
    var r1 = sheet.getRange(1, 1, 1, 2);
    if (typeof r1.merge === 'function') r1.merge();
    if (typeof r1.setBackground === 'function') r1.setBackground('#1d4ed8');
    if (typeof r1.setFontColor === 'function') r1.setFontColor('#ffffff');
    if (typeof r1.setFontSize === 'function') r1.setFontSize(18);
    r1.setFontWeight('bold');

    var r2 = sheet.getRange(2, 1, 1, 2);
    if (typeof r2.merge === 'function') r2.merge();
    if (typeof r2.setBackground === 'function') r2.setBackground('#eff6ff');
    if (typeof r2.setFontColor === 'function') r2.setFontColor('#1e3a8a');
    if (typeof r2.setFontSize === 'function') r2.setFontSize(11);
    if (typeof r2.setWrap === 'function') r2.setWrap(true);

    var badges = sheet.getRange(4, 1, 4, 1);
    if (typeof badges.setBackground === 'function') badges.setBackground('#dbeafe');
    if (typeof badges.setFontColor === 'function') badges.setFontColor('#1e40af');
    if (typeof badges.setFontSize === 'function') badges.setFontSize(12);
    badges.setFontWeight('bold');
    if (typeof badges.setVerticalAlignment === 'function') badges.setVerticalAlignment('middle');

    var stepsRange = sheet.getRange(4, 2, 4, 1);
    if (typeof stepsRange.setBackground === 'function') stepsRange.setBackground('#f8fafc');
    if (typeof stepsRange.setFontColor === 'function') stepsRange.setFontColor('#0f172a');
    if (typeof stepsRange.setFontSize === 'function') stepsRange.setFontSize(12);
    if (typeof stepsRange.setWrap === 'function') stepsRange.setWrap(true);
    if (typeof stepsRange.setVerticalAlignment === 'function') stepsRange.setVerticalAlignment('middle');

    var r9 = sheet.getRange(9, 1, 1, 2);
    if (typeof r9.merge === 'function') r9.merge();
    if (typeof r9.setFontColor === 'function') r9.setFontColor('#475569');
    if (typeof r9.setFontSize === 'function') r9.setFontSize(10);
    if (typeof r9.setWrap === 'function') r9.setWrap(true);

    if (typeof sheet.setRowHeight === 'function') {
      sheet.setRowHeight(1, 48);
      sheet.setRowHeight(2, 38);
      sheet.setRowHeight(4, 54);
      sheet.setRowHeight(5, 72);
      sheet.setRowHeight(6, 54);
      sheet.setRowHeight(7, 54);
    }
  } catch (styleErr) {
    // Estilo visual opcional: nunca debe bloquear la inicialización
  }

  return sheet;
}

/**
 * Escribe o valida la fila 1 de encabezados sin borrar filas de datos.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} expectedHeaders
 */
function ensureHeaders_(sheet, expectedHeaders) {
  var lastCol = Math.max(sheet.getLastColumn(), expectedHeaders.length);
  var currentRow = [];
  if (sheet.getLastRow() >= 1 && lastCol > 0) {
    currentRow = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];
  }

  var needsUpdate = false;
  for (var i = 0; i < expectedHeaders.length; i++) {
    if (String(currentRow[i] || '').trim() !== expectedHeaders[i]) {
      needsUpdate = true;
      break;
    }
  }

  if (needsUpdate) {
    var headerRange = sheet.getRange(1, 1, 1, expectedHeaders.length);
    headerRange.setValues([expectedHeaders]);
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

/**
 * Aplica formato texto ('@') a todas las columnas de fecha de una hoja.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} headers
 * @param {string[]} dateColumns
 */
function ensureDateColumnsTextFormat_(sheet, headers, dateColumns) {
  if (!dateColumns || dateColumns.length === 0) return;
  var maxRows = Math.max(sheet.getMaxRows(), 2);

  for (var i = 0; i < dateColumns.length; i++) {
    var colIdx = headers.indexOf(dateColumns[i]) + 1;
    if (colIdx > 0) {
      sheet.getRange(2, colIdx, maxRows - 1, 1).setNumberFormat('@');
    }
  }
}

/**
 * Parsea fechas de forma segura soportando tanto objetos Date como strings ISO o locales.
 * @param {Date|string|number} val
 * @param {string=} optTz
 * @return {Date|null}
 */
function parseDateSafe(val, optTz) {
  if (val === null || val === undefined || val === '') return null;

  if (Object.prototype.toString.call(val) === '[object Date]') {
    return isNaN(val.getTime()) ? null : new Date(val.getTime());
  }

  if (typeof val === 'number') {
    var fromNum = new Date(val);
    return isNaN(fromNum.getTime()) ? null : fromNum;
  }

  var str = String(val).trim();
  if (!str) return null;

  // Caso YYYY-MM-DD puro: construir en fecha local/media mañana o medianoche segura
  var dateOnlyMatch = /^(\d{4})-(\d{2})-(\d{2})$/.exec(str);
  if (dateOnlyMatch) {
    var y = parseInt(dateOnlyMatch[1], 10);
    var m = parseInt(dateOnlyMatch[2], 10) - 1;
    var d = parseInt(dateOnlyMatch[3], 10);
    var tz = optTz || (typeof getUserConfig === 'function' ? getUserConfig().TIMEZONE : 'UTC');
    return buildDateInTimezone_(y, m, d, 0, 0, 0, tz);
  }

  // Caso YYYY-MM-DD HH:mm(:ss)? sin offset explícito
  var localDtMatch = /^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})(?::(\d{2}))?$/.exec(str);
  if (localDtMatch) {
    var ly = parseInt(localDtMatch[1], 10);
    var lm = parseInt(localDtMatch[2], 10) - 1;
    var ld = parseInt(localDtMatch[3], 10);
    var lh = parseInt(localDtMatch[4], 10);
    var lmin = parseInt(localDtMatch[5], 10);
    var ls = localDtMatch[6] ? parseInt(localDtMatch[6], 10) : 0;
    var ltz = optTz || (typeof getUserConfig === 'function' ? getUserConfig().TIMEZONE : 'UTC');
    return buildDateInTimezone_(ly, lm, ld, lh, lmin, ls, ltz);
  }

  var parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Construye un Date ajustado a una zona horaria específica para fechas sin offset.
 * @private
 */
function buildDateInTimezone_(year, monthZeroBased, day, hour, minute, second, tz) {
  var utcGuess = new Date(Date.UTC(year, monthZeroBased, day, hour, minute, second));
  if (!tz || tz === 'UTC' || tz === 'Etc/UTC') {
    return utcGuess;
  }
  try {
    if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.formatDate === 'function') {
      var formatted = Utilities.formatDate(utcGuess, tz, "yyyy-MM-dd'T'HH:mm:ss");
      var parts = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/.exec(formatted);
      if (parts) {
        var asUtc = Date.UTC(
          parseInt(parts[1], 10),
          parseInt(parts[2], 10) - 1,
          parseInt(parts[3], 10),
          parseInt(parts[4], 10),
          parseInt(parts[5], 10),
          parseInt(parts[6], 10)
        );
        var offsetMs = asUtc - utcGuess.getTime();
        return new Date(utcGuess.getTime() - offsetMs);
      }
    }
  } catch (e) {
    // Fallback a fecha local del runtime
  }
  return new Date(year, monthZeroBased, day, hour, minute, second);
}

/**
 * Formatea una fecha a ISO en la zona horaria indicada (o YYYY-MM-DD'T'HH:mm:ssXXX).
 * @param {Date|string} dateVal
 * @param {string=} optTz
 * @return {string}
 */
function formatIsoInTimezone(dateVal, optTz) {
  var d = parseDateSafe(dateVal, optTz);
  if (!d) return '';
  var tz = optTz || (typeof getUserConfig === 'function' ? getUserConfig().TIMEZONE : 'UTC');
  if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.formatDate === 'function') {
    return Utilities.formatDate(d, tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
  }
  return d.toISOString();
}

/**
 * Convierte una fila de Resumen_Tareas en objeto indexado por nombre de columna.
 * @param {Array} rowValues
 * @param {number} rowNumber Número de fila 1-based en la hoja.
 * @return {Object}
 */
function mapResumenRow_(rowValues, rowNumber) {
  var headers = SHEET_SCHEMAS.Resumen_Tareas.headers;
  var obj = { _rowNumber: rowNumber };
  for (var i = 0; i < headers.length; i++) {
    obj[headers[i]] = rowValues[i] !== undefined ? rowValues[i] : '';
  }
  obj.Completada = obj.Completada === true || String(obj.Completada).toUpperCase() === 'TRUE';
  return obj;
}

/**
 * Lee todas las tareas existentes en Resumen_Tareas.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} optSpreadsheet
 * @return {Object[]}
 */
function readAllTasks(optSpreadsheet) {
  var sheets = ensureSheetStructure(optSpreadsheet);
  var sheet = sheets[SHEET_SCHEMAS.Resumen_Tareas.name];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var headers = SHEET_SCHEMAS.Resumen_Tareas.headers;
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var tasks = [];

  for (var i = 0; i < values.length; i++) {
    var clave = String(values[i][0] || '').trim();
    if (!clave) continue;
    tasks.push(mapResumenRow_(values[i], i + 2));
  }
  return tasks;
}

/**
 * Escritura quirúrgica de tareas en Resumen_Tareas por Clave_Semantica.
 * Nunca sobrescribe Completada (si ya existe la fila, salvo que se pase _forceCompletada)
 * ni Notas_Personales. Garantiza formato '@' en columnas de fecha.
 * @param {Object[]} taskPatches
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} optSpreadsheet
 * @return {{inserted: number, updated: number}}
 */
function upsertTasksSurgical(taskPatches, optSpreadsheet) {
  if (!taskPatches || taskPatches.length === 0) {
    return { inserted: 0, updated: 0 };
  }

  var sheets = ensureSheetStructure(optSpreadsheet);
  var sheet = sheets[SHEET_SCHEMAS.Resumen_Tareas.name];
  var headers = SHEET_SCHEMAS.Resumen_Tareas.headers;
  var existing = readAllTasks(optSpreadsheet);

  var byKey = {};
  for (var i = 0; i < existing.length; i++) {
    byKey[existing[i].Clave_Semantica] = existing[i];
  }

  var inserted = 0;
  var updated = 0;
  var rowsToAppend = [];

  for (var p = 0; p < taskPatches.length; p++) {
    var patch = taskPatches[p];
    if (!patch || !patch.Clave_Semantica) continue;

    var current = byKey[patch.Clave_Semantica];
    if (!current) {
      var newRowObj = {
        Clave_Semantica: String(patch.Clave_Semantica),
        Materia: patch.Materia !== undefined ? String(patch.Materia) : '',
        Actividad: patch.Actividad !== undefined ? String(patch.Actividad) : '',
        Fecha_Limite: patch.Fecha_Limite !== undefined ? String(patch.Fecha_Limite) : '',
        Fecha_Inicio_ISO: patch.Fecha_Inicio_ISO !== undefined ? String(patch.Fecha_Inicio_ISO) : '',
        Estado: patch.Estado !== undefined ? String(patch.Estado) : 'Pendiente',
        Completada: patch.Completada === true,
        Origen: patch.Origen !== undefined ? String(patch.Origen) : '',
        ID_Referencia: patch.ID_Referencia !== undefined ? String(patch.ID_Referencia) : '',
        URLs: patch.URLs !== undefined ? String(patch.URLs) : '',
        Ultima_Actualizacion_ISO: patch.Ultima_Actualizacion_ISO !== undefined ? String(patch.Ultima_Actualizacion_ISO) : '',
        Calendar_Event_ID: patch.Calendar_Event_ID !== undefined ? String(patch.Calendar_Event_ID) : '',
        Tasks_ID: patch.Tasks_ID !== undefined ? String(patch.Tasks_ID) : '',
        Notas_Personales: patch.Notas_Personales !== undefined ? String(patch.Notas_Personales) : ''
      };
      var rowArr = [];
      for (var h = 0; h < headers.length; h++) {
        rowArr.push(newRowObj[headers[h]]);
      }
      rowsToAppend.push(rowArr);
      byKey[patch.Clave_Semantica] = newRowObj;
      inserted++;
    } else {
      // Actualización quirúrgica: preservar siempre Completada (salvo flag explícito) y Notas_Personales
      var merged = [];
      var changed = false;
      for (var c = 0; c < headers.length; c++) {
        var colName = headers[c];
        var oldVal = current[colName];
        var nextVal = oldVal;

        if (colName === 'Notas_Personales') {
          nextVal = oldVal;
        } else if (colName === 'Completada') {
          nextVal = patch._forceCompletada === true ? Boolean(patch.Completada) : Boolean(oldVal);
        } else if (Object.prototype.hasOwnProperty.call(patch, colName) && patch[colName] !== undefined) {
          nextVal = patch[colName];
        }

        if (String(nextVal) !== String(oldVal)) {
          changed = true;
        }
        merged.push(nextVal);
      }

      if (changed && current._rowNumber) {
        var rowRange = sheet.getRange(current._rowNumber, 1, 1, headers.length);
        rowRange.setValues([merged]);
        ensureRowDateFormat_(sheet, current._rowNumber, headers, SHEET_SCHEMAS.Resumen_Tareas.dateTextColumns);
        updated++;
      }
    }
  }

  if (rowsToAppend.length > 0) {
    var startRow = Math.max(sheet.getLastRow() + 1, 2);
    var appendRange = sheet.getRange(startRow, 1, rowsToAppend.length, headers.length);
    appendRange.setValues(rowsToAppend);
    for (var d = 0; d < SHEET_SCHEMAS.Resumen_Tareas.dateTextColumns.length; d++) {
      var colIdx = headers.indexOf(SHEET_SCHEMAS.Resumen_Tareas.dateTextColumns[d]) + 1;
      if (colIdx > 0) {
        sheet.getRange(startRow, colIdx, rowsToAppend.length, 1).setNumberFormat('@');
      }
    }
  }

  return { inserted: inserted, updated: updated };
}

/**
 * Reaplica formato '@' en columnas de fecha para una fila puntual.
 * @private
 */
function ensureRowDateFormat_(sheet, rowNumber, headers, dateColumns) {
  for (var i = 0; i < dateColumns.length; i++) {
    var colIdx = headers.indexOf(dateColumns[i]) + 1;
    if (colIdx > 0) {
      sheet.getRange(rowNumber, colIdx, 1, 1).setNumberFormat('@');
    }
  }
}

/**
 * Poda registros en IDs_Procesados con antigüedad mayor a maxAgeDays (default 60 días).
 * @param {number=} maxAgeDays
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} optSpreadsheet
 * @return {number} Cantidad de filas eliminadas.
 */
function pruneProcessedIds(maxAgeDays, optSpreadsheet) {
  var days = typeof maxAgeDays === 'number' && maxAgeDays > 0
    ? maxAgeDays
    : APP_CONSTANTS.RETENTION_DAYS_PROCESSED;

  var sheets = ensureSheetStructure(optSpreadsheet);
  var sheet = sheets[SHEET_SCHEMAS.IDs_Procesados.name];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  var headers = SHEET_SCHEMAS.IDs_Procesados.headers;
  var fechaColIdx = headers.indexOf('Fecha_ISO');
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var cutoffMs = Date.now() - days * 24 * 60 * 60 * 1000;
  var tz = typeof getUserConfig === 'function' ? getUserConfig().TIMEZONE : 'UTC';

  var keptRows = [];
  var removedCount = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[0] && !row[1] && !row[2]) continue;

    var parsedDate = parseDateSafe(row[fechaColIdx], tz);
    if (parsedDate && parsedDate.getTime() < cutoffMs) {
      removedCount++;
    } else {
      keptRows.push(row);
    }
  }

  if (removedCount === 0) return 0;

  sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  if (keptRows.length > 0) {
    sheet.getRange(2, 1, keptRows.length, headers.length).setValues(keptRows);
    ensureDateColumnsTextFormat_(sheet, headers, SHEET_SCHEMAS.IDs_Procesados.dateTextColumns);
  }

  return removedCount;
}

/**
 * Escribe el latido estructurado en la nota de la celda A1 de Resumen_Tareas.
 * @param {{estado?: string, ics_err?: number, gmail_err?: number, gemini_err?: number, sync_err?: number, dl_err?: number, detalle?: string}} metrics
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} optSpreadsheet
 * @return {string} Texto escrito en la nota A1.
 */
function writeHeartbeatNote(metrics, optSpreadsheet) {
  var sheets = ensureSheetStructure(optSpreadsheet);
  var sheet = sheets[SHEET_SCHEMAS.Resumen_Tareas.name];
  var m = metrics || {};

  var icsErr = Number(m.ics_err || 0);
  var gmailErr = Number(m.gmail_err || 0);
  var geminiErr = Number(m.gemini_err || 0);
  var syncErr = Number(m.sync_err || 0);
  var dlErr = Number(m.dl_err || 0);

  var hasErrors = (icsErr + gmailErr + geminiErr + syncErr + dlErr) > 0;
  var estado = m.estado || (hasErrors ? 'DEGRADADO' : 'OK');
  var tz = typeof getUserConfig === 'function' ? getUserConfig().TIMEZONE : 'UTC';
  var timestamp = formatIsoInTimezone(new Date(), tz);

  var lines = [
    'Pendientes Uni — Latido de Sincronización',
    'estado=' + estado,
    'timestamp=' + timestamp,
    'ics_err=' + icsErr,
    'gmail_err=' + gmailErr,
    'gemini_err=' + geminiErr,
    'sync_err=' + syncErr,
    'dl_err=' + dlErr
  ];

  if (m.detalle) {
    lines.push('detalle=' + String(m.detalle));
  }

  var noteText = lines.join('\n');
  sheet.getRange('A1').setNote(noteText);
  return noteText;
}
