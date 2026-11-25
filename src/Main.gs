/**
 * Main.gs — Orquestador principal, menú en Sheets, triggers 24/7 con LockService
 * y endpoints API consumidos por Sidebar.html.
 */

var SYNC_TRIGGER_HANDLER = 'ejecutarSincronizacion';

/**
 * Se ejecuta al abrir la hoja de cálculo y crea el menú amigable.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎓 Pendientes Uni')
    .addItem('🚀 Abrir asistente (Empezar aquí)', 'mostrarSidebar')
    .addItem('🔄 Actualizar mis tareas ahora', 'ejecutarSincronizacionManual')
    .addToUi();
}

/**
 * Muestra el asistente visual en la barra lateral derecha.
 */
function mostrarSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('src/Sidebar')
    .setTitle('🎓 Pendientes Uni')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Ejecución manual desde el menú con feedback visual amigable en Sheets.
 */
function ejecutarSincronizacionManual() {
  var res = ejecutarSincronizacion();
  var ui = SpreadsheetApp.getUi();
  if (res.skippedLock) {
    ui.alert('Ya estoy actualizando tus tareas en este momento. Espera unos segundos y revisa de nuevo.');
    return;
  }
  if (res.ok) {
    ui.alert(
      '¡Listo! Tus tareas quedaron actualizadas.\n\n' +
      '• Tareas nuevas: ' + res.created + '\n' +
      '• Actualizadas: ' + res.updated + '\n\n' +
      'Ya puedes verlas en Google Calendar y Google Tasks en tu celular.'
    );
    return;
  }
  ui.alert(
    'Revisé tus tareas, pero hubo un problema temporal con alguna conexión.\n\n' +
    '• Tareas nuevas: ' + res.created + '\n' +
    '• Actualizadas: ' + res.updated + '\n\n' +
    'Intenta de nuevo en unos minutos o revisa tu configuración en el asistente.'
  );
}

/**
 * Orquestador completo protegido por LockService.getScriptLock().
 * Orden estricto:
 * 1) Lectura bidireccional de completados
 * 2) IcsConnector (si hay URL)
 * 3) GmailConnector + GeminiClient (solo si hay API Key y Gmail activo)
 * 4) TaskEngine -> SheetStore
 * 5) CalendarTasksSync
 * 6) Poda >60d -> Resumen/Versión -> Latido A1
 * @return {Object}
 */
function ejecutarSincronizacion() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return { ok: false, skippedLock: true, estado: 'OCUPADO' };
  }

  var runStartMs = Date.now();
  var metrics = {
    estado: 'OK',
    ics_err: 0,
    gmail_err: 0,
    gemini_err: 0,
    sync_err: 0,
    dl_err: 0,
    detalle: ''
  };

  var createdCount = 0;
  var updatedCount = 0;
  var calendarSynced = 0;
  var tasksSynced = 0;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheetStructure(ss);
    var config = getUserConfig();

    // 1. Lectura bidireccional de completados desde Google Tasks
    try {
      var bidirRes = syncCompletedFromTasksToSheet(ss);
      metrics.sync_err += bidirRes.errors || 0;
    } catch (errBidir) {
      metrics.sync_err++;
    }

    var incomingItems = [];

    // 2. Capa 1: Conector .ics (0 tokens, sin API Key)
    if (config.ICS_URL) {
      var icsRes = fetchAndParseIcs(config, ss);
      if (icsRes.error) {
        metrics.ics_err++;
        metrics.detalle = icsRes.error;
      } else {
        incomingItems = incomingItems.concat(icsRes.items);
      }
    }

    // 3. Capa 2: Gmail + Gemini (solo si hay API Key y modo Gmail activo)
    var processedEntries = [];
    if (config.GEMINI_API_KEY && config.GMAIL_MODE !== 'off') {
      var gmailRes = fetchPendingAcademicEmails(config, ss);
      if (gmailRes.error) {
        metrics.gmail_err++;
        metrics.detalle = gmailRes.error;
      } else {
        for (var i = 0; i < gmailRes.messages.length; i++) {
          var msg = gmailRes.messages[i];
          var aiRes = analizarTextoConGemini(
            msg.body,
            {
              origen: 'GMAIL',
              idReferencia: msg.id,
              fechaMensajeIso: msg.dateIso,
              subject: msg.subject,
              urls: msg.urls
            },
            config,
            { runStartMs: runStartMs }
          );

          if (aiRes.ok) {
            incomingItems = incomingItems.concat(aiRes.items);
            processedEntries.push({
              id: msg.id,
              origen: 'GMAIL',
              fechaIso: msg.dateIso,
              resultado: 'OK'
            });
          } else {
            metrics.gemini_err++;
            // Dead-Letter estricto: errores de cuota/red nunca se marcan como permanentes
            if (aiRes.isPermanentError) {
              metrics.dl_err++;
              processedEntries.push({
                id: msg.id,
                origen: 'GMAIL',
                fechaIso: msg.dateIso,
                resultado: 'ERROR_PERMANENTE'
              });
            }
          }
        }
      }
    }

    // 4. Motor de tareas + escritura quirúrgica en SheetStore
    if (incomingItems.length > 0) {
      var existingTasks = readAllTasks(ss);
      var merged = mergeTasksWithEngine(existingTasks, incomingItems, config, new Date());
      if (merged.patches.length > 0) {
        var writeRes = upsertTasksSurgical(merged.patches, ss);
        createdCount = writeRes.inserted;
        updatedCount = writeRes.updated;
      }
    }

    if (processedEntries.length > 0) {
      recordProcessedIds(processedEntries, ss);
    }

    // 5. Sincronización idempotente a Google Calendar y Google Tasks
    try {
      var syncRes = syncAllToCalendarAndTasks(config, ss);
      calendarSynced = syncRes.calendarSynced;
      tasksSynced = syncRes.tasksSynced;
      metrics.sync_err += syncRes.errors || 0;
    } catch (syncErr) {
      metrics.sync_err++;
      metrics.detalle = syncErr.message;
    }

    // 6. Poda >60 días y tareas semanales no bloqueantes
    pruneProcessedIds(APP_CONSTANTS.RETENTION_DAYS_PROCESSED, ss);
    try {
      sendWeeklyDigestIfDue(config, ss);
    } catch (digestErr) {
      // No degradar latido principal por resumen semanal
    }
    checkForUpdatesCached(config);

    var totalErr = metrics.ics_err + metrics.gmail_err + metrics.gemini_err + metrics.sync_err + metrics.dl_err;
    metrics.estado = totalErr > 0 ? 'DEGRADADO' : 'OK';
    writeHeartbeatNote(metrics, ss);

    return {
      ok: totalErr === 0,
      skippedLock: false,
      estado: metrics.estado,
      created: createdCount,
      updated: updatedCount,
      calendarSynced: calendarSynced,
      tasksSynced: tasksSynced,
      metrics: metrics
    };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================================
 * Endpoints puente llamados por Sidebar.html (google.script.run)
 * ========================================================================= */

/**
 * Traduce la nota técnica A1 del latido a un mensaje claro en español para el Sidebar.
 * @param {string} noteText
 * @param {string=} optTz
 * @param {Date=} optNow
 * @return {{hasSynced: boolean, ok: boolean|null, lastSyncText: string, summary: string}}
 */
function formatHeartbeatForUser(noteText, optTz, optNow) {
  var raw = String(noteText || '').trim();
  if (!raw) {
    return {
      hasSynced: false,
      ok: null,
      lastSyncText: '',
      summary: 'Todavía no hemos revisado tus tareas. Completa el Paso 1 abajo para empezar.'
    };
  }

  var stateMatch = /^estado=([A-Z_]+)/m.exec(raw);
  var tsMatch = /^timestamp=([^\r\n]+)/m.exec(raw);
  if (!tsMatch || !tsMatch[1]) {
    return {
      hasSynced: false,
      ok: null,
      lastSyncText: '',
      summary: 'Todavía no hemos revisado tus tareas. Completa el Paso 1 abajo para empezar.'
    };
  }

  var tz = optTz || getUserConfig().TIMEZONE || 'UTC';
  var tsDate = parseDateSafe(tsMatch[1].trim(), tz);
  var nowDate = optNow instanceof Date ? optNow : new Date();
  var formattedWhen = tsDate ? formatHumanSpanishDateTime_(tsDate, tz, nowDate) : tsMatch[1].trim();
  var isOk = !stateMatch || stateMatch[1] === 'OK';

  return {
    hasSynced: true,
    ok: isOk,
    lastSyncText: formattedWhen,
    summary: isOk
      ? 'Última actualización: ' + formattedWhen + ' (todo al día).'
      : 'Última actualización: ' + formattedWhen + ' (hubo un pequeño problema; puedes darle al botón para intentar de nuevo).'
  };
}

/**
 * Formatea una fecha/hora en español legible ("Hoy a las 2:45 p. m." o "25 de nov, 2:45 p. m.").
 * @private
 */
function formatHumanSpanishDateTime_(dateObj, tz, nowDate) {
  var months = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  var isoLocal = formatIsoInTimezone(dateObj, tz);
  var nowLocal = formatIsoInTimezone(nowDate || new Date(), tz);
  var m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(isoLocal);
  if (!m) return isoLocal;

  var datePart = m[1] + '-' + m[2] + '-' + m[3];
  var nowPart = nowLocal.slice(0, 10);
  var monthIdx = parseInt(m[2], 10) - 1;
  var dayNum = parseInt(m[3], 10);
  var hour24 = parseInt(m[4], 10);
  var minStr = m[5];
  var period = hour24 >= 12 ? 'p. m.' : 'a. m.';
  var hour12 = hour24 % 12;
  if (hour12 === 0) hour12 = 12;
  var timeLabel = hour12 + ':' + minStr + ' ' + period;

  if (datePart === nowPart) {
    return 'Hoy a las ' + timeLabel;
  }
  return dayNum + ' de ' + (months[monthIdx] || m[2]) + ', ' + timeLabel;
}

/**
 * Retorna estado inicial para renderizar el Sidebar.
 * @return {Object}
 */
function apiGetState() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetStructure(ss);
  var config = getUserConfig();
  var noteA1 = ss.getSheetByName(APP_CONSTANTS.SHEETS.RESUMEN).getRange('A1').getNote() || '';
  var triggerActive = hasSyncTrigger_();
  var versionInfo = checkForUpdatesCached(config);
  var humanStatus = formatHeartbeatForUser(noteA1, config.TIMEZONE);

  return {
    config: {
      TIMEZONE: config.TIMEZONE,
      ICS_URL: config.ICS_URL,
      GMAIL_MODE: config.GMAIL_MODE,
      GMAIL_LABEL: config.GMAIL_LABEL,
      HAS_GEMINI_KEY: Boolean(config.GEMINI_API_KEY),
      DIGEST_ENABLED: config.DIGEST_ENABLED
    },
    heartbeatNote: noteA1,
    humanStatus: humanStatus,
    triggerActive: triggerActive,
    versionInfo: versionInfo
  };
}

/**
 * Guarda configuración desde el Sidebar.
 * @param {Object} data
 * @return {Object}
 */
function apiSaveSettings(data) {
  var patch = {};
  if (data && Object.prototype.hasOwnProperty.call(data, 'TIMEZONE')) patch.TIMEZONE = data.TIMEZONE;
  if (data && Object.prototype.hasOwnProperty.call(data, 'ICS_URL')) patch.ICS_URL = data.ICS_URL;
  if (data && Object.prototype.hasOwnProperty.call(data, 'GMAIL_MODE')) patch.GMAIL_MODE = data.GMAIL_MODE;
  if (data && Object.prototype.hasOwnProperty.call(data, 'GMAIL_LABEL')) patch.GMAIL_LABEL = data.GMAIL_LABEL;
  if (data && Object.prototype.hasOwnProperty.call(data, 'DIGEST_ENABLED')) patch.DIGEST_ENABLED = data.DIGEST_ENABLED;
  if (data && typeof data.GEMINI_API_KEY === 'string' && data.GEMINI_API_KEY.trim() !== '') {
    patch.GEMINI_API_KEY = data.GEMINI_API_KEY.trim();
  }
  var saved = saveUserConfig(patch);
  return {
    ok: true,
    HAS_GEMINI_KEY: Boolean(saved.GEMINI_API_KEY),
    TIMEZONE: saved.TIMEZONE
  };
}

/**
 * Prueba la API Key de Gemini desde el Sidebar y la guarda si es válida.
 * @param {string} apiKey
 * @return {Object}
 */
function apiTestGemini(apiKey) {
  var keyToTest = String(apiKey || '').trim() || getUserConfig().GEMINI_API_KEY;
  var res = probarApiKeyGemini(keyToTest);
  if (res.ok && apiKey && String(apiKey).trim()) {
    saveUserConfig({ GEMINI_API_KEY: String(apiKey).trim() });
  }
  return res;
}

/**
 * Analiza un aviso pegado manualmente de WhatsApp o Chat y lo sincroniza.
 * @param {string} text
 * @return {Object}
 */
function apiProcessManualNotice(text) {
  var clean = String(text || '').trim();
  if (!clean) {
    return { ok: false, message: 'Pega primero el mensaje de Google Chat en la caja.' };
  }

  var config = getUserConfig();
  if (!config.GEMINI_API_KEY) {
    return {
      ok: false,
      message: 'Primero completa el Paso 1 con tu código gratuito de Google para poder leer avisos de Google Chat.'
    };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetStructure(ss);
  var nowIso = formatIsoInTimezone(new Date(), config.TIMEZONE);
  var aiRes = analizarTextoConGemini(
    clean,
    {
      origen: 'GOOGLE_CHAT',
      idReferencia: 'manual:' + Date.now(),
      fechaMensajeIso: nowIso,
      subject: 'Aviso pegado manualmente',
      urls: extractUrlsFromText_(clean)
    },
    config,
    { runStartMs: Date.now() }
  );

  if (!aiRes.ok) {
    return {
      ok: false,
      message: 'No pude leer ese mensaje en este momento. Intenta de nuevo en unos segundos.'
    };
  }

  var existing = readAllTasks(ss);
  var merged = mergeTasksWithEngine(existing, aiRes.items, config, new Date());
  var writeRes = upsertTasksSurgical(merged.patches, ss);
  var syncRes = syncAllToCalendarAndTasks(config, ss);

  return {
    ok: true,
    items: aiRes.items,
    stats: merged.stats,
    inserted: writeRes.inserted,
    updated: writeRes.updated,
    calendarSynced: syncRes.calendarSynced,
    tasksSynced: syncRes.tasksSynced
  };
}

/**
 * Ejecuta sincronización inmediata desde el Sidebar.
 * @return {Object}
 */
function apiRunSyncNow() {
  return ejecutarSincronizacion();
}

/**
 * Activa o desactiva el trigger automático único cada 2 horas.
 * @param {boolean} enabled
 * @return {{ok: boolean, triggerActive: boolean}}
 */
function apiToggleAutoTrigger(enabled) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === SYNC_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  if (enabled) {
    ScriptApp.newTrigger(SYNC_TRIGGER_HANDLER)
      .timeBased()
      .everyHours(2)
      .create();
  }

  return {
    ok: true,
    triggerActive: Boolean(enabled)
  };
}

/**
 * Verifica si existe un trigger activo para ejecutarSincronizacion.
 * @private
 */
function hasSyncTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === SYNC_TRIGGER_HANDLER) {
      return true;
    }
  }
  return false;
}
