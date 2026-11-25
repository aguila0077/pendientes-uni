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
    .addItem('Abrir asistente', 'mostrarSidebar')
    .addItem('Sincronizar ahora', 'ejecutarSincronizacionManual')
    .addItem('Pegar aviso de WhatsApp/Chat', 'mostrarSidebar')
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
 * Ejecución manual desde el menú con feedback visual en Sheets.
 */
function ejecutarSincronizacionManual() {
  var res = ejecutarSincronizacion();
  var ui = SpreadsheetApp.getUi();
  if (res.skippedLock) {
    ui.alert('Ya hay una sincronización en curso. Intenta de nuevo en unos segundos.');
    return;
  }
  ui.alert(
    'Sincronización completada (' + res.estado + ').\n' +
    'Nuevas: ' + res.created + ' | Actualizadas: ' + res.updated +
    ' | Calendar: ' + res.calendarSynced + ' | Tasks: ' + res.tasksSynced
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
    return { ok: false, message: 'Pega primero el texto del aviso.' };
  }

  var config = getUserConfig();
  if (!config.GEMINI_API_KEY) {
    return { ok: false, message: 'Configura primero tu API Key gratuita de Gemini en el Paso 2.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetStructure(ss);
  var nowIso = formatIsoInTimezone(new Date(), config.TIMEZONE);
  var aiRes = analizarTextoConGemini(
    clean,
    {
      origen: 'WHATSAPP',
      idReferencia: 'manual:' + Date.now(),
      fechaMensajeIso: nowIso,
      subject: 'Aviso pegado manualmente',
      urls: extractUrlsFromText_(clean)
    },
    config,
    { runStartMs: Date.now() }
  );

  if (!aiRes.ok) {
    return { ok: false, message: aiRes.error || 'No se pudo analizar el aviso.' };
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
