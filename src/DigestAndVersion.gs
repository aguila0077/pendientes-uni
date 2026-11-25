/**
 * DigestAndVersion.gs — Resumen semanal por correo (3 bloques) deduplicado por semana ISO
 * y verificación semanal cacheada contra version.json en GitHub.
 */

var VERSION_CACHE_TTL_MS = 7 * 24 * 60 * 60 * 1000;

/**
 * Calcula el identificador de semana ISO (ej. "2026-W48") para una fecha dada.
 * @param {Date} dateObj
 * @return {string}
 */
function getIsoWeekKey(dateObj) {
  var d = new Date(Date.UTC(dateObj.getUTCFullYear(), dateObj.getUTCMonth(), dateObj.getUTCDate()));
  var dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return d.getUTCFullYear() + '-W' + (weekNo < 10 ? '0' + weekNo : String(weekNo));
}

/**
 * Agrupa tareas pendientes en los 3 bloques del resumen semanal:
 * 1) Para esta semana (0 a 7 días)
 * 2) Próximas semanas (>7 días)
 * 3) Sin fecha clara o atrasadas (<0 días o sin Fecha_Limite)
 * @param {Object[]} tasks
 * @param {string=} optTz
 * @param {Date=} optNow
 * @return {{thisWeek: Object[], upcoming: Object[], overdueOrUndated: Object[]}}
 */
function groupTasksForWeeklyDigest(tasks, optTz, optNow) {
  var tz = optTz || 'UTC';
  var now = optNow instanceof Date ? optNow : new Date();
  var startOfToday = new Date(now.getTime());
  startOfToday.setUTCHours(0, 0, 0, 0);

  var in7DaysMs = startOfToday.getTime() + 7 * 24 * 60 * 60 * 1000;
  var thisWeek = [];
  var upcoming = [];
  var overdueOrUndated = [];

  for (var i = 0; i < (tasks || []).length; i++) {
    var t = tasks[i];
    if (!t || t.Completada) continue;
    var estadoLower = String(t.Estado || '').toLowerCase();
    if (estadoLower === 'cancelada' || estadoLower === 'completada') continue;

    var due = parseDateSafe(t.Fecha_Limite, tz);
    if (!due) {
      overdueOrUndated.push(t);
      continue;
    }

    var dueMs = due.getTime();
    if (dueMs < startOfToday.getTime()) {
      overdueOrUndated.push(t);
    } else if (dueMs <= in7DaysMs) {
      thisWeek.push(t);
    } else {
      upcoming.push(t);
    }
  }

  return {
    thisWeek: thisWeek,
    upcoming: upcoming,
    overdueOrUndated: overdueOrUndated
  };
}

/**
 * Formatea el cuerpo del correo de resumen semanal.
 * @private
 */
function buildDigestEmailBody_(groups) {
  function formatSection(title, list) {
    if (!list || list.length === 0) {
      return title + ':\n  - Ninguna por ahora.\n';
    }
    var lines = [title + ' (' + list.length + '):'];
    for (var i = 0; i < list.length; i++) {
      var item = list[i];
      var fecha = item.Fecha_Limite ? item.Fecha_Limite : 'Sin fecha';
      lines.push('  • [' + (item.Materia || 'General') + '] ' + (item.Actividad || 'Actividad') + ' — ' + fecha);
    }
    return lines.join('\n') + '\n';
  }

  return [
    'Hola,',
    '',
    'Este es tu resumen semanal de Pendientes Uni:',
    '',
    formatSection('📅 Para esta semana', groups.thisWeek),
    formatSection('🗓️ Próximas semanas', groups.upcoming),
    formatSection('⚠️ Sin fecha clara o atrasadas', groups.overdueOrUndated),
    'Puedes marcar tus tareas como completadas directamente desde Google Tasks o en tu hoja de cálculo.',
    '— Pendientes Uni'
  ].join('\n');
}

/**
 * Envía el resumen semanal si está habilitado y no se ha enviado esta semana ISO.
 * Usa GmailApp.sendEmail (scope gmail.send ya autorizado en appsscript.json).
 * @param {Object=} optConfig
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} optSpreadsheet
 * @return {{sent: boolean, weekKey: string, reason?: string}}
 */
function sendWeeklyDigestIfDue(optConfig, optSpreadsheet) {
  var config = optConfig || getUserConfig();
  var now = new Date();
  var weekKey = getIsoWeekKey(now);

  if (!config.DIGEST_ENABLED) {
    return { sent: false, weekKey: weekKey, reason: 'disabled' };
  }

  var props = PropertiesService.getScriptProperties();
  var lastWeek = props.getProperties()[APP_CONSTANTS.PROP_KEYS.LAST_DIGEST_WEEK];
  if (lastWeek === weekKey) {
    return { sent: false, weekKey: weekKey, reason: 'already_sent' };
  }

  var tasks = readAllTasks(optSpreadsheet);
  var groups = groupTasksForWeeklyDigest(tasks, config.TIMEZONE, now);
  var totalPending = groups.thisWeek.length + groups.upcoming.length + groups.overdueOrUndated.length;
  if (totalPending === 0) {
    return { sent: false, weekKey: weekKey, reason: 'no_pending_tasks' };
  }

  var recipient = Session.getActiveUser().getEmail();
  if (!recipient) {
    return { sent: false, weekKey: weekKey, reason: 'no_user_email' };
  }

  var subject = '🎓 Pendientes Uni — Resumen semanal (' + weekKey + ')';
  var body = buildDigestEmailBody_(groups);
  GmailApp.sendEmail(recipient, subject, body);

  var toSave = {};
  toSave[APP_CONSTANTS.PROP_KEYS.LAST_DIGEST_WEEK] = weekKey;
  props.setProperties(toSave, false);

  return { sent: true, weekKey: weekKey };
}

/**
 * Compara dos versiones semver simples (ej. "1.1.0" > "1.0.0").
 * @param {string} latest
 * @param {string} current
 * @return {boolean}
 */
function isNewerVersion(latest, current) {
  var a = String(latest || '0.0.0').replace(/^v/i, '').split('.').map(Number);
  var b = String(current || '0.0.0').replace(/^v/i, '').split('.').map(Number);
  for (var i = 0; i < 3; i++) {
    var va = isNaN(a[i]) ? 0 : a[i];
    var vb = isNaN(b[i]) ? 0 : b[i];
    if (va > vb) return true;
    if (va < vb) return false;
  }
  return false;
}

/**
 * Verifica actualizaciones contra version.json con caché semanal de 7 días.
 * Nunca bloquea ni lanza errores si no hay red o no hay URL configurada.
 * @param {Object=} optConfig
 * @return {{hasUpdate: boolean, latestVersion: string, currentVersion: string, highlights: string[], templateUrl: string}|null}
 */
function checkForUpdatesCached(optConfig) {
  var config = optConfig || getUserConfig();
  var currentVersion = APP_CONSTANTS.APP_VERSION;
  var props = PropertiesService.getScriptProperties();
  var cacheKey = APP_CONSTANTS.PROP_KEYS.VERSION_CACHE;
  var rawCache = props.getProperties()[cacheKey];

  if (rawCache) {
    try {
      var cached = JSON.parse(rawCache);
      if (cached && (Date.now() - Number(cached.checkedAt || 0)) < VERSION_CACHE_TTL_MS && cached.payload) {
        return cached.payload;
      }
    } catch (e) {
      // Ignorar caché inválida
    }
  }

  if (!config.VERSION_CHECK_URL) {
    return {
      hasUpdate: false,
      latestVersion: currentVersion,
      currentVersion: currentVersion,
      highlights: [],
      templateUrl: ''
    };
  }

  try {
    var res = UrlFetchApp.fetch(config.VERSION_CHECK_URL, {
      muteHttpExceptions: true,
      followRedirects: true
    });
    if (res.getResponseCode() >= 200 && res.getResponseCode() < 300) {
      var data = JSON.parse(res.getContentText('UTF-8') || '{}');
      var latest = String(data.latestVersion || currentVersion);
      var payload = {
        hasUpdate: isNewerVersion(latest, currentVersion),
        latestVersion: latest,
        currentVersion: currentVersion,
        highlights: Array.isArray(data.highlights) ? data.highlights : [],
        templateUrl: String(data.templateUrl || '')
      };
      var toSave = {};
      toSave[cacheKey] = JSON.stringify({ checkedAt: Date.now(), payload: payload });
      props.setProperties(toSave, false);
      return payload;
    }
  } catch (err) {
    // No bloquear ejecución si no hay red
  }

  return null;
}
