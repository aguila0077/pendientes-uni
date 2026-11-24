/**
 * CalendarTasksSync.gs — Sincronización idempotente a Google Calendar y Google Tasks
 * (calendario y lista dedicados "🎓 Pendientes Uni") y sincronización bidireccional de completado [✅].
 */

var COMPLETED_PREFIX = '[✅] ';

/**
 * Formatea el título obligatorio para Google Tasks:
 * "[HH:mm] Actividad — Materia" o "[Sin hora] Actividad — Materia".
 * @param {string} actividad
 * @param {string} materia
 * @param {string} fechaLimiteVal
 * @param {string=} optTz
 * @return {string}
 */
function formatTaskTitle(actividad, materia, fechaLimiteVal, optTz) {
  var act = String(actividad || 'Actividad').trim();
  var mat = String(materia || 'General').trim();
  var timeTag = '[Sin hora]';

  var raw = String(fechaLimiteVal || '').trim();
  var hasExplicitTime = /T\d{2}:\d{2}|\s\d{2}:\d{2}/.test(raw) || Object.prototype.toString.call(fechaLimiteVal) === '[object Date]';
  if (raw && hasExplicitTime) {
    var d = parseDateSafe(fechaLimiteVal, optTz);
    if (d) {
      var tz = optTz || (typeof getUserConfig === 'function' ? getUserConfig().TIMEZONE : 'UTC');
      var hhmm = formatHHmmInTimezone_(d, tz);
      if (hhmm) {
        timeTag = '[' + hhmm + ']';
      }
    }
  }

  return timeTag + ' ' + act + ' — ' + mat;
}

/**
 * Formatea HH:mm en la zona horaria indicada.
 * @private
 */
function formatHHmmInTimezone_(dateObj, tz) {
  if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.formatDate === 'function') {
    return Utilities.formatDate(dateObj, tz || 'UTC', 'HH:mm');
  }
  return new Intl.DateTimeFormat('en-GB', {
    timeZone: tz || 'UTC',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false
  }).format(dateObj);
}

/**
 * Antepone '[✅] ' al título de un evento de Calendar de forma idempotente.
 * @param {string} title
 * @param {boolean} isCompleted
 * @return {string}
 */
function applyCompletionPrefixToTitle(title, isCompleted) {
  var clean = String(title || '').trim();
  if (isCompleted) {
    return clean.indexOf(COMPLETED_PREFIX) === 0 ? clean : COMPLETED_PREFIX + clean;
  }
  return clean.indexOf(COMPLETED_PREFIX) === 0 ? clean.slice(COMPLETED_PREFIX.length) : clean;
}

/**
 * Reconcilia el estado de completado entre el Sheet y Google Tasks (bidireccional).
 * @param {Object} sheetTask
 * @param {Object|null} remoteGoogleTask
 * @return {{completed: boolean, sheetNeedsUpdate: boolean, tasksNeedsUpdate: boolean, calendarNeedsPrefix: boolean}}
 */
function reconcileCompletionState(sheetTask, remoteGoogleTask) {
  var sheetDone = Boolean(sheetTask && sheetTask.Completada);
  var remoteDone = Boolean(remoteGoogleTask && remoteGoogleTask.status === 'completed');
  var finalDone = sheetDone || remoteDone;

  return {
    completed: finalDone,
    sheetNeedsUpdate: finalDone && !sheetDone,
    tasksNeedsUpdate: finalDone && remoteGoogleTask ? !remoteDone : false,
    calendarNeedsPrefix: finalDone
  };
}

/**
 * Obtiene o crea de forma idempotente el calendario dedicado "🎓 Pendientes Uni".
 * @return {GoogleAppsScript.Calendar.Calendar}
 */
function getOrCreatePendientesCalendar() {
  var name = APP_CONSTANTS.CALENDAR_NAME;
  var calendars = CalendarApp.getCalendarsByName(name);
  if (calendars && calendars.length > 0) {
    return calendars[0];
  }
  var tz = getUserConfig().TIMEZONE;
  return CalendarApp.createCalendar(name, { timeZone: tz });
}

/**
 * Obtiene o crea de forma idempotente la lista dedicada "🎓 Pendientes Uni" en Google Tasks.
 * @return {string} ID de la lista de tareas.
 */
function getOrCreatePendientesTaskListId() {
  var listName = APP_CONSTANTS.TASKS_LIST_NAME;
  var res = Tasks.Tasklists.list({ maxResults: 100 });
  var items = (res && res.items) ? res.items : [];
  for (var i = 0; i < items.length; i++) {
    if (items[i].title === listName) {
      return items[i].id;
    }
  }
  var created = Tasks.Tasklists.insert({ title: listName });
  return created.id;
}

/**
 * Paso 1 de sincronización: lee completados desde Google Tasks y los propaga al Sheet y Calendar.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} optSpreadsheet
 * @return {{completedFromTasks: number, errors: number}}
 */
function syncCompletedFromTasksToSheet(optSpreadsheet) {
  var tasks = readAllTasks(optSpreadsheet);
  if (tasks.length === 0) return { completedFromTasks: 0, errors: 0 };

  var taskListId = getOrCreatePendientesTaskListId();
  var calendar = getOrCreatePendientesCalendar();
  var remoteMap = fetchTasksMap_(taskListId);

  var patches = [];
  var completedCount = 0;
  var errors = 0;

  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    var remote = t.Tasks_ID ? remoteMap[t.Tasks_ID] : null;
    var state = reconcileCompletionState(t, remote);

    if (state.sheetNeedsUpdate) {
      patches.push({
        Clave_Semantica: t.Clave_Semantica,
        Completada: true,
        _forceCompletada: true,
        Estado: 'Completada'
      });
      completedCount++;
    }

    if (state.calendarNeedsPrefix && t.Calendar_Event_ID) {
      try {
        var ev = calendar.getEventById(t.Calendar_Event_ID);
        if (ev) {
          var nextTitle = applyCompletionPrefixToTitle(ev.getTitle(), true);
          if (ev.getTitle() !== nextTitle) {
            ev.setTitle(nextTitle);
          }
          ev.removeAllReminders();
        }
      } catch (err) {
        errors++;
      }
    }
  }

  if (patches.length > 0) {
    upsertTasksSurgical(patches, optSpreadsheet);
  }

  return { completedFromTasks: completedCount, errors: errors };
}

/**
 * Sincroniza todas las tareas del Sheet hacia Google Calendar y Google Tasks de forma idempotente.
 * @param {Object=} optConfig
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} optSpreadsheet
 * @return {{calendarSynced: number, tasksSynced: number, errors: number}}
 */
function syncAllToCalendarAndTasks(optConfig, optSpreadsheet) {
  var config = optConfig || getUserConfig();
  var tz = config.TIMEZONE || 'UTC';
  var tasks = readAllTasks(optSpreadsheet);
  if (tasks.length === 0) return { calendarSynced: 0, tasksSynced: 0, errors: 0 };

  var calendar = getOrCreatePendientesCalendar();
  var taskListId = getOrCreatePendientesTaskListId();
  var remoteMap = fetchTasksMap_(taskListId);

  var idPatches = [];
  var calendarSynced = 0;
  var tasksSynced = 0;
  var errors = 0;

  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    var isCancelled = String(t.Estado || '').toLowerCase() === 'cancelada';
    var isPostponed = String(t.Estado || '').toLowerCase().indexOf('aplazada') === 0;
    var isDone = Boolean(t.Completada);

    var calId = String(t.Calendar_Event_ID || '').trim();
    var gTaskId = String(t.Tasks_ID || '').trim();
    var needsPatch = false;

    // 1. Google Calendar
    try {
      var baseTitle = t.Actividad + ' — ' + t.Materia;
      if (isCancelled) {
        baseTitle = '[CANCELADA] ' + baseTitle;
      } else if (isDone) {
        baseTitle = applyCompletionPrefixToTitle(baseTitle, true);
      }

      var dueDate = parseDateSafe(t.Fecha_Limite, tz);
      var startDate = parseDateSafe(t.Fecha_Inicio_ISO, tz) || dueDate;

      if (dueDate && !isPostponed) {
        var ev = calId ? safeGetEvent_(calendar, calId) : null;
        var isAllDay = /^\d{4}-\d{2}-\d{2}$/.test(String(t.Fecha_Limite || '').trim());
        var evEnd = dueDate;
        var evStart = startDate && startDate.getTime() <= evEnd.getTime()
          ? startDate
          : new Date(evEnd.getTime() - 30 * 60 * 1000);

        if (!ev) {
          if (isAllDay) {
            ev = calendar.createAllDayEvent(baseTitle, dueDate, { description: buildEventDescription_(t) });
          } else {
            ev = calendar.createEvent(baseTitle, evStart, evEnd, { description: buildEventDescription_(t) });
          }
          calId = ev.getId();
          needsPatch = true;
        } else {
          if (ev.getTitle() !== baseTitle) ev.setTitle(baseTitle);
          ev.setDescription(buildEventDescription_(t));
          if (!isAllDay) {
            ev.setTime(evStart, evEnd);
          }
        }

        if (isDone || isCancelled) {
          ev.removeAllReminders();
        }
        calendarSynced++;
      }
    } catch (calErr) {
      errors++;
    }

    // 2. Google Tasks
    try {
      var taskTitle = formatTaskTitle(t.Actividad, t.Materia, t.Fecha_Limite, tz);
      if (isCancelled) {
        taskTitle = '[CANCELADA] ' + taskTitle;
      }
      var dueIsoTasks = toGoogleTasksDueIso_(t.Fecha_Limite, tz);
      var taskPayload = {
        title: taskTitle,
        notes: buildEventDescription_(t),
        status: (isDone || isCancelled) ? 'completed' : 'needsAction'
      };
      if (dueIsoTasks) {
        taskPayload.due = dueIsoTasks;
      }

      var existingRemote = gTaskId ? remoteMap[gTaskId] : null;
      if (!existingRemote) {
        var inserted = Tasks.Tasks.insert(taskPayload, taskListId);
        gTaskId = inserted.id;
        needsPatch = true;
      } else {
        taskPayload.id = existingRemote.id;
        Tasks.Tasks.patch(taskPayload, taskListId, existingRemote.id);
      }
      tasksSynced++;
    } catch (taskErr) {
      errors++;
    }

    if (needsPatch) {
      idPatches.push({
        Clave_Semantica: t.Clave_Semantica,
        Calendar_Event_ID: calId,
        Tasks_ID: gTaskId
      });
    }
  }

  if (idPatches.length > 0) {
    upsertTasksSurgical(idPatches, optSpreadsheet);
  }

  return {
    calendarSynced: calendarSynced,
    tasksSynced: tasksSynced,
    errors: errors
  };
}

/**
 * Obtiene un mapa id -> Task de Google Tasks incluyendo completadas.
 * @private
 */
function fetchTasksMap_(taskListId) {
  var map = {};
  var pageToken = null;
  do {
    var res = Tasks.Tasks.list(taskListId, {
      showCompleted: true,
      showHidden: true,
      maxResults: 100,
      pageToken: pageToken || undefined
    });
    var items = (res && res.items) ? res.items : [];
    for (var i = 0; i < items.length; i++) {
      map[items[i].id] = items[i];
    }
    pageToken = res && res.nextPageToken ? res.nextPageToken : null;
  } while (pageToken);
  return map;
}

/**
 * Busca evento de Calendar sin lanzar excepción si fue borrado.
 * @private
 */
function safeGetEvent_(calendar, eventId) {
  try {
    return calendar.getEventById(eventId);
  } catch (e) {
    return null;
  }
}

/**
 * Construye la descripción legible para Calendar y Tasks.
 * @private
 */
function buildEventDescription_(task) {
  var lines = [
    'Materia: ' + (task.Materia || 'General'),
    'Estado: ' + (task.Estado || 'Pendiente'),
    'Origen: ' + (task.Origen || '')
  ];
  if (task.URLs) lines.push('Enlaces: ' + task.URLs);
  if (task.Notas_Personales) lines.push('Notas: ' + task.Notas_Personales);
  return lines.join('\n');
}

/**
 * Convierte Fecha_Limite al formato RFC3339 que espera Google Tasks API.
 * @private
 */
function toGoogleTasksDueIso_(fechaVal, tz) {
  var d = parseDateSafe(fechaVal, tz);
  if (!d) return null;
  var ymd = formatAllDayDate_(d, tz);
  return ymd + 'T00:00:00.000Z';
}
