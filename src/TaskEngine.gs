/**
 * TaskEngine.gs — Motor determinista de claves semánticas, fusión .ics + IA,
 * evaluación previa de cancelar/aplazar_sin_fecha y CHRONO_GUARD.
 */

/**
 * Normaliza un texto para construir claves semánticas deterministas.
 * @param {string} str
 * @return {string}
 */
function normalizeSemanticToken_(str) {
  return String(str || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

/**
 * Genera una Clave_Semantica determinista a partir de Materia y Actividad.
 * @param {string} materia
 * @param {string} actividad
 * @return {string}
 */
function buildSemanticKey(materia, actividad) {
  var m = normalizeSemanticToken_(materia) || 'general';
  var a = normalizeSemanticToken_(actividad) || 'actividad';
  return m + '::' + a;
}

/**
 * Deduplica y combina listas de URLs separadas por ' | '.
 * @param {string} existingUrls
 * @param {string} incomingUrls
 * @return {string}
 */
function mergeDeduplicatedUrls_(existingUrls, incomingUrls) {
  var list = [];
  var rawParts = []
    .concat(String(existingUrls || '').split(/\s*\|\s*|\s+/))
    .concat(String(incomingUrls || '').split(/\s*\|\s*|\s+/));

  for (var i = 0; i < rawParts.length; i++) {
    var token = rawParts[i].trim();
    if (token && list.indexOf(token) === -1) {
      list.push(token);
    }
  }
  return list.join(' | ');
}

/**
 * Busca una tarea existente por Clave_Semantica o por ID_Referencia.
 * @private
 */
function findExistingTask_(existingList, incomingItem, semanticKey) {
  for (var i = 0; i < existingList.length; i++) {
    if (existingList[i].Clave_Semantica === semanticKey) {
      return existingList[i];
    }
  }
  var refId = String(incomingItem.ID_Referencia || incomingItem.uid || '').trim();
  var isIcs = String(incomingItem.Origen || '').toUpperCase() === 'ICS';
  if (refId && isIcs) {
    for (var j = 0; j < existingList.length; j++) {
      if (String(existingList[j].ID_Referencia || '').trim() === refId) {
        return existingList[j];
      }
    }
  }
  return null;
}

/**
 * Procesa y fusiona ítems entrantes (.ics o IA) contra las tareas existentes en Sheets.
 * Reglas invariantes:
 * 1) Evaluar `accion` (`cancelar` / `aplazar_sin_fecha`) ANTES de `es_pendiente`;
 *    nunca crear filas huérfanas para cancelaciones sin tarea existente.
 * 2) CHRONO_GUARD: si `msgEsMasAntiguo === true`, jamás sobrescribir Materia, Actividad,
 *    Origen ni ID_Referencia (ni fechas/estado); solo anexar URLs nuevas deduplicadas.
 * 3) Preservar siempre Completada y Notas_Personales.
 *
 * @param {Object[]} existingTasks
 * @param {Object[]} incomingItems
 * @param {Object=} optConfig
 * @param {Date=} optNow
 * @return {{patches: Object[], stats: {created: number, updated: number, cancelled: number, postponed: number, chronoBlocked: number, orphanCancelledSkipped: number}}}
 */
function mergeTasksWithEngine(existingTasks, incomingItems, optConfig, optNow) {
  var config = optConfig || { TIMEZONE: 'UTC' };
  var tz = config.TIMEZONE || 'UTC';
  var nowIso = formatIsoInTimezone(optNow instanceof Date ? optNow : new Date(), tz);

  var workingMap = {};
  var workingList = [];
  for (var i = 0; i < (existingTasks || []).length; i++) {
    var clone = Object.assign({}, existingTasks[i]);
    workingMap[clone.Clave_Semantica] = clone;
    workingList.push(clone);
  }

  var patchesByKey = {};
  var stats = {
    created: 0,
    updated: 0,
    cancelled: 0,
    postponed: 0,
    chronoBlocked: 0,
    orphanCancelledSkipped: 0
  };

  for (var j = 0; j < (incomingItems || []).length; j++) {
    var item = incomingItems[j];
    if (!item) continue;

    var accion = String(item.accion || 'crear_o_actualizar').trim().toLowerCase();
    var materia = String(item.Materia || item.materia || 'General').trim();
    var actividad = String(item.Actividad || item.actividad || '').trim();
    var semanticKey = item.Clave_Semantica || buildSemanticKey(materia, actividad);
    var existing = findExistingTask_(workingList, item, semanticKey);

    var incomingUpdateDate = parseDateSafe(item.Ultima_Actualizacion_ISO || item.fecha_mensaje_iso, tz);
    var existingUpdateDate = existing ? parseDateSafe(existing.Ultima_Actualizacion_ISO, tz) : null;
    var msgEsMasAntiguo = Boolean(
      existing &&
      incomingUpdateDate &&
      existingUpdateDate &&
      incomingUpdateDate.getTime() < existingUpdateDate.getTime()
    );

    // 1. Evaluar cancelación / aplazamiento ANTES de filtrar por es_pendiente
    if (accion === 'cancelar') {
      if (!existing) {
        stats.orphanCancelledSkipped++;
        continue;
      }
      if (msgEsMasAntiguo) {
        var mergedUrlsCancel = mergeDeduplicatedUrls_(existing.URLs, item.URLs);
        if (mergedUrlsCancel !== existing.URLs) {
          existing.URLs = mergedUrlsCancel;
          patchesByKey[existing.Clave_Semantica] = Object.assign({}, existing);
        }
        stats.chronoBlocked++;
        continue;
      }

      existing.Estado = 'Cancelada';
      existing.URLs = mergeDeduplicatedUrls_(existing.URLs, item.URLs);
      existing.Ultima_Actualizacion_ISO = item.Ultima_Actualizacion_ISO || nowIso;
      patchesByKey[existing.Clave_Semantica] = Object.assign({}, existing);
      stats.cancelled++;
      continue;
    }

    if (accion === 'aplazar_sin_fecha') {
      if (!existing) {
        stats.orphanCancelledSkipped++;
        continue;
      }
      if (msgEsMasAntiguo) {
        var mergedUrlsPostpone = mergeDeduplicatedUrls_(existing.URLs, item.URLs);
        if (mergedUrlsPostpone !== existing.URLs) {
          existing.URLs = mergedUrlsPostpone;
          patchesByKey[existing.Clave_Semantica] = Object.assign({}, existing);
        }
        stats.chronoBlocked++;
        continue;
      }

      existing.Estado = 'Aplazada / Sin fecha';
      existing.Fecha_Limite = '';
      existing.Fecha_Inicio_ISO = '';
      existing.URLs = mergeDeduplicatedUrls_(existing.URLs, item.URLs);
      existing.Ultima_Actualizacion_ISO = item.Ultima_Actualizacion_ISO || nowIso;
      patchesByKey[existing.Clave_Semantica] = Object.assign({}, existing);
      stats.postponed++;
      continue;
    }

    // 2. Filtrar por es_pendiente solo después de evaluar cancelar / aplazar_sin_fecha
    if (item.es_pendiente === false) {
      continue;
    }

    if (!actividad) continue;

    // 3. CHRONO_GUARD en actualización sobre tarea existente
    if (existing && msgEsMasAntiguo) {
      var urlsOnly = mergeDeduplicatedUrls_(existing.URLs, item.URLs);
      if (urlsOnly !== existing.URLs) {
        existing.URLs = urlsOnly;
        patchesByKey[existing.Clave_Semantica] = Object.assign({}, existing);
      }
      stats.chronoBlocked++;
      continue;
    }

    if (!existing) {
      var createdObj = {
        Clave_Semantica: semanticKey,
        Materia: materia,
        Actividad: actividad,
        Fecha_Limite: String(item.Fecha_Limite || ''),
        Fecha_Inicio_ISO: String(item.Fecha_Inicio_ISO || ''),
        Estado: String(item.Estado || 'Pendiente'),
        Completada: false,
        Origen: String(item.Origen || 'ICS'),
        ID_Referencia: String(item.ID_Referencia || item.uid || ''),
        URLs: mergeDeduplicatedUrls_('', item.URLs),
        Ultima_Actualizacion_ISO: String(item.Ultima_Actualizacion_ISO || nowIso),
        Calendar_Event_ID: '',
        Tasks_ID: '',
        Notas_Personales: ''
      };
      workingMap[semanticKey] = createdObj;
      workingList.push(createdObj);
      patchesByKey[semanticKey] = Object.assign({}, createdObj);
      stats.created++;
    } else {
      existing.Materia = materia || existing.Materia;
      existing.Actividad = actividad || existing.Actividad;
      existing.Fecha_Limite = item.Fecha_Limite !== undefined ? String(item.Fecha_Limite) : existing.Fecha_Limite;
      existing.Fecha_Inicio_ISO = item.Fecha_Inicio_ISO !== undefined ? String(item.Fecha_Inicio_ISO) : existing.Fecha_Inicio_ISO;
      existing.Estado = item.Estado || (existing.Completada ? 'Completada' : 'Pendiente');
      existing.Origen = combineOrigins_(existing.Origen, item.Origen);
      if (item.ID_Referencia || item.uid) {
        existing.ID_Referencia = String(item.ID_Referencia || item.uid);
      }
      existing.URLs = mergeDeduplicatedUrls_(existing.URLs, item.URLs);
      existing.Ultima_Actualizacion_ISO = String(item.Ultima_Actualizacion_ISO || nowIso);

      patchesByKey[existing.Clave_Semantica] = Object.assign({}, existing);
      stats.updated++;
    }
  }

  var patches = [];
  var keys = Object.keys(patchesByKey);
  for (var k = 0; k < keys.length; k++) {
    patches.push(patchesByKey[keys[k]]);
  }

  return {
    patches: patches,
    stats: stats
  };
}

/**
 * Combina orígenes sin duplicarlos (ej. 'ICS+IA').
 * @private
 */
function combineOrigins_(prevOrigin, nextOrigin) {
  var parts = [];
  var all = String(prevOrigin || '').split('+').concat(String(nextOrigin || '').split('+'));
  for (var i = 0; i < all.length; i++) {
    var p = all[i].trim();
    if (p && parts.indexOf(p) === -1) {
      parts.push(p);
    }
  }
  return parts.join('+');
}
