/**
 * GeminiClient.gs — Cliente Gemini Flash resiliente (SRE v5.1).
 * - Descubrimiento dinámico de modelos Flash estables vía /models cacheado 24 h.
 * - API Key enviada siempre en header 'x-goog-api-key' (nunca en URL).
 * - Rotación ante errores recuperables (400 modelo, 404, 408, 429, 500, 502, 503, 504 o red)
 *   respetando margen de seguridad de 90-120 s antes del límite de 360 s de Apps Script.
 * - Dead-Letter estricto: errores de cuota/red nunca se marcan como permanentes.
 */

var GEMINI_API_BASE = 'https://generativelanguage.googleapis.com/v1beta';
var GEMINI_MODEL_CACHE_TTL_MS = 24 * 60 * 60 * 1000;
var GAS_EXECUTION_LIMIT_MS = 360 * 1000;
var GAS_SAFETY_MARGIN_MS = 105 * 1000; // 105 s de margen (rango 90-120 s)
var DEFAULT_GEMINI_FALLBACK_MODELS = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash'];

/**
 * Clasifica si un código HTTP o error de red es recuperable (permite rotar de modelo)
 * y si corresponde a cuota/infraestructura (nunca Dead-Letter permanente).
 * @param {number} statusCode
 * @param {string=} optBody
 * @return {{recoverable: boolean, isInfraOrQuota: boolean}}
 */
function classifyGeminiHttpError(statusCode, optBody) {
  if (statusCode === 0) {
    return { recoverable: true, isInfraOrQuota: true };
  }
  if (statusCode === 429 || statusCode === 408 || statusCode === 500 || statusCode === 502 || statusCode === 503 || statusCode === 504) {
    return { recoverable: true, isInfraOrQuota: true };
  }
  if (statusCode === 404) {
    return { recoverable: true, isInfraOrQuota: false };
  }
  if (statusCode === 400) {
    var txt = String(optBody || '').toLowerCase();
    var isModelErr = txt.indexOf('model') !== -1 || txt.indexOf('not found') !== -1 || txt.indexOf('not supported') !== -1;
    return { recoverable: isModelErr, isInfraOrQuota: false };
  }
  return { recoverable: false, isInfraOrQuota: false };
}

/**
 * Selecciona y ordena modelos Flash estables desde la respuesta de /models.
 * Excluye '-preview' y '-exp' salvo que no exista ningún modelo estable.
 * @param {Object[]} rawModels
 * @return {string[]}
 */
function selectStableFlashModels(rawModels) {
  var stable = [];
  var experimental = [];

  for (var i = 0; i < (rawModels || []).length; i++) {
    var m = rawModels[i];
    if (!m || !m.name) continue;
    var shortName = String(m.name).replace(/^models\//, '');
    var lower = shortName.toLowerCase();

    if (lower.indexOf('flash') === -1) continue;
    if (lower.indexOf('lite') !== -1 && lower.indexOf('8b') === -1) {
      // Permitir flash-lite solo después de flash estándar si hiciera falta
    }
    var methods = m.supportedGenerationMethods || [];
    if (methods.length > 0 && methods.indexOf('generateContent') === -1) continue;

    if (lower.indexOf('-preview') !== -1 || lower.indexOf('-exp') !== -1 || lower.indexOf('experimental') !== -1) {
      experimental.push(shortName);
    } else {
      stable.push(shortName);
    }
  }

  // Priorizar versiones más recientes primero (orden descendente lexicográfico natural)
  stable.sort().reverse();
  experimental.sort().reverse();

  var result = stable.length > 0 ? stable : experimental;
  for (var j = 0; j < DEFAULT_GEMINI_FALLBACK_MODELS.length; j++) {
    if (result.indexOf(DEFAULT_GEMINI_FALLBACK_MODELS[j]) === -1) {
      result.push(DEFAULT_GEMINI_FALLBACK_MODELS[j]);
    }
  }
  return result;
}

/**
 * Obtiene la lista de modelos Flash ordenada, usando caché de 24 h en PropertiesService.
 * @param {string} apiKey
 * @param {Function=} optFetcher
 * @return {string[]}
 */
function discoverFlashModels(apiKey, optFetcher) {
  var cacheKey = APP_CONSTANTS.PROP_KEYS.GEMINI_MODEL_CACHE;
  var props = PropertiesService.getScriptProperties();
  var rawCache = props.getProperties()[cacheKey];

  if (rawCache) {
    try {
      var parsedCache = JSON.parse(rawCache);
      if (parsedCache && Array.isArray(parsedCache.models) && parsedCache.models.length > 0 &&
          (Date.now() - Number(parsedCache.cachedAt || 0)) < GEMINI_MODEL_CACHE_TTL_MS) {
        return parsedCache.models;
      }
    } catch (e) {
      // Ignorar caché corrupta
    }
  }

  var fetcher = optFetcher || defaultGeminiFetcher_;
  try {
    var res = fetcher(GEMINI_API_BASE + '/models', {
      method: 'get',
      headers: { 'x-goog-api-key': apiKey },
      muteHttpExceptions: true
    });
    if (res.statusCode >= 200 && res.statusCode < 300) {
      var data = JSON.parse(res.body || '{}');
      var selected = selectStableFlashModels(data.models || []);
      var toStore = {};
      toStore[cacheKey] = JSON.stringify({ cachedAt: Date.now(), models: selected });
      props.setProperties(toStore, false);
      return selected;
    }
  } catch (err) {
    // Fallback a modelos por defecto
  }

  return DEFAULT_GEMINI_FALLBACK_MODELS.slice();
}

/**
 * Adaptador por defecto sobre UrlFetchApp para llamadas HTTP a Gemini.
 * @private
 */
function defaultGeminiFetcher_(url, options) {
  var response = UrlFetchApp.fetch(url, options);
  return {
    statusCode: response.getResponseCode(),
    body: response.getContentText('UTF-8')
  };
}

/**
 * Prueba una API Key de Gemini consultando /models con header x-goog-api-key.
 * @param {string} apiKey
 * @param {Function=} optFetcher
 * @return {{ok: boolean, message: string, models: string[]}}
 */
function probarApiKeyGemini(apiKey, optFetcher) {
  var cleanKey = String(apiKey || '').trim();
  if (!cleanKey) {
    return { ok: false, message: 'Pega primero tu código gratuito de Google.', models: [] };
  }

  var fetcher = optFetcher || defaultGeminiFetcher_;
  try {
    var res = fetcher(GEMINI_API_BASE + '/models', {
      method: 'get',
      headers: { 'x-goog-api-key': cleanKey },
      muteHttpExceptions: true
    });

    if (res.statusCode >= 200 && res.statusCode < 300) {
      var data = JSON.parse(res.body || '{}');
      var models = selectStableFlashModels(data.models || []);
      return {
        ok: true,
        message: '¡Código verificado y guardado! Ya puedes leer correos y avisos de Google Chat.',
        models: models
      };
    }
    if (res.statusCode === 401 || res.statusCode === 403 || res.statusCode === 400) {
      return {
        ok: false,
        message: 'Ese código no funcionó. Revisa que lo hayas copiado completo desde Google AI Studio (aistudio.google.com/apikey).',
        models: []
      };
    }
    return {
      ok: false,
      message: 'No pude verificar el código en este momento. Intenta de nuevo en unos segundos.',
      models: []
    };
  } catch (err) {
    return {
      ok: false,
      message: 'No pude conectarme para revisar el código. Revisa tu internet e intenta de nuevo.',
      models: []
    };
  }
}

/**
 * Analiza un correo o aviso pegado de WhatsApp/Chat usando Gemini Flash con rotación de modelos.
 * @param {string} texto
 * @param {Object=} metadata ({ origen, idReferencia, fechaMensajeIso, subject, urls })
 * @param {Object=} optConfig
 * @param {{fetcher?: Function, runStartMs?: number}=} optRuntime
 * @return {{items: Object[], ok: boolean, isPermanentError: boolean, error: string|null, modelUsed: string|null}}
 */
function analizarTextoConGemini(texto, metadata, optConfig, optRuntime) {
  var config = optConfig || getUserConfig();
  var apiKey = String(config.GEMINI_API_KEY || '').trim();
  if (!apiKey) {
    return {
      items: [],
      ok: false,
      isPermanentError: false,
      error: 'GEMINI_API_KEY no configurada.',
      modelUsed: null
    };
  }

  var runtime = optRuntime || {};
  var fetcher = runtime.fetcher || defaultGeminiFetcher_;
  var runStartMs = typeof runtime.runStartMs === 'number' ? runtime.runStartMs : Date.now();
  var meta = metadata || {};
  var tz = config.TIMEZONE || 'UTC';
  var msgIso = meta.fechaMensajeIso || formatIsoInTimezone(new Date(), tz);

  var models = discoverFlashModels(apiKey, fetcher);
  var prompt = buildGeminiExtractionPrompt_(texto, meta, tz, msgIso);
  var lastErr = 'Sin modelos disponibles.';
  var lastIsPermanent = false;

  for (var i = 0; i < models.length; i++) {
    var elapsed = Date.now() - runStartMs;
    if (elapsed > (GAS_EXECUTION_LIMIT_MS - GAS_SAFETY_MARGIN_MS)) {
      return {
        items: [],
        ok: false,
        isPermanentError: false,
        error: 'Margen de tiempo de Apps Script alcanzado antes de completar rotación Gemini.',
        modelUsed: null
      };
    }

    var model = models[i];
    var endpoint = GEMINI_API_BASE + '/models/' + encodeURIComponent(model) + ':generateContent';

    var res;
    try {
      res = fetcher(endpoint, {
        method: 'post',
        contentType: 'application/json',
        headers: { 'x-goog-api-key': apiKey },
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: {
            temperature: 0.1,
            responseMimeType: 'application/json'
          }
        }),
        muteHttpExceptions: true
      });
    } catch (netErr) {
      lastErr = 'Error de red con ' + model + ': ' + netErr.message;
      lastIsPermanent = false;
      continue;
    }

    if (res.statusCode >= 200 && res.statusCode < 300) {
      var parsed = parseGeminiExtractionResponse_(res.body, meta, tz, msgIso);
      if (parsed.ok) {
        return {
          items: parsed.items,
          ok: true,
          isPermanentError: false,
          error: null,
          modelUsed: model
        };
      }
      lastErr = parsed.error || 'Respuesta JSON inválida de Gemini';
      lastIsPermanent = true;
      continue;
    }

    var cls = classifyGeminiHttpError(res.statusCode, res.body);
    lastErr = 'HTTP ' + res.statusCode + ' en ' + model;
    lastIsPermanent = !cls.isInfraOrQuota;
    if (!cls.recoverable) {
      break;
    }
  }

  return {
    items: [],
    ok: false,
    isPermanentError: lastIsPermanent,
    error: lastErr,
    modelUsed: null
  };
}

/**
 * Construye el prompt estructurado para extraer tareas universitarias en JSON estricto.
 * @private
 */
function buildGeminiExtractionPrompt_(texto, meta, tz, msgIso) {
  return [
    'Eres un asistente académico que extrae tareas, entregas, exámenes, cancelaciones y aplazamientos.',
    'Zona horaria del estudiante: ' + tz,
    'Fecha de referencia del mensaje (ISO): ' + msgIso,
    'Asunto/Contexto: ' + String(meta.subject || 'Aviso manual'),
    'Texto a analizar:\n"""\n' + String(texto || '').slice(0, 6000) + '\n"""',
    'Devuelve ÚNICAMENTE un JSON válido con la estructura:',
    '{ "items": [ {',
    '  "es_pendiente": true|false,',
    '  "accion": "crear_o_actualizar"|"cancelar"|"aplazar_sin_fecha",',
    '  "materia": "Nombre de la materia o General",',
    '  "actividad": "Nombre concreto de la actividad",',
    '  "fecha_limite_iso": "YYYY-MM-DDTHH:mm:ssXXX o YYYY-MM-DD o vacío",',
    '  "urls": ["https://..."]',
    '} ] }'
  ].join('\n');
}

/**
 * Parsea la respuesta JSON de generateContent y normaliza ítems para TaskEngine.
 * @private
 */
function parseGeminiExtractionResponse_(rawBody, meta, tz, msgIso) {
  try {
    var envelope = JSON.parse(rawBody || '{}');
    var textOut = envelope &&
      envelope.candidates &&
      envelope.candidates[0] &&
      envelope.candidates[0].content &&
      envelope.candidates[0].content.parts &&
      envelope.candidates[0].content.parts[0] &&
      envelope.candidates[0].content.parts[0].text;

    if (!textOut) {
      return { ok: false, items: [], error: 'Respuesta vacía de Gemini.' };
    }

    var cleaned = String(textOut)
      .replace(/^```json\s*/i, '')
      .replace(/^```\s*/i, '')
      .replace(/\s*```$/, '')
      .trim();

    var payload = JSON.parse(cleaned);
    var rawItems = Array.isArray(payload) ? payload : (Array.isArray(payload.items) ? payload.items : []);
    var out = [];

    for (var i = 0; i < rawItems.length; i++) {
      var it = rawItems[i];
      if (!it || typeof it !== 'object') continue;

      var urlsList = [];
      if (Array.isArray(it.urls)) {
        for (var u = 0; u < it.urls.length; u++) {
          if (it.urls[u]) urlsList.push(String(it.urls[u]).trim());
        }
      }
      if (Array.isArray(meta.urls)) {
        for (var mu = 0; mu < meta.urls.length; mu++) {
          if (meta.urls[mu] && urlsList.indexOf(meta.urls[mu]) === -1) {
            urlsList.push(meta.urls[mu]);
          }
        }
      }

      var fechaLim = String(it.fecha_limite_iso || '').trim();
      out.push({
        es_pendiente: it.es_pendiente !== false,
        accion: String(it.accion || 'crear_o_actualizar').trim().toLowerCase(),
        Materia: String(it.materia || 'General').trim(),
        Actividad: String(it.actividad || '').trim(),
        Fecha_Limite: fechaLim,
        Fecha_Inicio_ISO: fechaLim,
        Origen: String(meta.origen || 'IA'),
        ID_Referencia: String(meta.idReferencia || ''),
        URLs: urlsList.join(' | '),
        Ultima_Actualizacion_ISO: msgIso
      });
    }

    return { ok: true, items: out, error: null };
  } catch (err) {
    return { ok: false, items: [], error: 'JSON inválido en respuesta Gemini: ' + err.message };
  }
}
