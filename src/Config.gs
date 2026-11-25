/**
 * Config.gs — Configuración centralizada y persistencia en PropertiesService.
 * Sin valores hardcodeados de usuario, universidad o hoja de cálculo.
 */

var APP_CONSTANTS = {
  APP_NAME: 'Pendientes Uni',
  APP_VERSION: '1.0.0',
  CALENDAR_NAME: '🎓 Pendientes Uni',
  TASKS_LIST_NAME: '🎓 Pendientes Uni',
  SHEETS: {
    RESUMEN: 'Resumen_Tareas',
    MATERIAS: 'Config_Materias',
    IDS_PROCESADOS: 'IDs_Procesados'
  },
  GMAIL_MODES: ['label', 'institutional', 'off'],
  EDTECH_DOMAINS: [
    'eltngl.com',
    'ngl.cengage.com',
    'moodle',
    'canvas',
    'blackboard',
    'brightspace'
  ],
  PROP_KEYS: {
    TIMEZONE: 'PU_TIMEZONE',
    ICS_URL: 'PU_ICS_URL',
    GMAIL_MODE: 'PU_GMAIL_MODE',
    GMAIL_LABEL: 'PU_GMAIL_LABEL',
    GEMINI_API_KEY: 'PU_GEMINI_API_KEY',
    DIGEST_ENABLED: 'PU_DIGEST_ENABLED',
    VERSION_CHECK_URL: 'PU_VERSION_CHECK_URL',
    GEMINI_MODEL_CACHE: 'PU_GEMINI_MODEL_CACHE',
    VERSION_CACHE: 'PU_VERSION_CACHE',
    LAST_DIGEST_WEEK: 'PU_LAST_DIGEST_WEEK'
  },
  RETENTION_DAYS_PROCESSED: 60
};

/**
 * Obtiene la zona horaria por defecto del script de forma segura.
 * @return {string}
 */
function getDefaultTimezone_() {
  try {
    if (typeof Session !== 'undefined' && Session && typeof Session.getScriptTimeZone === 'function') {
      var tz = Session.getScriptTimeZone();
      if (tz && String(tz).trim()) {
        return String(tz).trim();
      }
    }
  } catch (e) {
    // Fallback seguro para entornos sin Session
  }
  return 'UTC';
}

/**
 * Valores por defecto de configuración de usuario.
 * @return {Object}
 */
function getDefaultConfig_() {
  return {
    TIMEZONE: getDefaultTimezone_(),
    ICS_URL: '',
    GMAIL_MODE: 'label',
    GMAIL_LABEL: 'Universidad',
    GEMINI_API_KEY: '',
    DIGEST_ENABLED: true,
    VERSION_CHECK_URL: 'https://raw.githubusercontent.com/aguila0077/pendientes-uni/main/version.json'
  };
}

/**
 * Normaliza un valor booleano leído desde PropertiesService o UI.
 * @param {*} val
 * @param {boolean} fallback
 * @return {boolean}
 */
function parseBooleanConfig_(val, fallback) {
  if (typeof val === 'boolean') return val;
  if (val === null || val === undefined || val === '') return fallback;
  var s = String(val).trim().toLowerCase();
  if (s === 'true' || s === '1' || s === 'yes' || s === 'si') return true;
  if (s === 'false' || s === '0' || s === 'no') return false;
  return fallback;
}

/**
 * Valida que una zona horaria IANA sea utilizable por Utilities.formatDate o Intl.
 * @param {string} tz
 * @return {boolean}
 */
function isValidTimezone_(tz) {
  if (!tz || typeof tz !== 'string' || !tz.trim()) return false;
  var clean = tz.trim();
  try {
    if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.formatDate === 'function') {
      Utilities.formatDate(new Date(), clean, 'yyyy-MM-dd');
      return true;
    }
    if (typeof Intl !== 'undefined' && Intl.DateTimeFormat) {
      Intl.DateTimeFormat('en-US', { timeZone: clean });
      return true;
    }
  } catch (e) {
    return false;
  }
  return true;
}

/**
 * Lee y valida la configuración actual del usuario desde PropertiesService.
 * @return {Object}
 */
function getUserConfig() {
  var defaults = getDefaultConfig_();
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperties();
  var keys = APP_CONSTANTS.PROP_KEYS;

  var tz = raw[keys.TIMEZONE] ? String(raw[keys.TIMEZONE]).trim() : defaults.TIMEZONE;
  if (!isValidTimezone_(tz)) {
    tz = defaults.TIMEZONE;
  }

  var gmailMode = raw[keys.GMAIL_MODE] ? String(raw[keys.GMAIL_MODE]).trim().toLowerCase() : defaults.GMAIL_MODE;
  if (APP_CONSTANTS.GMAIL_MODES.indexOf(gmailMode) === -1) {
    gmailMode = defaults.GMAIL_MODE;
  }

  var gmailLabel = raw[keys.GMAIL_LABEL] !== undefined && String(raw[keys.GMAIL_LABEL]).trim() !== ''
    ? String(raw[keys.GMAIL_LABEL]).trim()
    : defaults.GMAIL_LABEL;

  var icsUrl = raw[keys.ICS_URL] ? String(raw[keys.ICS_URL]).trim() : defaults.ICS_URL;
  var apiKey = raw[keys.GEMINI_API_KEY] ? String(raw[keys.GEMINI_API_KEY]).trim() : defaults.GEMINI_API_KEY;
  var digestEnabled = parseBooleanConfig_(raw[keys.DIGEST_ENABLED], defaults.DIGEST_ENABLED);
  var versionUrl = raw[keys.VERSION_CHECK_URL] ? String(raw[keys.VERSION_CHECK_URL]).trim() : defaults.VERSION_CHECK_URL;

  return {
    TIMEZONE: tz,
    ICS_URL: icsUrl,
    GMAIL_MODE: gmailMode,
    GMAIL_LABEL: gmailLabel,
    GEMINI_API_KEY: apiKey,
    DIGEST_ENABLED: digestEnabled,
    VERSION_CHECK_URL: versionUrl
  };
}

/**
 * Valida y aplica un parche parcial sobre la configuración persistida.
 * @param {Object} patch
 * @return {Object} Configuración resultante actualizada.
 */
function saveUserConfig(patch) {
  if (!patch || typeof patch !== 'object') {
    throw new Error('saveUserConfig requiere un objeto patch válido.');
  }

  var current = getUserConfig();
  var next = {
    TIMEZONE: current.TIMEZONE,
    ICS_URL: current.ICS_URL,
    GMAIL_MODE: current.GMAIL_MODE,
    GMAIL_LABEL: current.GMAIL_LABEL,
    GEMINI_API_KEY: current.GEMINI_API_KEY,
    DIGEST_ENABLED: current.DIGEST_ENABLED,
    VERSION_CHECK_URL: current.VERSION_CHECK_URL
  };

  if (Object.prototype.hasOwnProperty.call(patch, 'TIMEZONE')) {
    var tz = String(patch.TIMEZONE || '').trim();
    if (!isValidTimezone_(tz)) {
      throw new Error('Zona horaria inválida: ' + tz);
    }
    next.TIMEZONE = tz;
  }

  if (Object.prototype.hasOwnProperty.call(patch, 'ICS_URL')) {
    var url = String(patch.ICS_URL || '').trim();
    if (url && !/^https?:\/\/.+/i.test(url)) {
      throw new Error('ICS_URL debe iniciar con http:// o https://');
    }
    next.ICS_URL = url;
  }

  if (Object.prototype.hasOwnProperty.call(patch, 'GMAIL_MODE')) {
    var mode = String(patch.GMAIL_MODE || '').trim().toLowerCase();
    if (APP_CONSTANTS.GMAIL_MODES.indexOf(mode) === -1) {
      throw new Error('GMAIL_MODE inválido: ' + mode + '. Valores permitidos: label, institutional, off');
    }
    next.GMAIL_MODE = mode;
  }

  if (Object.prototype.hasOwnProperty.call(patch, 'GMAIL_LABEL')) {
    var label = String(patch.GMAIL_LABEL || '').trim();
    next.GMAIL_LABEL = label || 'Universidad';
  }

  if (Object.prototype.hasOwnProperty.call(patch, 'GEMINI_API_KEY')) {
    next.GEMINI_API_KEY = String(patch.GEMINI_API_KEY || '').trim();
  }

  if (Object.prototype.hasOwnProperty.call(patch, 'DIGEST_ENABLED')) {
    next.DIGEST_ENABLED = parseBooleanConfig_(patch.DIGEST_ENABLED, true);
  }

  if (Object.prototype.hasOwnProperty.call(patch, 'VERSION_CHECK_URL')) {
    var vUrl = String(patch.VERSION_CHECK_URL || '').trim();
    if (vUrl && !/^https?:\/\/.+/i.test(vUrl)) {
      throw new Error('VERSION_CHECK_URL debe iniciar con http:// o https://');
    }
    next.VERSION_CHECK_URL = vUrl;
  }

  var keys = APP_CONSTANTS.PROP_KEYS;
  var toSave = {};
  toSave[keys.TIMEZONE] = next.TIMEZONE;
  toSave[keys.ICS_URL] = next.ICS_URL;
  toSave[keys.GMAIL_MODE] = next.GMAIL_MODE;
  toSave[keys.GMAIL_LABEL] = next.GMAIL_LABEL;
  toSave[keys.GEMINI_API_KEY] = next.GEMINI_API_KEY;
  toSave[keys.DIGEST_ENABLED] = next.DIGEST_ENABLED ? 'true' : 'false';
  toSave[keys.VERSION_CHECK_URL] = next.VERSION_CHECK_URL;

  PropertiesService.getScriptProperties().setProperties(toSave, false);
  return next;
}

/**
 * Restablece la configuración de usuario a los valores por defecto.
 * @return {Object} Configuración restablecida.
 */
function resetConfig() {
  var props = PropertiesService.getScriptProperties();
  var keys = APP_CONSTANTS.PROP_KEYS;
  var keyList = [
    keys.TIMEZONE,
    keys.ICS_URL,
    keys.GMAIL_MODE,
    keys.GMAIL_LABEL,
    keys.GEMINI_API_KEY,
    keys.DIGEST_ENABLED,
    keys.VERSION_CHECK_URL
  ];

  for (var i = 0; i < keyList.length; i++) {
    props.deleteProperty(keyList[i]);
  }
  return getUserConfig();
}
