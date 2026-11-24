import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.resolve(__dirname, '..');

// Mock mínimo de entorno Google Apps Script V8 para pruebas offline
const sandbox = {
  console,
  Date,
  Math,
  JSON,
  Object,
  Array,
  String,
  Number,
  Boolean,
  RegExp,
  Error,
  Intl,
  parseInt,
  parseFloat,
  isNaN,
  globalThis: null,
  Session: {
    getScriptTimeZone: () => 'America/Bogota'
  },
  PropertiesService: {
    _store: {},
    getScriptProperties() {
      const store = this._store;
      return {
        getProperties: () => ({ ...store }),
        setProperties: (obj) => Object.assign(store, obj),
        deleteProperty: (k) => { delete store[k]; }
      };
    }
  },
  Utilities: {
    formatDate(date, tz, pattern) {
      const parts = new Intl.DateTimeFormat('en-CA', {
        timeZone: tz || 'UTC',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
      }).formatToParts(date);
      const map = {};
      for (const p of parts) map[p.type] = p.value;
      const hh = map.hour === '24' ? '00' : map.hour;
      if (pattern === 'yyyy-MM-dd') return `${map.year}-${map.month}-${map.day}`;
      if (pattern === 'HH:mm') return `${hh}:${map.minute}`;
      if (pattern === "yyyy-MM-dd'T'HH:mm:ss") return `${map.year}-${map.month}-${map.day}T${hh}:${map.minute}:${map.second}`;
      // Offset aproximado para tests
      const offsetPart = new Intl.DateTimeFormat('en-US', {
        timeZone: tz || 'UTC',
        timeZoneName: 'shortOffset'
      }).formatToParts(date).find((p) => p.type === 'timeZoneName')?.value || 'GMT';
      let offset = '+00:00';
      const m = /^GMT([+-])(\d{1,2})(?::?(\d{2}))?$/.exec(offsetPart);
      if (m) {
        offset = `${m[1]}${m[2].padStart(2, '0')}:${(m[3] || '00').padStart(2, '0')}`;
      }
      return `${map.year}-${map.month}-${map.day}T${hh}:${map.minute}:${map.second}${offset}`;
    }
  }
};
sandbox.globalThis = sandbox;
vm.createContext(sandbox);

const gsFiles = [
  'src/VendorIcal.gs',
  'src/Config.gs',
  'src/SheetStore.gs',
  'src/IcsConnector.gs',
  'src/TaskEngine.gs',
  'src/CalendarTasksSync.gs'
];

for (const rel of gsFiles) {
  const code = fs.readFileSync(path.join(ROOT, rel), 'utf8');
  vm.runInContext(code, sandbox, { filename: rel });
}

// Helper mock de Spreadsheet para ensureSubjectsInConfig sin duplicados
function createMockSpreadsheet() {
  const sheets = {};
  function makeSheet(name) {
    const data = [];
    const formats = {};
    return {
      _data: data,
      _formats: formats,
      getLastRow: () => data.length,
      getLastColumn: () => (data[0] ? data[0].length : 0),
      getMaxRows: () => Math.max(data.length + 10, 20),
      setFrozenRows: () => {},
      getRange(row, col, numRows = 1, numCols = 1) {
        return {
          getValues() {
            const out = [];
            for (let r = 0; r < numRows; r++) {
              const rowArr = data[row - 1 + r] || [];
              const sliced = [];
              for (let c = 0; c < numCols; c++) {
                sliced.push(rowArr[col - 1 + c] !== undefined ? rowArr[col - 1 + c] : '');
              }
              out.push(sliced);
            }
            return out;
          },
          setValues(vals) {
            for (let r = 0; r < vals.length; r++) {
              const targetRow = row - 1 + r;
              if (!data[targetRow]) data[targetRow] = [];
              for (let c = 0; c < vals[r].length; c++) {
                data[targetRow][col - 1 + c] = vals[r][c];
              }
            }
          },
          setFontWeight() {},
          setNumberFormat(fmt) { formats[`${row}:${col}:${numRows}`] = fmt; },
          clearContent() { data.splice(row - 1, numRows); },
          setNote() {}
        };
      }
    };
  }
  return {
    getSheetByName: (name) => sheets[name] || null,
    insertSheet: (name) => {
      sheets[name] = makeSheet(name);
      return sheets[name];
    }
  };
}

let passed = 0;
function test(name, fn) {
  fn();
  passed++;
  console.log(`PASS: ${name}`);
}

// 1. Parseo .ics de Moodle + ventana [-7, +120] días + autodetección de materias
test('Parseo .ics Moodle, ventana temporal y autodetección de materias', () => {
  const now = new Date('2026-11-24T12:00:00Z');
  const icsFixture = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Moodle Pty Ltd//NONSGML Moodle Version 4.1//EN',
    'BEGIN:VEVENT',
    'UID:moodle-101@uni.edu',
    'SUMMARY:Taller 1 de Derivadas',
    'DESCRIPTION:Subir PDF en https://campus.example.edu/mod/assign/view.php?id=101',
    'CATEGORIES:CALC101 - Cálculo I',
    'DTSTART:20261129T045900Z',
    'DTEND:20261129T045900Z',
    'LAST-MODIFIED:20261124T100000Z',
    'END:VEVENT',
    'BEGIN:VEVENT',
    'UID:moodle-old@uni.edu',
    'SUMMARY:Quiz pasado hace mucho',
    'CATEGORIES:CALC101 - Cálculo I',
    'DTSTART:20261001T120000Z',
    'DTEND:20261001T120000Z',
    'END:VEVENT',
    'BEGIN:VEVENT',
    'UID:moodle-allday@uni.edu',
    'SUMMARY:Lectura Obligatoria',
    'CATEGORIES:ETICA200',
    'DTSTART;VALUE=DATE:20261201',
    'DTEND;VALUE=DATE:20261201',
    'END:VEVENT',
    'END:VCALENDAR'
  ].join('\r\n');

  const res = sandbox.parseIcsContent(icsFixture, { TIMEZONE: 'America/Bogota' }, now);
  assert.equal(res.items.length, 2, 'Debe filtrar el evento fuera de la ventana [-7, +120] días');
  assert.equal(res.detectedSubjects.length, 2, 'Debe autodetectar 2 materias únicas');

  const taller = res.items[0];
  assert.equal(taller.Materia, 'CALC101 - Cálculo I');
  assert.equal(taller.Actividad, 'Taller 1 de Derivadas');
  assert.equal(taller.hasTime, true);
  assert.ok(taller.Fecha_Limite.startsWith('2026-11-28T23:59:00'), 'Debe convertir 04:59Z a 23:59-05:00 en Bogota');
  assert.ok(taller.URLs.includes('https://campus.example.edu/mod/assign/view.php?id=101'));

  const lectura = res.items[1];
  assert.equal(lectura.hasTime, false);
  assert.equal(lectura.Fecha_Limite, '2026-12-01');

  // Inserción idempotente en Config_Materias sin duplicar
  const mockSs = createMockSpreadsheet();
  const inserted1 = sandbox.ensureSubjectsInConfig(res.detectedSubjects, mockSs);
  const inserted2 = sandbox.ensureSubjectsInConfig(res.detectedSubjects, mockSs);
  assert.equal(inserted1, 2);
  assert.equal(inserted2, 0, 'Segunda corrida no debe duplicar materias');
});

// 2. CHRONO_GUARD: mensaje antiguo no sobrescribe datos principales y solo anexa URL nueva
test('CHRONO_GUARD protege campos ante mensaje más antiguo y anexa URLs', () => {
  const existing = [{
    Clave_Semantica: 'calculo_i::taller_1',
    Materia: 'Cálculo I',
    Actividad: 'Taller 1',
    Fecha_Limite: '2026-11-30T23:59:00-05:00',
    Fecha_Inicio_ISO: '2026-11-30T23:59:00-05:00',
    Estado: 'Pendiente',
    Completada: false,
    Origen: 'ICS',
    ID_Referencia: 'uid-nuevo',
    URLs: 'https://campus.example.edu/1',
    Ultima_Actualizacion_ISO: '2026-11-24T15:00:00-05:00',
    Notas_Personales: 'Revisar ejercicio 4'
  }];

  const olderMsg = [{
    Clave_Semantica: 'calculo_i::taller_1',
    Materia: 'Materia Vieja',
    Actividad: 'Nombre Viejo',
    Fecha_Limite: '2026-11-25T10:00:00-05:00',
    Origen: 'GMAIL',
    ID_Referencia: 'msg-viejo',
    URLs: 'https://campus.example.edu/1 | https://drive.example.edu/guia.pdf',
    Ultima_Actualizacion_ISO: '2026-11-23T09:00:00-05:00',
    accion: 'crear_o_actualizar',
    es_pendiente: true
  }];

  const res = sandbox.mergeTasksWithEngine(existing, olderMsg, { TIMEZONE: 'America/Bogota' });
  assert.equal(res.stats.chronoBlocked, 1);
  assert.equal(res.patches.length, 1);
  const patch = res.patches[0];
  assert.equal(patch.Materia, 'Cálculo I');
  assert.equal(patch.Actividad, 'Taller 1');
  assert.equal(patch.Origen, 'ICS');
  assert.equal(patch.ID_Referencia, 'uid-nuevo');
  assert.equal(patch.Fecha_Limite, '2026-11-30T23:59:00-05:00');
  assert.equal(patch.URLs, 'https://campus.example.edu/1 | https://drive.example.edu/guia.pdf');
});

// 3. Cancelación sin fila huérfana + aplazar_sin_fecha evaluados antes de es_pendiente
test('Cancelación sin fila huérfana y aplazamiento sin fecha', () => {
  const existing = [{
    Clave_Semantica: 'fisica_i::laboratorio_2',
    Materia: 'Física I',
    Actividad: 'Laboratorio 2',
    Fecha_Limite: '2026-12-02T18:00:00-05:00',
    Fecha_Inicio_ISO: '2026-12-02T18:00:00-05:00',
    Estado: 'Pendiente',
    Completada: false,
    Origen: 'ICS',
    ID_Referencia: 'lab-2',
    URLs: '',
    Ultima_Actualizacion_ISO: '2026-11-24T10:00:00-05:00'
  }];

  const incoming = [
    {
      Materia: 'Física I',
      Actividad: 'Examen Inexistente',
      accion: 'cancelar',
      es_pendiente: false,
      Ultima_Actualizacion_ISO: '2026-11-24T12:00:00-05:00'
    },
    {
      Materia: 'Física I',
      Actividad: 'Laboratorio 2',
      accion: 'cancelar',
      es_pendiente: false,
      Ultima_Actualizacion_ISO: '2026-11-24T12:00:00-05:00'
    }
  ];

  const res = sandbox.mergeTasksWithEngine(existing, incoming, { TIMEZONE: 'America/Bogota' });
  assert.equal(res.stats.orphanCancelledSkipped, 1, 'No debe crear fila huérfana');
  assert.equal(res.stats.cancelled, 1);
  assert.equal(res.patches.length, 1);
  assert.equal(res.patches[0].Estado, 'Cancelada');

  const resPostpone = sandbox.mergeTasksWithEngine(existing, [{
    Materia: 'Física I',
    Actividad: 'Laboratorio 2',
    accion: 'aplazar_sin_fecha',
    es_pendiente: false,
    Ultima_Actualizacion_ISO: '2026-11-24T13:00:00-05:00'
  }], { TIMEZONE: 'America/Bogota' });
  assert.equal(resPostpone.stats.postponed, 1);
  assert.equal(resPostpone.patches[0].Estado, 'Aplazada / Sin fecha');
  assert.equal(resPostpone.patches[0].Fecha_Limite, '');
});

// 4. Formato [HH:mm] y [Sin hora] para Google Tasks
test('Formato obligatorio [HH:mm] y [Sin hora] en Google Tasks', () => {
  const withTime = sandbox.formatTaskTitle(
    'Taller 1 de Derivadas',
    'Cálculo I',
    '2026-11-28T23:59:00-05:00',
    'America/Bogota'
  );
  assert.equal(withTime, '[23:59] Taller 1 de Derivadas — Cálculo I');

  const withoutTime = sandbox.formatTaskTitle(
    'Lectura Obligatoria',
    'Ética',
    '2026-12-01',
    'America/Bogota'
  );
  assert.equal(withoutTime, '[Sin hora] Lectura Obligatoria — Ética');
});

// 5. Propagación bidireccional [✅]
test('Propagación bidireccional [✅] entre Tasks, Sheet y Calendar', () => {
  const state = sandbox.reconcileCompletionState(
    { Completada: false },
    { status: 'completed' }
  );
  assert.equal(state.completed, true);
  assert.equal(state.sheetNeedsUpdate, true);
  assert.equal(state.calendarNeedsPrefix, true);

  const prefixedOnce = sandbox.applyCompletionPrefixToTitle('Taller 1 — Cálculo I', true);
  const prefixedTwice = sandbox.applyCompletionPrefixToTitle(prefixedOnce, true);
  assert.equal(prefixedOnce, '[✅] Taller 1 — Cálculo I');
  assert.equal(prefixedTwice, '[✅] Taller 1 — Cálculo I', 'No debe duplicar [✅] ');
});

console.log(`\nResumen smoke tests: ${passed}/5 pasaron.`);
