# PLAN.md — Pendientes Uni (v1.0 Pública)

> Documento ejecutable y autocontenido. Todas las decisiones técnicas, de producto y de arquitectura están cerradas. No introducir opciones abiertas ni modificar el proyecto personal previo del usuario (`UAO_Pendientes_Nicolas` / `Sistema Autónomo Pendientes UAO`).

---

## 1. Contexto y arquitectura

### Objetivo
Construir **Pendientes Uni (v1.0)**, una plantilla pública y autocontenida en Google Sheets + Google Apps Script diseñada para que cualquier estudiante universitario sin conocimientos de programación, APIs ni hojas de cálculo pueda:
1. Instalar el sistema en su propia cuenta de Google en ~3 minutos mediante un enlace `/copy` y un asistente visual (Sidebar HTML), sin abrir nunca el editor de Apps Script.
2. Unificar entregas oficiales de su universidad (enlace `.ics` de Moodle/Canvas/Blackboard) con avisos informales (correos reenviados de profesores y avisos pegados de WhatsApp/Chat).
3. Consultar y completar sus tareas diariamente desde las apps nativas que ya tiene en su celular (**Google Calendar** y **Google Tasks**), sin necesidad de abrir Chrome ni Google Sheets tras la instalación.

### Stack tecnológico y restricciones
- **Runtime:** Google Apps Script (motor **V8**), sin servidor externo, sin base de datos central y sin `npm` nativo.
- **Almacenamiento por usuario:** Google Sheets activo (`SpreadsheetApp.getActiveSpreadsheet()`) + `PropertiesService.getScriptProperties()`.
- **Servicios Google:** `CalendarApp`, `GmailApp`, `SpreadsheetApp`, `UrlFetchApp`, `LockService`, `ScriptApp` y servicio avanzado **Google Tasks API v1** (`Tasks`).
- **Librería externa permitida (única):** `ical.min.js` de `kewisch/ical.js` (**MPL-2.0**), embebida intacta en `src/VendorIcal.gs` conservando su cabecera de licencia MPL-2.0.
- **IA (progresiva / BYOK):** API REST de Gemini (`generativelanguage.googleapis.com`) usando la API Key gratuita del propio estudiante. El conector `.ics` debe funcionar al 100% **sin API Key**; la API Key solo desbloquea el análisis de correos reenviados y avisos pegados.
- **Actualizaciones:** Archivo estático `version.json` alojado en GitHub (lectura semanal cacheada).
- **Costo operativo:** `$0/mes` para el creador y `$0/mes` para el estudiante.

### Decisiones técnicas cerradas (invariantes obligatorios)
1. **Una sola cuenta por defecto (`@gmail.com` personal):** Evita bloqueos de permisos de TI en cuentas universitarias y funciona con cualquier universidad (Google Workspace o Microsoft 365/Outlook). También puede correr en cuentas institucionales si la universidad no bloquea Apps Script.
2. **Ingestión en 2 capas:**
   - *Capa 1 (0 tokens, sin API Key):* Enlace `.ics` de Moodle/Canvas/Blackboard + autodetección de materias desde el campo `CATEGORIES` del `.ics`.
   - *Capa 2 (IA opcional BYOK):* Correos reenviados filtrados por etiqueta (por defecto `Universidad`, editable) o query académica directa + cuadro de texto en `Sidebar.html` para pegar avisos de WhatsApp/Chat.
3. **Salida nativa dual (Calendar + Tasks):**
   - **Google Calendar** (calendario dedicado `🎓 Pendientes Uni`): guarda la hora exacta de entrega y recordatorios nativos.
   - **Google Tasks** (lista dedicada `🎓 Pendientes Uni`): como la API de Google Tasks trunca `due` a medianoche UTC (sin hora exacta), el título **debe incluir la hora exacta** con el formato `[HH:mm] Actividad — Materia` (o `[Sin hora] Actividad — Materia` si no hay hora).
   - **Sincronización bidireccional de completado:** Marcar una tarea como completada en Google Tasks o en la columna `Completada` del Sheet propaga `[✅] ` al título del evento en Calendar, elimina recordatorios futuros y sincroniza ambos destinos.
4. **Invariantes heredados del motor SRE v5.1:**
   - **Zona horaria parametrizada:** Leer desde `CONFIG.TIMEZONE` (default `Session.getScriptTimeZone()`, editable en Sidebar); nunca hardcodear `America/Bogota` ni `SPREADSHEET_ID` ni materias fijas.
   - **Fechas en Sheets:** Columnas `Fecha_Limite`, `Fecha_Inicio_ISO` y `Fecha_ISO` siempre en formato texto (`@`). Al leer con `.getValues()`, soportar tanto `Date` como `string`.
   - **Gemini resiliente:** Descubrir modelo Flash **estable** vía `/models` (excluir `-preview` y `-exp` salvo último recurso) y cachear 24 h en `PropertiesService`. Enviar la API Key siempre en header `x-goog-api-key` (nunca en query string). Ante errores recuperables (`400` de modelo, `404`, `408`, `429`, `500`, `502`, `503`, `504` o excepción de red), rotar al siguiente modelo fallback respetando margen de 90–120 s antes del límite de 360 s de GAS. Errores de cuota/red nunca incrementan Dead-Letter.
   - **`CHRONO_GUARD` y cancelaciones:** Evaluar `accion` (`cancelar` / `aplazar_sin_fecha`) **antes** de filtrar por `es_pendiente`; nunca crear filas huérfanas para cancelaciones sin tarea existente. Si un mensaje es más antiguo que la última actualización de la tarea (`msgEsMasAntiguo === true`), jamás sobrescribir `Materia`, `Actividad`, `Origen` ni `ID_Referencia`; solo anexar URLs/notas nuevas deduplicadas.
   - **Integridad y concurrencia:** Usar `LockService.getScriptLock()`, escritura quirúrgica que nunca sobrescribe `Completada` ni `Notas_Personales`, poda automática >60 días en `IDs_Procesados`, y latido estructurado en nota `A1` de `Resumen_Tareas` (`estado=OK|DEGRADADO`, `ics_err`, `gmail_err`, `gemini_err`, `sync_err`, `dl_err`).
   - **Protección EdTech en Gmail:** Proteger dominios académicos/EdTech (`eltngl.com`, `ngl.cengage.com`, `moodle`, `canvas`, `blackboard`, `brightspace`) con precedencia booleana explícita frente a `-category:promotions`.
5. **Onboarding orientado a usuarios no técnicos:** El `Sidebar.html` y `GUIA_ESTUDIANTES.md` deben cubrir explícitamente cómo pasar la pantalla de *"Google no ha verificado esta app"* (*Opciones avanzadas → Ir a Pendientes Uni*), cómo copiar el enlace `.ics` y cómo crear la API Key gratuita en Google AI Studio.

### Decisiones descartadas (NO reintroducir)
- **NO** crear backend central, servidor Node/Python, Supabase, Firebase ni base de datos compartida.
- **NO** implementar sistema de cuentas, login, códigos de invitación ni panel de administración de usuarios.
- **NO** usar doble proyecto ni doble cuenta (`Vigilante Latido` / espejo `[UAO_MIRROR_ID:...]`).
- **NO** copiar código fuente de `derekantrican/GAS-ICS-Sync` (`Helpers.gs` / `Code.gs`) por incompatibilidad de licencia **GPL-3.0**.
- **NO** incluir integración obligatoria con Google Chat institucional ni hardcodear materias o IDs de la UAO.
- **NO** tocar ni modificar el sistema personal del usuario (`1jLjAD4g-gzVaPXZ7-7zK7dBb8Ky_m665aQF1CwaSbec` / `10X-_jlY5YGTH_VRxItPpDkTR7jw9Z3yUk-yHIzwbhqw76pDZl-YLT-DL`).

---

## 2. Mapa de archivos

Raíz del proyecto: `C:\dev\pendientes-uni`

### Archivos existentes a modificar
- Ninguno. Proyecto nuevo y aislado; no modificar archivos fuera de esta carpeta.

### Archivos nuevos a crear
| Ruta relativa | Responsabilidad |
| :--- | :--- |
| `appsscript.json` | Manifiesto V8, `oauthScopes` mínimos y servicio avanzado `tasks` (`v1`). |
| `version.json` | Metadatos públicos para GitHub (`latestVersion`, `releaseNotes`, `templateUrl`). |
| `GUIA_ESTUDIANTES.md` | Guía visual de 1 página para principiantes (pantalla "app no verificada", `.ics`, API Key y reenvío). |
| `src/Config.gs` | Constantes, lectura/escritura en `PropertiesService`, defaults y validación de configuración. |
| `src/SheetStore.gs` | Inicialización idempotente de pestañas (`Resumen_Tareas`, `Config_Materias`, `IDs_Procesados`), formato `@`, lectura segura `Date|string`, escritura quirúrgica, poda >60d y latido `A1`. |
| `src/VendorIcal.gs` | `ical.min.js` de `kewisch/ical.js` intacto con cabecera MPL-2.0. |
| `src/IcsConnector.gs` | Descarga y parseo de `.ics` con `ICAL.parse()`, normalización de fechas a `CONFIG.TIMEZONE`, ventana `-7` a `+120` días y autodetección de materias vía `CATEGORIES`. |
| `src/TaskEngine.gs` | Normalización de `Clave_Semantica`, fusión determinista `.ics` + IA, `CHRONO_GUARD` y evaluación previa de `cancelar` / `aplazar_sin_fecha`. |
| `src/CalendarTasksSync.gs` | Sincronización idempotente a calendario `🎓 Pendientes Uni` (hora exacta) y lista `🎓 Pendientes Uni` en Google Tasks (`[HH:mm] Actividad — Materia`), más sincronización bidireccional de `[✅]`. |
| `src/GmailConnector.gs` | Lectura de correos por etiqueta de reenvío o modo institucional, precedencia booleana EdTech y pre-filtrado determinista antes de llamar a IA. |
| `src/GeminiClient.gs` | Cliente Gemini Flash v5.1 (`/models` cacheado 24 h, `x-goog-api-key`, fallbacks, margen de tiempo 90–120 s, Dead-Letter) + función para analizar correos y avisos pegados de WhatsApp/Chat. |
| `src/DigestAndVersion.gs` | Resumen semanal por correo (`MailApp`/`GmailApp`) y chequeo semanal cacheado contra `version.json`. |
| `src/Main.gs` | `onOpen()`, menú `🎓 Pendientes Uni`, `ejecutarSincronizacion()`, gestión automática de triggers (`ScriptApp`), `LockService` y endpoints llamados por `Sidebar.html`. |
| `src/Sidebar.html` | Asistente visual cero-código: estado actual, configuración `.ics` + etiqueta Gmail + zona horaria, prueba de API Key Gemini, pegado rápido de avisos WhatsApp/Chat y activación 24/7. |
| `tests/smoke.mjs` | Suite ligera en Node.js (sin dependencias externas) para validar `TaskEngine`, `IcsConnector`, formato `[HH:mm]` de Tasks y `CHRONO_GUARD` localmente antes de subir a Apps Script. |

---

## 3. Fases y tareas atómicas

### Fase 1 — Base del proyecto, configuración y capa de Sheets
- [x] Crear `appsscript.json` con `runtimeVersion: "V8"`, servicio avanzado `tasks` (`v1`, `userSymbol: "Tasks"`), y `oauthScopes` mínimos (`spreadsheets.currentonly`, `calendar`, `tasks`, `gmail.readonly`, `gmail.send`, `script.external_request`, `script.scriptapp`, `script.container.ui`).
- [x] Crear `version.json` con estructura `{ "latestVersion": "1.0.0", "releasedAt": "YYYY-MM-DD", "highlights": [], "templateUrl": "" }`.
- [x] Implementar `src/Config.gs` con `getUserConfig()`, `saveUserConfig(patch)`, `resetConfig()`, soporte de `TIMEZONE` (default `Session.getScriptTimeZone()`), `ICS_URL`, `GMAIL_MODE` (`label` | `institutional` | `off`), `GMAIL_LABEL` (default `Universidad`), `GEMINI_API_KEY`, `DIGEST_ENABLED` y `VERSION_CHECK_URL`.
- [x] Implementar `src/SheetStore.gs` con inicialización idempotente de pestañas y encabezados exactos:
  - `Resumen_Tareas`: `Clave_Semantica`, `Materia`, `Actividad`, `Fecha_Limite`, `Fecha_Inicio_ISO`, `Estado`, `Completada`, `Origen`, `ID_Referencia`, `URLs`, `Ultima_Actualizacion_ISO`, `Calendar_Event_ID`, `Tasks_ID`, `Notas_Personales`.
  - `Config_Materias`: `Codigo_o_Alias`, `Nombre_Materia`, `Activa`, `Origen_Deteccion`.
  - `IDs_Procesados`: `ID_Referencia`, `Origen`, `Fecha_ISO`, `Resultado`.
- [x] Garantizar en `src/SheetStore.gs` formato texto (`@`) en columnas de fecha, helper `parseDateSafe(val, tz)` que soporte `Date` y `string`, escritura quirúrgica preservando `Completada` y `Notas_Personales`, poda >60 días en `IDs_Procesados` y registro de latido en nota `A1`.

#### Criterio de verificación — Fase 1
- Ejecutar `node -e "JSON.parse(require('fs').readFileSync('appsscript.json','utf8')); JSON.parse(require('fs').readFileSync('version.json','utf8')); console.log('JSON OK')"` y validar sintaxis JS de `src/Config.gs` y `src/SheetStore.gs` con `node --check`.
- Confirmar que no existen referencias hardcodeadas a `nicolas`, `uao.edu.co` ni IDs de hojas personales (`rg -i "nicolas|1jLjAD4g|10X-_jlY|UAO_MIRROR_ID" src/ appsscript.json` devuelve 0 coincidencias).

---

### Fase 2 — Conector `.ics` (modo 0 API Key) + Motor de tareas + Calendar & Google Tasks
- [x] Descargar e incorporar `ical.min.js` (release oficial de `kewisch/ical.js`, licencia MPL-2.0 intacta) en `src/VendorIcal.gs`.
- [x] Implementar `src/IcsConnector.gs` usando `ICAL.parse()` sobre `UrlFetchApp.fetch(icsUrl)` con ventana temporal `-7` a `+120` días, manejo de `TZID`/UTC hacia `CONFIG.TIMEZONE`, extracción de `UID`, `SUMMARY`, `DESCRIPTION`, `DTSTART`/`DTEND` y autodetección de materias desde `CATEGORIES` (insertando nuevas materias en `Config_Materias` sin duplicar).
- [x] Implementar `src/TaskEngine.gs` con generación determinista de `Clave_Semantica`, fusión de ítems de `.ics` e IA, evaluación de `accion` (`cancelar` / `aplazar_sin_fecha`) previa a `es_pendiente` sin crear filas huérfanas, y `CHRONO_GUARD` (`msgEsMasAntiguo === true` solo anexa URLs/notas deduplicadas).
- [x] Implementar `src/CalendarTasksSync.gs` con:
  - Creación/obtención idempotente del calendario `🎓 Pendientes Uni` y la lista de Google Tasks `🎓 Pendientes Uni`.
  - Formato obligatorio del título en Google Tasks: `[HH:mm] Actividad — Materia` (o `[Sin hora] Actividad — Materia`).
  - Sincronización bidireccional de completado: si la tarea está `completed` en Google Tasks o `Completada === true` en el Sheet, marcar ambos como completados, anteponer `[✅] ` al título del evento en Calendar y eliminar recordatorios futuros.
- [x] Crear `tests/smoke.mjs` con pruebas unitarias offline (usando fixtures `.ics` de Moodle y eventos simulados) que verifiquen: parseo `.ics`, autodetección de materias, `CHRONO_GUARD`, cancelación sin fila huérfana, formato `[HH:mm]` para Google Tasks y propagación `[✅]`.

#### Criterio de verificación — Fase 2
- Ejecutar `node tests/smoke.mjs` y confirmar que **100%** de los casos pasan (exit code `0`).
- Verificar sintaxis de todos los `.gs` creados con `node --check`.

---

### Fase 3 — Capa IA progresiva (Gmail + Gemini v5.1 + avisos WhatsApp/Chat)
- [ ] Implementar `src/GmailConnector.gs` con soporte para modo `label` (correos reenviados bajo `label:<GMAIL_LABEL>`) y modo `institutional` (query académica directa), precedencia booleana explícita para dominios EdTech frente a `-category:promotions` y filtrado previo de `IDs_Procesados`.
- [ ] Implementar `src/GeminiClient.gs` portando el cliente SRE v5.1:
  - Descubrimiento de modelos Flash estables vía `/models` cacheado 24 h en `PropertiesService` (excluyendo `-preview`/`-exp` salvo último recurso).
  - Envío obligatorio de API Key en header `x-goog-api-key`.
  - Rotación ante errores recuperables (`400` de modelo, `404`, `408`, `429`, `500`, `502`, `503`, `504` y excepciones de red) respetando margen de 90–120 s antes del límite de 360 s.
  - Dead-Letter estricto: errores de cuota/infraestructura nunca marcan un mensaje como error permanente.
  - Exponer `probarApiKeyGemini(apiKey)` para el Sidebar y `analizarTextoConGemini(texto, metadata)` reutilizable tanto para correos como para avisos pegados de WhatsApp/Chat.
- [ ] Ampliar `tests/smoke.mjs` con mocks de respuesta Gemini (creación, actualización más reciente, bloqueo `CHRONO_GUARD` por mensaje antiguo, cancelación existente vs. inexistente y aplazamiento sin fecha).

#### Criterio de verificación — Fase 3
- Ejecutar `node tests/smoke.mjs` y confirmar 100% de pruebas en verde (exit code `0`).
- Confirmar mediante búsqueda estática que la API Key nunca aparece en query strings (`rg "\?key=|&key=" src/` devuelve 0 coincidencias).

---

### Fase 4 — Orquestador, Asistente Visual (Sidebar), Resumen Semanal y Guía
- [ ] Implementar `src/DigestAndVersion.gs` con:
  - Resumen semanal por correo agrupado en 3 bloques (*Para esta semana*, *Próximas semanas*, *Sin fecha clara o atrasadas*) deduplicado por semana ISO.
  - Verificación semanal cacheada contra `version.json` en GitHub sin bloquear la ejecución si no hay red.
- [ ] Implementar `src/Main.gs` con:
  - `onOpen()` creando el menú `🎓 Pendientes Uni` (*Abrir asistente*, *Sincronizar ahora*, *Pegar aviso de WhatsApp/Chat*).
  - `ejecutarSincronizacion()` protegido con `LockService.getScriptLock()`, ejecutando en orden: lectura bidireccional de completados → `IcsConnector` (si hay URL) → `GmailConnector` + `GeminiClient` (solo si hay API Key y Gmail activo) → `TaskEngine` → `SheetStore` → `CalendarTasksSync` → poda >60d → latido `A1`.
  - Funciones puente para `Sidebar.html`: `apiGetState()`, `apiSaveSettings(data)`, `apiTestGemini(key)`, `apiProcessManualNotice(text)`, `apiRunSyncNow()`, `apiToggleAutoTrigger(enabled)` (crea/borra trigger único cada 2 horas con `ScriptApp`).
- [ ] Implementar `src/Sidebar.html` con interfaz limpia en español, pensada para usuarios sin experiencia técnica:
  - Sección de estado (último latido, trigger activo/inactivo, aviso si hay versión nueva).
  - Paso 1: Zona horaria + URL `.ics` del Campus Virtual + botón *"Guardar y sincronizar ahora"* (funciona sin API Key).
  - Paso 2 (IA opcional): Instrucciones breves para obtener API Key gratuita en Google AI Studio + campo protegido + botón *"Probar conexión"* + selector de etiqueta Gmail para reenvío.
  - Paso 3: Botón único *"Activar automático 24/7"* (toggle del trigger cada 2 h) y botón *"Pegar aviso de WhatsApp/Chat"* con previsualización del resultado.
- [ ] Redactar `GUIA_ESTUDIANTES.md` (máximo 1 página escaneable) cubriendo:
  1. Cómo hacer copia con el enlace `/copy`.
  2. Cómo pasar la pantalla *"Google no ha verificado esta app"* (*Opciones avanzadas → Ir a Pendientes Uni*).
  3. Cómo obtener el enlace `.ics` (Moodle / Canvas / Blackboard) y ver las tareas en Google Calendar + Google Tasks en el celular.
  4. Paso opcional: cómo crear la API Key gratuita de Gemini en 30 segundos y configurar filtro de reenvío desde el correo universitario.

#### Criterio de verificación — Fase 4
- Ejecutar `node --check` en todos los archivos `src/*.gs` y `node tests/smoke.mjs` (100% en verde).
- Desplegar el proyecto en una **hoja de cálculo de prueba limpia** en cuenta `@gmail.com` personal (nunca en la hoja personal de producción) y ejecutar las 4 pruebas en vivo:
  1. **Modo básico (sin API Key):** Configurar URL `.ics` desde el Sidebar → verificar creación de pestañas, fechas en formato `@`, materias autodetectadas en `Config_Materias`, eventos con hora exacta en Calendar y tareas `[HH:mm] ...` en Google Tasks.
  2. **Completado bidireccional:** Marcar 1 tarea como completada en Google Tasks → ejecutar *Sincronizar ahora* → confirmar `Completada = TRUE` en Sheet y `[✅] ` en Calendar sin duplicados.
  3. **Modo IA (con API Key):** Probar botón *"Probar conexión"*, pegar un aviso de prueba con cambio de fecha/cancelación en el Sidebar y verificar actualización correcta respetando `CHRONO_GUARD`.
  4. **Idempotencia:** Ejecutar una segunda sincronización inmediata y confirmar 0 duplicados en Sheet, Calendar y Google Tasks, y nota `A1` con `estado=OK`.
