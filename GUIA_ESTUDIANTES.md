# 🎓 Pendientes Uni — Guía rápida para estudiantes (3 minutos)

Organiza automáticamente las entregas de tu universidad en **Google Calendar** y **Google Tasks** en tu celular, sin programar ni tocar código.

---

## 1. Crea tu propia copia privada
1. Abre el enlace `/copy` que recibiste y haz clic en **Hacer una copia**.
2. Inicia sesión con tu **cuenta personal `@gmail.com`** (recomendado para evitar bloqueos de TI de la universidad).
3. Espera unos 5 segundos a que aparezca el menú **🎓 Pendientes Uni** en la barra superior de Google Sheets.

## 2. Concede permisos la primera vez ("App no verificada")
Al abrir el asistente por primera vez, Google mostrará un aviso de seguridad porque tu copia corre bajo tu propia cuenta:
1. Haz clic en **Continuar** y elige tu cuenta `@gmail.com`.
2. Si aparece *"Google no ha verificado esta app"*, haz clic en **Opciones avanzadas** (abajo a la izquierda).
3. Haz clic en **Ir a Pendientes Uni (no seguro)** y luego en **Permitir**.

## 3. Conecta tu Campus Virtual (sin API Key)
1. En el menú superior ve a **🎓 Pendientes Uni → Abrir asistente**.
2. En el **Paso 1**, elige tu zona horaria y pega tu enlace `.ics` del Campus Virtual:
   - **Moodle:** *Calendario → Exportar calendario → Todos los eventos → Obtener URL del calendario*.
   - **Canvas:** *Calendario → Fuente del calendario (abajo a la derecha) → Copiar enlace*.
   - **Blackboard:** *Calendario → Configuración → Obtener enlace externo*.
3. Haz clic en **Guardar y sincronizar ahora**.
4. Listo: verás tus entregas con hora exacta en **Google Calendar** (calendario `🎓 Pendientes Uni`) y como lista de pendientes en **Google Tasks** (lista `🎓 Pendientes Uni`). Cuando marques una tarea como completada en Google Tasks, también se marcará con `✅` en tu calendario.

## 4. (Opcional) Activa la IA para correos y WhatsApp
1. Entra a [Google AI Studio](https://aistudio.google.com/apikey) y haz clic en **Create API key** (es gratuita).
2. Pégala en el **Paso 2** del asistente y haz clic en **Probar conexión**.
3. Si tu universidad usa otro correo (como Outlook), crea un reenvío automático de correos de profesores hacia tu `@gmail.com` con un filtro que les ponga la etiqueta **`Universidad`**.
4. En el **Paso 3**, haz clic en **Activar automático 24/7** para que todo se actualice solo cada 2 horas, y usa la caja de texto cuando quieras pegar un aviso rápido de WhatsApp.
