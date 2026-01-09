# Guía de Migración a v2.0.0

## Cambios Principales

### 1. Estructura Modular

El código se ha dividido en 6 archivos separados:

- **Config.js** (73 líneas): Constantes SHEETS, COLUMNS, ACTIONS, STATUS
- **Utils.js** (104 líneas): Funciones auxiliares (showToast, validateSheet, etc.)
- **DataProcessor.js** (185 líneas): Procesamiento de datos (startProcess, deleteData)
- **Emails.js** (209 líneas): **NUEVO** - Funciones de email ahora implementadas
- **Actions.js** (231 líneas): Ejecución de acciones (executeActions, moveToXXX)
- **UI.js** (45 líneas): Interfaz de usuario (onOpen, hasScript)
- **Main.js** (51 líneas): Documentación y punto de entrada

**Total**: ~898 líneas (vs 957 líneas anteriores)

### 2. Funciones de Email Implementadas ✅

Las funciones de email ahora están **funcionalmente completas**:

```javascript
// ANTES: Función vacía
const sendFirstReminder = (data, rowNumber) => {
  // TODO: Implementar
};

// AHORA: Función completa
const sendFirstReminder = (data, rowNumber) => {
  const email = data[COLUMNS.EMAIL];
  const subject = "📚 Recordatorio: Devolución de recurso pendiente";
  const body = createReminderEmailBody(data);
  
  if (sendEmail(email, subject, body)) {
    updateActionLog(rowNumber, "✉️ Primer recordatorio enviado", currentLog);
  }
};
```

**Implementado:**
- ✅ `sendFirstReminder()` - Email con recordatorio básico
- ✅ `sendSecondReminder()` - Email con tono más urgente
- ✅ `sendRechargeNotice()` - Aviso de recarga (pendiente personalización)
- ✅ `sendRechargeConfirmation()` - Confirmación de pago (pendiente personalización)

### 3. Comentarios Reducidos

Se eliminaron comentarios redundantes manteniendo:
- JSDoc con tipos para ayuda del IDE
- Explicaciones de POR QUÉ (no QUÉ hace el código)
- Diagramas de flujo principales

**Reducción**: ~400 líneas de comentarios → ~150 líneas

### 4. Configuración Actualizada

#### appsscript.json

Se agregaron permisos OAuth para envío de emails:

```json
{
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.send"
  ]
}
```

## Instalación en Google Apps Script

### Opción 1: Subir manualmente (Recomendado para primera vez)

1. Abre tu proyecto en [script.google.com](https://script.google.com)
2. **Borra el archivo Main.gs existente**
3. Crea los siguientes archivos (Archivo → Nuevo → Archivo de comandos):
   - `Config.js`
   - `Utils.js`
   - `DataProcessor.js`
   - `Emails.js`
   - `Actions.js`
   - `UI.js`
   - `Main.js`
4. Copia el contenido de cada archivo local a su correspondiente en el editor
5. Guarda el proyecto (Ctrl+S)

### Opción 2: Usar clasp (Para desarrolladores)

```bash
# Asegúrate de tener clasp instalado
npm install -g @google/clasp

# Push todos los archivos al proyecto
clasp push
```

## Verificación Post-Migración

### 1. Verificar que no hay errores de sintaxis

En el editor de Apps Script:
- Revisa que no aparezcan subrayados rojos
- Ejecuta "Ver" → "Registros" para ver si hay errores

### 2. Probar el menú

1. Abre el Google Sheet vinculado
2. Recarga la página (F5)
3. Verifica que aparezca el menú "Scripts 🟢"
4. Prueba "⚠️ Información del script"

### 3. Probar funcionalidad básica

**TEST 1: Limpiar datos**
- Scripts 🟢 → 🗑️ Borrar datos de: Reporte de deudores - Widget
- Debe mostrar toast con cantidad de filas eliminadas

**TEST 2: Procesar datos** (requiere datos de prueba)
- Agregar filas de prueba en la hoja Alma
- Scripts 🟢 → ➡️ Procesar datos de: Reporte de deudores - Widget
- Verificar toast con resumen

**TEST 3: Enviar email de prueba**
1. En "Préstamos vencidos / Deudores", selecciona una fila
2. En columna L, selecciona "✉️ Primer recordatorio"
3. Scripts 🟢 → 🧪 Ejecutar acciones (L) de: Préstamos vencidos
4. Verifica que:
   - Se envíe el correo
   - La bitácora (columna M) se actualice
   - La acción (columna L) se limpie

## Resolución de Problemas

### Error: "Cannot find name 'SHEETS'"

**Causa**: Los archivos no se cargaron en el orden correcto

**Solución**: Asegúrate de que `Config.js` esté cargado primero. En Apps Script, todos los archivos comparten el mismo namespace, así que el orden no debería importar, pero puedes intentar:
1. Cerrar y reabrir el editor
2. Recargar el Sheet

### Error: "Exception: Service invoked too many times for one day: email"

**Causa**: Límite de cuota de Gmail alcanzado (500 emails/día para cuentas gratuitas)

**Solución**: Espera 24 horas o usa una cuenta de Google Workspace

### Error: "Authorization required"

**Causa**: El script necesita permisos para enviar emails

**Solución**:
1. Ejecuta cualquier función manualmente desde el editor
2. Acepta los permisos solicitados
3. Vuelve a intentar desde el menú del Sheet

## Diferencias de Comportamiento

### ⚠️ IMPORTANTE: No hay cambios en la funcionalidad

La refactorización **NO cambia el comportamiento** del sistema:
- ✅ Mismas hojas de Google Sheets
- ✅ Misma estructura de datos
- ✅ Mismo flujo de trabajo
- ✅ Mismo menú de usuario
- ➕ **NUEVO**: Emails ahora se envían realmente (antes solo actualizaban log)

## Próximos Pasos Recomendados

1. **Personalizar plantillas de email**:
   - Editar `Emails.js` líneas 40-85 (función `createReminderEmailBody`)
   - Ajustar colores, logos, textos según necesidades

2. **Integrar plantillas HTML existentes** (templates/*.html):
   - Usar función `getEmailTemplate()` en `Utils.js`
   - Reemplazar variables `{{NOMBRE}}`, `{{TITULO}}`, etc.

3. **Agregar validaciones adicionales**:
   - Verificar formato de email antes de enviar
   - Confirmar con el usuario antes de enviar emails masivos

4. **Monitorear cuotas de Gmail**:
   - Implementar contador de emails enviados
   - Pausar envíos si se acerca al límite

## Contacto

Si encuentras problemas durante la migración:
- Revisa los logs del Apps Script (Ver → Registros)
- Compara con el código original en Main.js (backup recomendado)
- Consulta la documentación de Google Apps Script

---

**Versión**: 2.0.0  
**Fecha**: 2026-01-09  
**Autor**: Refactorización por Antigravity AI
