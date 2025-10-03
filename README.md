# Documentación del Sistema de Gestión de Deudores

## Información General

- **Nombre del Proyecto**: Sistema de Gestión de Deudores
- **Autor**: Fredy Romero <romeroespinoza.fp@gmail.com>
- **Versión**: 1.0.0
- **Licencia**: MIT
- **ID del Script**: 16b7rtb2oDqkQ2TOZ8TqOwNasxu9nkJMFpvAe4aodz17Uq3Qzs7FN4iae

## Descripción

Este sistema automatiza el proceso de gestión de deudores utilizando Google Apps Script, permitiendo una integración perfecta con las herramientas de Google Workspace. Está diseñado para gestionar registros de deudores, procesar préstamos vencidos y enviar notificaciones por correo electrónico.

## Estructura del Proyecto

```
/sp_deudores/
├── .clasp.json         # Configuración para Google Apps Script
├── .gitignore          # Archivos ignorados por Git
├── README.md           # Información básica del proyecto
├── appsscript.json     # Configuración del script para Google Apps Script
├── emailTemplate.html  # Plantilla HTML para correos electrónicos
├── main.js             # Código principal del sistema
└── package.json        # Dependencias y configuración del proyecto
```

## Hojas de Cálculo

El sistema utiliza las siguientes hojas de cálculo:

| Nombre de la Hoja | Variable en Código | Descripción |
|-------------------|-------------------|-------------|
| Reporte de deudores - Widget | `SHEETS.alma` | Contiene los datos importados del sistema Alma |
| Préstamos vencidos / Deudores | `SHEETS.prestamosVencidos` | Almacena información de préstamos vencidos y deudores activos |
| Seguimiento de préstamos | `SHEETS.seguimientoPrestamos` | Registra el seguimiento de préstamos en proceso |
| Recursos devueltos / Histórico | `SHEETS.recursosDevueltos` | Histórico de recursos que han sido devueltos |

## Estructura de Columnas

### Hoja "Reporte de deudores - Widget"
- Columnas A-L: Datos importados del sistema Alma
- Columna M (13): Estado del registro ("YA REGISTRADO" o "NUEVO DEUDOR")

### Hoja "Préstamos vencidos / Deudores"
- Columnas A-L (1-12): Datos del préstamo y deudor
- Columna N (14): Acciones a realizar

### Hoja "Recursos devueltos / Histórico"
- Columnas A-L (1-12): Datos del préstamo y deudor
- Columna M (13): Indicador de devolución ("Sí")
- Columna N (14): Fecha de devolución
- Columna O (15): Comentario sobre la devolución
- Columna P (16): Mes en texto (ej. "Abril")
- Columna Q (17): Mes en número (ej. "04")

## Funciones Principales

### Funciones de Procesamiento

| Función | Descripción |
|---------|-------------|
| `resetSheetForNewData()` | Limpia la hoja "Reporte de deudores - Widget" para nuevos datos |
| `startProcess()` | Procesa los datos de Alma, identifica nuevos deudores y recursos devueltos |

### Funciones de Acciones

| Función | Descripción |
|---------|-------------|
| `moverARecursosDevueltos()` | Mueve registros a la hoja "Recursos devueltos / Histórico" |
| `moverASeguimientoPrestamos()` | Mueve registros a la hoja "Seguimiento de préstamos" |
| `enviarPrimerRecordatorio()` | Envía correo de primer recordatorio (pendiente de implementación) |
| `enviarSegundoRecordatorio()` | Envía correo de segundo recordatorio (pendiente de implementación) |
| `enviarAvisoRecarga()` | Envía correo de aviso de recarga (pendiente de implementación) |
| `enviarConfirmacionRecarga()` | Envía correo de confirmación de recarga (pendiente de implementación) |
| `executeActions()` | Ejecuta acciones basadas en los valores de la columna N (14) |

### Funciones de Interfaz

| Función | Descripción |
|---------|-------------|
| `onOpen()` | Crea el menú personalizado en la interfaz de Google Sheets |

## Flujo de Trabajo

1. **Importación de Datos**:
   - Los datos se importan a la hoja "Reporte de deudores - Widget"
   - Se ejecuta la función `startProcess()`

2. **Procesamiento**:
   - Se identifican nuevos deudores y se agregan a "Préstamos vencidos / Deudores"
   - Se identifican recursos devueltos y se mueven a "Recursos devueltos / Histórico"
   - Se actualiza el estado en la hoja "Reporte de deudores - Widget"

3. **Acciones**:
   - Se ejecuta la función `executeActions()` para procesar las acciones definidas
   - Las acciones pueden incluir envío de correos o movimiento de registros

## Relaciones entre Hojas

```
Reporte de deudores - Widget
         ↓
         ↓ (Nuevos deudores)
         ↓
Préstamos vencidos / Deudores
         ↓
         ↓ (Según acción)
         ↓
  ┌──────┴──────┐
  ↓             ↓
Seguimiento    Recursos
de préstamos   devueltos / Histórico
```

## Identificación de Registros

Los registros se identifican mediante una clave compuesta:
```javascript
const recordKey = `${row[2]}__${row[8]}__${row[9]}__${row[10]}`;
```

Esta clave permite identificar de manera única cada registro a través de las diferentes hojas.

## Acciones Disponibles

Las acciones disponibles en la columna N (14) de "Préstamos vencidos / Deudores" son:

- "Enviar correo: Primer recordatorio"
- "Enviar correo: Segundo recordatorio"
- "Enviar correo: Aviso de recarga"
- "Enviar correo: Confirmación de la recarga"
- "Mover a: Recursos devueltos / Histórico"
- "Mover a: Seguimiento de préstamos"

## Menú Personalizado

El sistema crea un menú personalizado en Google Sheets con las siguientes opciones:

- 🔄 Procesar reporte de Alma
- ⚡ Ejecutar acciones por ítem
- 🗑️ Limpiar información
- ⚙️ Avanzado (solo para usuario autorizado)
  - Mover a: Seguimiento de préstamos
  - Mover a: Recursos devueltos

## Notas Adicionales

- El sistema está configurado para funcionar en la zona horaria "America/Lima"
- Existe un usuario autorizado (`AUTHORIZED_USER`) con acceso a funciones avanzadas
- Las plantillas de correo electrónico se encuentran en el archivo `emailTemplate.html`
- Algunas funciones de envío de correo están pendientes de implementación