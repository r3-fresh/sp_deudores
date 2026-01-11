// ========================================
// CONFIGURACIÓN Y CONSTANTES
// ========================================

/**
 * Interfaz de usuario de Google Sheets
 * @type {GoogleAppsScript.Spreadsheet.Ui}
 */
const UI = SpreadsheetApp.getUi();

/**
 * Referencia al archivo de Google Sheets activo
 * @type {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
const SS = SpreadsheetApp.getActiveSpreadsheet();

/**
 * Mapeo de hojas del documento por ID
 * @typedef {Object} SheetRefs
 * @property {GoogleAppsScript.Spreadsheet.Sheet} alma - Datos importados desde Alma
 * @property {GoogleAppsScript.Spreadsheet.Sheet} overdueItems - Préstamos vencidos activos
 * @property {GoogleAppsScript.Spreadsheet.Sheet} trackingItems - Préstamos en seguimiento
 * @property {GoogleAppsScript.Spreadsheet.Sheet} returnedItems - Histórico de devoluciones
 */
const SHEETS = {
    alma: SS.getSheetById(563966915),
    overdueItems: SS.getSheetById(491373272),
    trackingItems: SS.getSheetById(687630222),
    returnedItems: SS.getSheetById(1634827826),
};

/**
 * Índices de columnas (0-indexed)
 * 
 * Estructura de overdueItems (17 columnas):
 * 0-Campus, 1-Tipo usuario, 2-ID Usuario, 3-Apellidos y Nombres, 4-Celular,
 * 5-Correo, 6-Título, 7-Código clasificación, 8-Código barras,
 * 9-Fecha Préstamo, 10-Fecha Vencimiento, 11-Acciones, 12-Bitácora,
 * 13-Fecha recargo boleta, 14-Fecha retiro boleta, 15-Costo, 16-Observaciones
 * 
 * Estructura trackingItems/returnedItems (20 columnas):
 * Las primeras 11 igual a overdueItems, luego:
 * 11-Fecha seguimiento/devolución, 12-Bitácora, 13-Fecha recargo,
 * 14-Fecha retiro, 15-Costo, 16-Observaciones, 17-Estado,
 * 18-Consulta pago caja, 19-¿Realizó pago?
 */
const COLUMNS = {
    // Columnas principales (comunes en todas las hojas)
    CAMPUS: 0,
    USER_TYPE: 1,
    USER_ID: 2,
    FULL_NAME: 3,
    PHONE: 4,
    EMAIL: 5,
    TITLE: 6,
    CLASSIFICATION: 7,
    BARCODE: 8,
    LOAN_DATE: 9,
    DUE_DATE: 10,

    // Columnas específicas de overdueItems
    ACTION: 11,           // Acciones (overdueItems)
    LOG: 12,              // Bitácora de acciones
    RECHARGE_DATE: 13,    // Fecha de recargo a la boleta
    WITHDRAWAL_DATE: 14,  // Fecha de retiro en la boleta
    COST: 15,             // Costo
    OBSERVATIONS: 16,     // Observaciones

    // Columnas adicionales en trackingItems/returnedItems
    TRACKING_DATE: 11,    // Fecha de seguimiento (trackingItems)
    RETURN_DATE: 11,      // Fecha de devolución (returnedItems)
    STATUS: 17,           // Estado
    PAYMENT_QUERY: 18,    // Consulta de pago a caja
    PAYMENT_DONE: 19,     // ¿Realizó el pago?

    // Columna específica de Alma (12 columnas total: 0-11)
    ALMA_STATUS: 11,      // Estado (NUEVO DEUDOR / YA REGISTRADO)

    // Contadores de columnas por hoja
    ALMA_TOTAL: 12,       // Total de columnas en Alma
    OVERDUE_TOTAL: 17,    // Total de columnas en overdueItems
    TRACKING_TOTAL: 20,   // Total de columnas en trackingItems
    RETURNED_TOTAL: 20,   // Total de columnas en returnedItems
};

/**
 * Acciones disponibles en el sistema
 */
const ACTIONS = {
    FIRST_REMINDER: "✉️ Primer recordatorio",
    SECOND_REMINDER: "✉️ Segundo recordatorio",
    RECHARGE_NOTICE: "✉️ Aviso de recarga",
    RECHARGE_CONFIRMATION: "✉️ Confirmación de la recarga",
    MOVE_TO_RETURNED: "Ítem devuelto/encontrado",
    MOVE_TO_TRACKING: "Dar seguimiento al ítem",
};

/**
 * Estados posibles en la columna de estado de Alma
 */
const STATUS = {
    REGISTERED: "YA REGISTRADO",
    NEW: "NUEVO DEUDOR",
};

/**
 * Configuración de emails
 */
const EMAIL_CONFIG = {
    CAMPUS_NAME: "Huancayo",  // Nombre del campus para asuntos de correo
    MAILBOX_IMAGE_URL: "https://hubinformacion.continental.edu.pe/web/wp-content/uploads/2026/01/buzones-hyo.png",     // URL de la imagen del buzón (agregar por el usuario)
};
