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
 * Índices de columnas para facilitar mantenimiento
 */
const COLUMNS = {
    DATE: 0,
    TIME: 1,
    NAME: 2,
    LASTNAME: 3,
    USER_ID: 4,
    EMAIL: 5,
    TITLE: 6,
    BARCODE: 7,
    LIBRARY: 8,
    LOCATION: 9,
    DUE_DATE: 10,
    ACTION: 11,
    LOG: 12,
    STATUS: 11,
    RETURN_DATE: 11,
    RETURN_COMMENT: 12,
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
