// ========================================
// FUNCIONES AUXILIARES
// ========================================

/**
 * Muestra notificación toast en Google Sheets
 * @param {string} message - Mensaje a mostrar
 * @param {string} title - Título de la notificación
 * @param {number} [duration=5] - Duración en segundos
 * @param {string} [icon=''] - Emoji o icono (ℹ️, ✅, ❌, ⚠️)
 */
const showToast = (message, title, duration = 5, icon = "") => {
    const fullTitle = icon ? `${icon} ${title}` : title;
    SS.toast(message, fullTitle, duration);
};

/**
 * Genera clave única para identificar un préstamo
 * Formato: Título__Campus__Código_de_barras__Fecha_Vencimiento
 * 
 * @param {Array} row - Fila de datos
 * @returns {string} Clave única del préstamo
 */
const generateRecordKey = (row) => {
    return `${row[COLUMNS.TITLE]}__${row[COLUMNS.CAMPUS]}__${row[COLUMNS.BARCODE]}__${row[COLUMNS.DUE_DATE]}`;
};

/**
 * Valida que una hoja exista y esté accesible
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Hoja a validar
 * @param {string} sheetName - Nombre para mostrar en errores
 * @returns {boolean} true si la hoja es válida
 */
const validateSheet = (sheet, sheetName) => {
    if (!sheet) {
        showToast(
            `No se encontró la hoja: ${sheetName}`,
            "Error de configuración",
            5,
            "❌"
        );
        return false;
    }
    return true;
};

/**
 * Actualiza la bitácora de acciones de un registro
 * @param {number} rowNumber - Número de fila (1-indexed)
 * @param {string} action - Descripción de la acción
 * @param {string} [currentLog=''] - Bitácora existente
 * @returns {string} Bitácora actualizada
 */
const updateActionLog = (rowNumber, action, currentLog = "") => {
    const timestamp = new Date().toLocaleString("es-PE", {
        timeZone: "America/Lima",
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
    });

    const newEntry = `${timestamp}: ${action}`;
    const updatedLog = currentLog ? `${currentLog}\n${newEntry}` : newEntry;

    SHEETS.overdueItems.getRange(rowNumber, COLUMNS.LOG + 1).setValue(updatedLog);
    SHEETS.overdueItems.getRange(rowNumber, COLUMNS.ACTION + 1).clearContent();

    return updatedLog;
};

/**
 * Obtiene el contenido HTML de una plantilla
 * @param {string} templateName - Nombre del archivo de plantilla (sin extensión)
 * @returns {string} Contenido HTML de la plantilla
 */
const getEmailTemplate = (templateName) => {
    try {
        return HtmlService.createHtmlOutputFromFile(`templates/${templateName}`).getContent();
    } catch (error) {
        console.error(`Error cargando plantilla ${templateName}:`, error);
        return "";
    }
};

/**
 * Reemplaza variables en plantilla HTML
 * @param {string} template - Contenido HTML de la plantilla
 * @param {Object} variables - Objeto con las variables a reemplazar
 * @returns {string} HTML con variables reemplazadas
 */
const fillTemplate = (template, variables) => {
    let result = template;
    for (const [key, value] of Object.entries(variables)) {
        const regex = new RegExp(`{{${key}}}`, "g");
        result = result.replace(regex, value);
    }
    return result;
};
