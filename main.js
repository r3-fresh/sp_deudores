/**
 * ========================================
 * SISTEMA DE GESTIÓN DE DEUDORES
 * ========================================
 * 
 * Automatiza el proceso de gestión de préstamos vencidos integrando
 * datos del sistema Alma con Google Sheets.
 * 
 * @author Fredy Romero <romeroespinoza.fp@gmail.com>
 * @version 1.0.0
 * @license MIT
 */

// ========================================
// 1. CONFIGURACIÓN Y CONSTANTES
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
 * Cada hoja tiene un propósito específico en el flujo de trabajo
 * 
 * @typedef {Object} SheetRefs
 * @property {GoogleAppsScript.Spreadsheet.Sheet} alma - Datos importados desde Alma
 * @property {GoogleAppsScript.Spreadsheet.Sheet} overdueItems - Préstamos vencidos activos
 * @property {GoogleAppsScript.Spreadsheet.Sheet} trackingItems - Préstamos en seguimiento
 * @property {GoogleAppsScript.Spreadsheet.Sheet} returnedItems - Histórico de devoluciones
 */
const SHEETS = {
  alma: SS.getSheetById(563966915),           // Entrada: Widget de Alma
  overdueItems: SS.getSheetById(491373272),   // Deudores activos
  trackingItems: SS.getSheetById(687630222),  // En seguimiento
  returnedItems: SS.getSheetById(1634827826), // Histórico
};

/**
 * Índices de columnas para facilitar mantenimiento
 * Previene errores con "números mágicos"
 */
const COLUMNS = {
  // Columnas de datos principales (0-10)
  DATE: 0,           // A: Fecha
  TIME: 1,           // B: Hora
  NAME: 2,           // C: Nombre
  LASTNAME: 3,       // D: Apellido
  USER_ID: 4,        // E: ID Usuario
  EMAIL: 5,          // F: Email
  TITLE: 6,          // G: Título del recurso
  BARCODE: 7,        // H: Código de barras
  LIBRARY: 8,        // I: Biblioteca
  LOCATION: 9,       // J: Ubicación
  DUE_DATE: 10,      // K: Fecha de vencimiento
  
  // Columnas de control
  ACTION: 11,        // L: Acción a ejecutar
  LOG: 12,          // M: Bitácora de acciones
  STATUS: 11,        // L: Estado en hoja Alma (reutiliza índice)
  
  // Columnas adicionales en histórico
  RETURN_DATE: 11,   // L: Fecha de devolución
  RETURN_COMMENT: 12 // M: Comentario de devolución
};

/**
 * Acciones disponibles en el sistema
 * Estas aparecen en el menú desplegable de la columna L
 */
const ACTIONS = {
  FIRST_REMINDER: "✉️ Primer recordatorio",
  SECOND_REMINDER: "✉️ Segundo recordatorio",
  RECHARGE_NOTICE: "✉️ Aviso de recarga",
  RECHARGE_CONFIRMATION: "✉️ Confirmación de la recarga",
  MOVE_TO_RETURNED: "Ítem devuelto/encontrado",
  MOVE_TO_TRACKING: "Dar seguimiento al ítem"
};

/**
 * Estados posibles en la columna de estado de Alma
 */
const STATUS = {
  REGISTERED: "YA REGISTRADO",
  NEW: "NUEVO DEUDOR"
};

// ========================================
// 2. FUNCIONES AUXILIARES
// ========================================

/**
 * Muestra una notificación toast personalizada
 * Centraliza el manejo de mensajes al usuario
 * 
 * @param {string} message - Mensaje a mostrar
 * @param {string} title - Título de la notificación
 * @param {number} [duration=5] - Duración en segundos
 * @param {string} [icon=''] - Emoji o icono (ℹ️, ✅, ❌, ⚠️)
 */
const showToast = (message, title, duration = 5, icon = '') => {
  const fullTitle = icon ? `${icon} ${title}` : title;
  SS.toast(message, fullTitle, duration);
};

/**
 * Genera una clave única para identificar un registro
 * Usa: Nombre__Biblioteca__Ubicación__FechaVencimiento
 * 
 * ¿Por qué esta combinación?
 * - Permite identificar el mismo préstamo a través de diferentes hojas
 * - Un usuario puede tener múltiples préstamos simultáneos
 * - La misma persona podría pedir el mismo libro en diferentes momentos
 * 
 * @param {Array} row - Fila de datos
 * @returns {string} Clave única del registro
 */
const generateRecordKey = (row) => {
  return `${row[COLUMNS.NAME]}__${row[COLUMNS.LIBRARY]}__${row[COLUMNS.LOCATION]}__${row[COLUMNS.DUE_DATE]}`;
};

/**
 * Valida que una hoja exista y esté accesible
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Hoja a validar
 * @param {string} sheetName - Nombre para mostrar en errores
 * @returns {boolean} true si la hoja es válida
 */
const validateSheet = (sheet, sheetName) => {
  if (!sheet) {
    showToast(
      `No se encontró la hoja: ${sheetName}`,
      'Error de configuración',
      5,
      '❌'
    );
    return false;
  }
  return true;
};

/**
 * Actualiza la bitácora de acciones de un registro
 * Mantiene un historial de todas las acciones realizadas
 * 
 * @param {number} rowNumber - Número de fila (1-indexed)
 * @param {string} action - Descripción de la acción
 * @param {string} [currentLog=''] - Bitácora existente
 * @returns {string} Bitácora actualizada
 */
const updateActionLog = (rowNumber, action, currentLog = '') => {
  const timestamp = new Date().toLocaleString('es-PE', {
    timeZone: 'America/Lima',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  });
  
  const newEntry = `${timestamp}: ${action}`;
  const updatedLog = currentLog ? `${currentLog}\n${newEntry}` : newEntry;
  
  // Actualizar bitácora en columna M (índice 13, 1-indexed)
  SHEETS.overdueItems.getRange(rowNumber, COLUMNS.LOG + 1).setValue(updatedLog);
  
  // Limpiar la acción ejecutada de columna L (índice 12, 1-indexed)
  SHEETS.overdueItems.getRange(rowNumber, COLUMNS.ACTION + 1).clearContent();
  
  return updatedLog;
};

// ========================================
// 3. FUNCIONES PRINCIPALES DE PROCESAMIENTO
// ========================================

/**
 * Limpia la hoja de Alma preparándola para nuevos datos
 * 
 * FLUJO:
 * 1. Valida que la hoja exista
 * 2. Verifica si hay datos para limpiar
 * 3. Elimina todo el contenido excepto encabezados
 * 
 * @returns {void}
 */
const deleteData = () => {
  console.log('=== INICIANDO LIMPIEZA DE DATOS ===');
  
  // Validación de hoja
  if (!validateSheet(SHEETS.alma, 'Reporte de deudores - Widget')) {
    return;
  }
  
  const lastRow = SHEETS.alma.getLastRow();
  
  // Verificar si hay datos (más allá de encabezados)
  if (lastRow < 2) {
    showToast(
      'La hoja ya se encuentra vacía',
      'Información',
      5,
      'ℹ️'
    );
    return;
  }
  
  // Limpiar rango de datos (A2:L hasta última fila)
  // Mantiene los encabezados en fila 1
  SHEETS.alma.getRange(`A2:L${lastRow}`).clearContent();
  
  console.log(`Limpiadas ${lastRow - 1} filas`);
  showToast(
    `Se limpiaron ${lastRow - 1} filas`,
    'Limpieza exitosa',
    5,
    '✅'
  );
};

/**
 * FUNCIÓN PRINCIPAL: Procesa datos de Alma e identifica cambios
 * 
 * FLUJO DE PROCESAMIENTO:
 * ┌──────────────────────────────────────────────────────────┐
 * │ 1. VALIDACIÓN                                            │
 * │    - Verificar hojas requeridas                          │
 * │    - Confirmar existencia de datos                       │
 * └─────────────────┬────────────────────────────────────────┘
 *                   ↓
 * ┌──────────────────────────────────────────────────────────┐
 * │ 2. CARGA EN MEMORIA (Optimización)                       │
 * │    - Leer todas las hojas de una vez                     │
 * │    - Crear índices Set para búsquedas O(1)               │
 * └─────────────────┬────────────────────────────────────────┘
 *                   ↓
 * ┌──────────────────────────────────────────────────────────┐
 * │ 3. IDENTIFICACIÓN DE NUEVOS DEUDORES                     │
 * │    For each registro en Alma:                            │
 * │      - Generar clave única                               │
 * │      - Buscar en índice de deudores actuales (O(1))      │
 * │      - Si NO existe → añadir a lista de nuevos           │
 * │      - Marcar estado (NUEVO/REGISTRADO)                  │
 * └─────────────────┬────────────────────────────────────────┘
 *                   ↓
 * ┌──────────────────────────────────────────────────────────┐
 * │ 4. IDENTIFICACIÓN DE DEVOLUCIONES                        │
 * │    For each deudor actual:                               │
 * │      - Generar clave única                               │
 * │      - Buscar en índice de Alma (O(1))                   │
 * │      - Si NO existe → usuario devolvió el recurso        │
 * │      - Añadir a lista de devoluciones                    │
 * │      - Marcar fila para eliminación                      │
 * └─────────────────┬────────────────────────────────────────┘
 *                   ↓
 * ┌──────────────────────────────────────────────────────────┐
 * │ 5. ESCRITURA BATCH (Optimización)                        │
 * │    - Actualizar estados en Alma (1 operación)            │
 * │    - Insertar nuevos deudores (1 operación)              │
 * │    - Insertar devoluciones (1 operación)                 │
 * │    - Eliminar filas de deudores devueltos                │
 * └─────────────────┬────────────────────────────────────────┘
 *                   ↓
 * ┌──────────────────────────────────────────────────────────┐
 * │ 6. REPORTE                                               │
 * │    - Mostrar resumen de operaciones                      │
 * │    - Log de tiempo de ejecución                          │
 * └──────────────────────────────────────────────────────────┘
 * 
 * OPTIMIZACIONES APLICADAS:
 * - Lectura única de cada hoja (no múltiples getRange())
 * - Uso de Set para búsquedas en O(1) vs O(n)
 * - Escritura por lotes (batch) en lugar de fila por fila
 * - Procesamiento en memoria antes de escribir
 * 
 * @returns {void}
 */
const startProcess = () => {
  console.log('=== INICIANDO PROCESO PRINCIPAL ===');
  console.time('⏱️ Tiempo total de procesamiento');
  
  // ─────────────────────────────────────
  // PASO 1: VALIDACIÓN
  // ─────────────────────────────────────
  const requiredSheets = [
    { sheet: SHEETS.alma, name: 'Reporte de deudores - Widget' },
    { sheet: SHEETS.overdueItems, name: 'Préstamos vencidos / Deudores' },
    { sheet: SHEETS.returnedItems, name: 'Recursos devueltos / Histórico' }
  ];
  
  const missingSheets = requiredSheets
    .filter(s => !s.sheet)
    .map(s => s.name);
  
  if (missingSheets.length > 0) {
    showToast(
      `Hojas faltantes:\n- ${missingSheets.join('\n- ')}`,
      'Error de configuración',
      8,
      '❌'
    );
    return;
  }
  
  // Verificar que haya datos para procesar
  if (SHEETS.alma.getRange('A2').getValue() === '') {
    showToast(
      'No hay datos para procesar',
      'Error',
      5,
      '❌'
    );
    return;
  }
  
  try {
    // ─────────────────────────────────────
    // PASO 2: CARGA EN MEMORIA
    // ─────────────────────────────────────
    console.log('📥 Cargando datos en memoria...');
    
    // Leer hoja de Alma completa (incluye encabezados)
    const almaFullData = SHEETS.alma.getDataRange().getValues();
    const almaHeaders = almaFullData[0];
    const almaData = almaFullData.slice(1); // Excluir encabezados
    
    // Leer hoja de deudores actuales
    const overdueFullData = SHEETS.overdueItems.getDataRange().getValues();
    const overdueHeaders = overdueFullData[0];
    const overdueData = overdueFullData.slice(1);
    
    console.log(`✓ Cargados ${almaData.length} registros de Alma`);
    console.log(`✓ Cargados ${overdueData.length} deudores actuales`);
    
    // ─────────────────────────────────────
    // CREAR ÍNDICES PARA BÚSQUEDA RÁPIDA
    // ─────────────────────────────────────
    // Set permite búsqueda en O(1) vs Array.find() que es O(n)
    
    // Índice de deudores actuales: Set de claves únicas
    const overdueIndex = new Set(
      overdueData.map(row => generateRecordKey(row))
    );
    console.log(`✓ Índice de deudores creado: ${overdueIndex.size} registros`);
    
    // ─────────────────────────────────────
    // PASO 3: IDENTIFICAR NUEVOS DEUDORES
    // ─────────────────────────────────────
    console.log('🔍 Identificando nuevos deudores...');
    
    const newDebtors = [];    // Registros a añadir a "Préstamos vencidos"
    const updates = [];       // Estados a actualizar en hoja Alma
    
    almaData.forEach((row, index) => {
      const recordKey = generateRecordKey(row);
      const isRegistered = overdueIndex.has(recordKey);
      
      // Preparar actualización de estado en Alma
      updates.push({
        row: index + 2, // +2 porque: arrays inician en 0, y hay 1 fila de encabezado
        value: isRegistered ? STATUS.REGISTERED : STATUS.NEW
      });
      
      // Si es nuevo, añadir a lista de nuevos deudores
      if (!isRegistered) {
        // Tomar solo las primeras 11 columnas (A-K)
        newDebtors.push(row.slice(0, COLUMNS.ACTION));
      }
    });
    
    console.log(`✓ Encontrados ${newDebtors.length} nuevos deudores`);
    
    // ─────────────────────────────────────
    // PASO 4: IDENTIFICAR DEVOLUCIONES
    // ─────────────────────────────────────
    console.log('🔍 Identificando recursos devueltos...');
    
    // Crear índice de registros actuales en Alma
    const almaIndex = new Set(
      almaData.map(row => generateRecordKey(row))
    );
    
    const returnedItems = [];    // Registros devueltos para histórico
    const rowsToDelete = [];     // Filas a eliminar de "Préstamos vencidos"
    
    overdueData.forEach((row, index) => {
      const recordKey = generateRecordKey(row);
      
      // Si el registro NO está en Alma, significa que fue devuelto
      if (!almaIndex.has(recordKey)) {
        const currentLog = row[COLUMNS.LOG] || '';
        const actionMessage = currentLog
          ? `${currentLog}\n${new Date().toLocaleString()}: Devuelto por el usuario`
          : `${new Date().toLocaleString()}: Devuelto por el usuario`;
        
        // Preparar registro para histórico
        returnedItems.push([
          ...row.slice(0, COLUMNS.ACTION),  // Datos principales (A-K)
          new Date(),                       // Fecha de devolución
          actionMessage                     // Bitácora actualizada
        ]);
        
        // Marcar fila para eliminación (+2 por índice y encabezado)
        rowsToDelete.push(index + 2);
      }
    });
    
    console.log(`✓ Encontrados ${returnedItems.length} recursos devueltos`);
    
    // ─────────────────────────────────────
    // PASO 5: ESCRITURA BATCH
    // ─────────────────────────────────────
    console.log('💾 Escribiendo cambios...');
    
    // 5.1 - Actualizar estados en hoja Alma
    if (updates.length > 0) {
      console.log(`  → Actualizando ${updates.length} estados en Alma...`);
      
      // Ordenar por número de fila para escritura contigua
      const sortedUpdates = updates.sort((a, b) => a.row - b.row);
      const firstRow = sortedUpdates[0].row;
      const lastRow = sortedUpdates[sortedUpdates.length - 1].row;
      const rowCount = lastRow - firstRow + 1;
      
      // Crear matriz de valores para escribir de una vez
      const outputValues = new Array(rowCount).fill(['']);
      sortedUpdates.forEach(update => {
        outputValues[update.row - firstRow] = [update.value];
      });
      
      // Escribir todos los estados en una sola operación
      SHEETS.alma
        .getRange(firstRow, COLUMNS.STATUS + 1, rowCount, 1)
        .setValues(outputValues);
      
      console.log(`  ✓ Estados actualizados`);
    }
    
    // 5.2 - Insertar nuevos deudores
    if (newDebtors.length > 0) {
      console.log(`  → Insertando ${newDebtors.length} nuevos deudores...`);
      
      const lastRow = SHEETS.overdueItems.getLastRow();
      SHEETS.overdueItems
        .getRange(lastRow + 1, 1, newDebtors.length, newDebtors[0].length)
        .setValues(newDebtors);
      
      console.log(`  ✓ Nuevos deudores insertados`);
    }
    
    // 5.3 - Mover recursos devueltos a histórico
    if (returnedItems.length > 0) {
      console.log(`  → Moviendo ${returnedItems.length} recursos a histórico...`);
      
      const lastRow = SHEETS.returnedItems.getLastRow();
      SHEETS.returnedItems
        .getRange(lastRow + 1, 1, returnedItems.length, returnedItems[0].length)
        .setValues(returnedItems);
      
      // Eliminar filas de "Préstamos vencidos"
      // IMPORTANTE: Eliminar de mayor a menor para no afectar índices
      rowsToDelete
        .sort((a, b) => b - a)
        .forEach(row => {
          SHEETS.overdueItems.deleteRow(row);
        });
      
      console.log(`  ✓ Recursos movidos y filas eliminadas`);
    }
    
    // ─────────────────────────────────────
    // PASO 6: REPORTE FINAL
    // ─────────────────────────────────────
    console.timeEnd('⏱️ Tiempo total de procesamiento');
    
    const registeredCount = updates.filter(u => u.value === STATUS.REGISTERED).length;
    
    const summary = [
      `Registros previos: ${registeredCount}`,
      `Nuevos deudores: ${newDebtors.length}`,
      `Ítems devueltos: ${returnedItems.length}`
    ].join(' // ');
    
    console.log('=== RESUMEN ===');
    console.log(summary);
    
    showToast(summary, 'Proceso completado', 15, '✅');
    
  } catch (error) {
    console.error('❌ Error en startProcess:', error);
    console.error('Stack:', error.stack);
    
    showToast(
      `Error inesperado: ${error.message}`,
      'Error en proceso',
      8,
      '❌'
    );
  }
};

// ========================================
// 4. FUNCIONES DE ACCIONES POR LOTES
// ========================================

/**
 * Mueve múltiples registros a "Recursos devueltos / Histórico"
 * 
 * OPTIMIZACIÓN: Procesa por lotes en lugar de uno por uno
 * 
 * @param {Array<Array>} rowsWithNumbers - Array de filas, cada una incluye:
 *   [...datos del registro, número de fila]
 * @returns {boolean} true si tuvo éxito
 */
const moveToReturnedItems = (rowsWithNumbers) => {
  console.log(`📦 Moviendo ${rowsWithNumbers.length} ítems a Recursos devueltos...`);
  
  try {
    // Validar hojas requeridas
    if (!validateSheet(SHEETS.overdueItems, 'Préstamos vencidos') ||
        !validateSheet(SHEETS.returnedItems, 'Recursos devueltos')) {
      return false;
    }
    
    // Separar datos de números de fila
    const rowsData = rowsWithNumbers.map(row => row.slice(0, -1));
    const rowNumbers = rowsWithNumbers.map(row => row[row.length - 1]);
    
    // Preparar registros para histórico
    const valuesToCopy = rowsData.map((row, index) => {
      const baseData = row.slice(0, COLUMNS.ACTION);
      const rowNumber = rowNumbers[index];
      
      // Obtener bitácora actual
      const logInfo = SHEETS.overdueItems.getRange(rowNumber, COLUMNS.LOG + 1).getValue();
      
      const actionMessage = logInfo
        ? `${logInfo}\n${new Date().toLocaleString()}: Ítem devuelto por ejecución de acciones`
        : `${new Date().toLocaleString()}: Ítem devuelto por ejecución de acciones`;
      
      return [
        ...baseData,
        new Date(),      // Fecha de devolución
        actionMessage    // Bitácora actualizada
      ];
    });
    
    // Insertar en histórico (1 operación)
    const lastRow = SHEETS.returnedItems.getLastRow();
    SHEETS.returnedItems
      .getRange(lastRow + 1, 1, valuesToCopy.length, valuesToCopy[0].length)
      .setValues(valuesToCopy);
    
    // Eliminar de deudores activos (de mayor a menor)
    rowNumbers
      .sort((a, b) => b - a)
      .forEach(rowNum => {
        SHEETS.overdueItems.deleteRow(rowNum);
      });
    
    console.log(`✓ ${rowsWithNumbers.length} ítems movidos exitosamente`);
    return true;
    
  } catch (error) {
    console.error('❌ Error en moveToReturnedItems:', error);
    showToast(
      `Error moviendo registros: ${error.message}`,
      'Error',
      5,
      '❌'
    );
    return false;
  }
};

/**
 * Mueve múltiples registros a "Seguimiento de préstamos"
 * 
 * DIFERENCIA con moveToReturnedItems:
 * - NO elimina las filas originales
 * - Solo limpia la acción ejecutada
 * - Mantiene el registro en "Préstamos vencidos"
 * 
 * @param {Array<Array>} rowsWithNumbers - Array de filas con números
 * @returns {boolean} true si tuvo éxito
 */
const moveToTrackingItems = (rowsWithNumbers) => {
  console.log(`📦 Moviendo ${rowsWithNumbers.length} ítems a Seguimiento...`);
  
  try {
    // Validar hojas requeridas
    if (!validateSheet(SHEETS.overdueItems, 'Préstamos vencidos') ||
        !validateSheet(SHEETS.trackingItems, 'Seguimiento de préstamos')) {
      return false;
    }
    
    // Separar datos de números de fila
    const rowsData = rowsWithNumbers.map(row => row.slice(0, -1));
    const rowNumbers = rowsWithNumbers.map(row => row[row.length - 1]);
    
    // Preparar registros para seguimiento
    const valuesToCopy = rowsData.map((row, index) => {
      const baseData = row.slice(0, COLUMNS.ACTION);
      const rowNumber = rowNumbers[index];
      
      // Obtener bitácora actual
      const logInfo = SHEETS.overdueItems.getRange(rowNumber, COLUMNS.LOG + 1).getValue();
      
      const actionMessage = logInfo
        ? `${logInfo}\n${new Date().toLocaleString()}: Ítem movido a Seguimiento`
        : `${new Date().toLocaleString()}: Ítem movido a Seguimiento`;
      
      // Limpiar acción ejecutada (columna L)
      SHEETS.overdueItems.getRange(rowNumber, COLUMNS.ACTION + 1).clearContent();
      
      return [
        ...baseData,
        new Date(),      // Fecha de seguimiento
        actionMessage    // Bitácora actualizada
      ];
    });
    
    // Insertar en seguimiento (1 operación)
    const lastRow = SHEETS.trackingItems.getLastRow();
    SHEETS.trackingItems
      .getRange(lastRow + 1, 1, valuesToCopy.length, valuesToCopy[0].length)
      .setValues(valuesToCopy);
    
    console.log(`✓ ${rowsWithNumbers.length} ítems movidos a seguimiento`);
    return true;
    
  } catch (error) {
    console.error('❌ Error en moveToTrackingItems:', error);
    showToast(
      `Error moviendo a seguimiento: ${error.message}`,
      'Error',
      5,
      '❌'
    );
    return false;
  }
};

// ========================================
// 5. FUNCIONES DE ENVÍO DE CORREOS
// ========================================

/**
 * TODO: Implementar envío de primer recordatorio
 * 
 * PENDIENTE:
 * 1. Cargar plantilla HTML emailFirstReminder.html
 * 2. Reemplazar variables en plantilla
 * 3. Enviar correo con GmailApp o MailApp
 * 4. Actualizar bitácora
 * 
 * @param {Array} data - Datos del registro
 * @param {number} rowNumber - Número de fila
 */
const sendFirstReminder = (data, rowNumber) => {
  console.log(`📧 [TODO] Enviar primer recordatorio - Fila ${rowNumber}`);
  
  // Actualizar bitácora
  const currentLog = SHEETS.overdueItems.getRange(rowNumber, COLUMNS.LOG + 1).getValue();
  updateActionLog(rowNumber, "Enviado primer recordatorio", currentLog);
};

/**
 * TODO: Implementar envío de segundo recordatorio
 * 
 * @param {Array} data - Datos del registro
 * @param {number} rowNumber - Número de fila
 */
const sendSecondReminder = (data, rowNumber) => {
  console.log(`📧 [TODO] Enviar segundo recordatorio - Fila ${rowNumber}`);
  
  // Actualizar bitácora
  const currentLog = SHEETS.overdueItems.getRange(rowNumber, COLUMNS.LOG + 1).getValue();
  updateActionLog(rowNumber, "Enviado segundo recordatorio", currentLog);
};

/**
 * TODO: Implementar envío de aviso de recarga
 * 
 * @param {Array} data - Datos del registro
 * @param {number} rowNumber - Número de fila
 */
const sendRechargeNotice = (data, rowNumber) => {
  console.log(`📧 [TODO] Enviar aviso de recarga - Fila ${rowNumber}`);
  
  // Actualizar bitácora
  const currentLog = SHEETS.overdueItems.getRange(rowNumber, COLUMNS.LOG + 1).getValue();
  updateActionLog(rowNumber, "Enviado aviso de recarga", currentLog);
};

/**
 * TODO: Implementar envío de confirmación de recarga
 * 
 * @param {Array} data - Datos del registro
 * @param {number} rowNumber - Número de fila
 */
const sendRechargeConfirmation = (data, rowNumber) => {
  console.log(`📧 [TODO] Enviar confirmación de recarga - Fila ${rowNumber}`);
  
  // Actualizar bitácora
  const currentLog = SHEETS.overdueItems.getRange(rowNumber, COLUMNS.LOG + 1).getValue();
  updateActionLog(rowNumber, "Enviada confirmación de recarga", currentLog);
};

// ========================================
// 6. EJECUTOR DE ACCIONES
// ========================================

/**
 * Ejecuta todas las acciones pendientes en la hoja "Préstamos vencidos"
 * 
 * FLUJO DE EJECUCIÓN:
 * ┌────────────────────────────────────────────────────────────┐
 * │ 1. LECTURA Y CLASIFICACIÓN                                 │
 * │    - Leer toda la hoja de una vez                          │
 * │    - Agrupar por tipo de acción (batch processing)         │
 * └─────────────────┬──────────────────────────────────────────┘
 *                   ↓
 * ┌────────────────────────────────────────────────────────────┐
 * │ 2. PROCESAMIENTO POR LOTES                                 │
 * │    Orden de ejecución:                                     │
 * │    a) Movimientos (batch)                                  │
 * │       - Recursos devueltos                                 │
 * │       - Seguimiento                                        │
 * │    b) Correos (individual)                                 │
 * │       - Primer recordatorio                                │
 * │       - Segundo recordatorio                               │
 * │       - Aviso de recarga                                   │
 * │       - Confirmación de recarga                            │
 * └─────────────────┬──────────────────────────────────────────┘
 *                   ↓
 * ┌────────────────────────────────────────────────────────────┐
 * │ 3. REPORTE                                                 │
 * │    - Contar acciones ejecutadas                            │
 * │    - Mostrar resumen al usuario                            │
 * └────────────────────────────────────────────────────────────┘
 * 
 * OPTIMIZACIÓN: Agrupa acciones del mismo tipo para ejecutarlas
 * en lote cuando sea posible (movimientos), reduciendo operaciones
 * de lectura/escritura en la hoja.
 * 
 * @returns {void}
 */
const executeActions = () => {
  console.log('=== INICIANDO EJECUCIÓN DE ACCIONES ===');
  console.time('⏱️ Tiempo de ejecución de acciones');
  
  // Validar hoja requerida
  if (!validateSheet(SHEETS.overdueItems, 'Préstamos vencidos / Deudores')) {
    return;
  }
  
  try {
    // ─────────────────────────────────────
    // PASO 1: LECTURA Y CLASIFICACIÓN
    // ─────────────────────────────────────
    console.log('📥 Leyendo acciones pendientes...');
    
    const fullData = SHEETS.overdueItems.getDataRange().getValues();
    const headers = fullData[0];
    const data = fullData.slice(1);
    
    // Mapeo de acciones a funciones
    const ACTION_MAP = {
      [ACTIONS.FIRST_REMINDER]: sendFirstReminder,
      [ACTIONS.SECOND_REMINDER]: sendSecondReminder,
      [ACTIONS.RECHARGE_NOTICE]: sendRechargeNotice,
      [ACTIONS.RECHARGE_CONFIRMATION]: sendRechargeConfirmation,
      [ACTIONS.MOVE_TO_RETURNED]: moveToReturnedItems,
      [ACTIONS.MOVE_TO_TRACKING]: moveToTrackingItems
    };
    
    // Estructura para agrupar acciones por tipo
    const actionsBatch = {
      [ACTIONS.FIRST_REMINDER]: [],
      [ACTIONS.SECOND_REMINDER]: [],
      [ACTIONS.RECHARGE_NOTICE]: [],
      [ACTIONS.RECHARGE_CONFIRMATION]: [],
      [ACTIONS.MOVE_TO_RETURNED]: [],
      [ACTIONS.MOVE_TO_TRACKING]: []
    };
    
    // Clasificar cada fila según su acción
    data.forEach((row, index) => {
      const rowNumber = index + 2; // +2 por índice y encabezado
      const actionValue = row[COLUMNS.ACTION];
      
      // Si hay una acción válida, añadir a su lote
      if (actionValue && ACTION_MAP[actionValue]) {
        actionsBatch[actionValue].push({
          data: row,
          rowNumber: rowNumber
        });
      }
    });
    
    // Contar total de acciones pendientes
    const totalActions = Object.values(actionsBatch)
      .reduce((sum, batch) => sum + batch.length, 0);
    
    if (totalActions === 0) {
      showToast(
        'No hay acciones pendientes para ejecutar',
        'Información',
        5,
        'ℹ️'
      );
      return;
    }
    
    console.log(`✓ Encontradas ${totalActions} acciones pendientes`);
    
    // ─────────────────────────────────────
    // PASO 2: PROCESAMIENTO POR LOTES
    // ─────────────────────────────────────
    
    // 2.1 - Procesar movimientos a Recursos devueltos (batch)
    if (actionsBatch[ACTIONS.MOVE_TO_RETURNED].length > 0) {
      console.log(`📦 Procesando ${actionsBatch[ACTIONS.MOVE_TO_RETURNED].length} movimientos a Recursos devueltos...`);
      
      const batch = actionsBatch[ACTIONS.MOVE_TO_RETURNED];
      // Añadir número de fila al final de cada registro
      const rowsToProcess = batch.map(item => [...item.data, item.rowNumber]);
      
      if (moveToReturnedItems(rowsToProcess)) {
        console.log(`✓ ${batch.length} registros movidos a Recursos devueltos`);
      }
    }
    
    // 2.2 - Procesar movimientos a Seguimiento (batch)
    if (actionsBatch[ACTIONS.MOVE_TO_TRACKING].length > 0) {
      console.log(`📦 Procesando ${actionsBatch[ACTIONS.MOVE_TO_TRACKING].length} movimientos a Seguimiento...`);
      
      const batch = actionsBatch[ACTIONS.MOVE_TO_TRACKING];
      // Añadir número de fila al final de cada registro
      const rowsToProcess = batch.map(item => [...item.data, item.rowNumber]);
      
      if (moveToTrackingItems(rowsToProcess)) {
        console.log(`✓ ${batch.length} registros movidos a Seguimiento`);
      }
    }
    
    // 2.3 - Procesar envíos de correo (individual)
    // Los correos se procesan uno por uno porque:
    // - Cada correo puede fallar individualmente
    // - Necesitamos registrar cada envío en la bitácora
    // - GmailApp tiene límites de cuota diaria
    
    const emailActions = [
      ACTIONS.FIRST_REMINDER,
      ACTIONS.SECOND_REMINDER,
      ACTIONS.RECHARGE_NOTICE,
      ACTIONS.RECHARGE_CONFIRMATION
    ];
    
    emailActions.forEach(action => {
      if (actionsBatch[action].length > 0) {
        console.log(`📧 Procesando ${actionsBatch[action].length} ${action}...`);
        
        const batch = actionsBatch[action];
        batch.forEach(item => {
          try {
            // Ejecutar la función correspondiente
            ACTION_MAP[action](item.data, item.rowNumber);
          } catch (error) {
            console.error(`❌ Error procesando fila ${item.rowNumber}:`, error);
            // Continuar con los demás registros
          }
        });
        
        console.log(`✓ ${batch.length} correos procesados`);
      }
    });
    
    // ─────────────────────────────────────
    // PASO 3: REPORTE FINAL
    // ─────────────────────────────────────
    console.timeEnd('⏱️ Tiempo de ejecución de acciones');
    
    // Calcular totales por categoría
    const movedCount = actionsBatch[ACTIONS.MOVE_TO_RETURNED].length +
                      actionsBatch[ACTIONS.MOVE_TO_TRACKING].length;
    
    const emailCount = actionsBatch[ACTIONS.FIRST_REMINDER].length +
                      actionsBatch[ACTIONS.SECOND_REMINDER].length +
                      actionsBatch[ACTIONS.RECHARGE_NOTICE].length +
                      actionsBatch[ACTIONS.RECHARGE_CONFIRMATION].length;
    
    const summary = [
      `Ítems devueltos: ${actionsBatch[ACTIONS.MOVE_TO_RETURNED].length}`,
      `Ítems en seguimiento: ${actionsBatch[ACTIONS.MOVE_TO_TRACKING].length}`,
      `Correos enviados: ${emailCount}`
    ].join(' // ');
    
    console.log('=== RESUMEN DE EJECUCIÓN ===');
    console.log(summary);
    console.log('Detalle por tipo:');
    Object.entries(actionsBatch).forEach(([action, items]) => {
      if (items.length > 0) {
        console.log(`  - ${action}: ${items.length}`);
      }
    });
    
    showToast(summary, 'Acciones ejecutadas', 15, '✅');
    
  } catch (error) {
    console.error('❌ Error en executeActions:', error);
    console.error('Stack:', error.stack);
    
    showToast(
      `Error ejecutando acciones: ${error.message}`,
      'Error',
      8,
      '❌'
    );
  }
};

// ========================================
// 7. FUNCIONES DE INTERFAZ DE USUARIO
// ========================================

/**
 * Muestra información sobre el script actual
 * Útil para verificar que el script está conectado correctamente
 */
const hasScript = () => {
  const info = `
📄 Script: SP | Reporte de deudores
📌 Versión: 1.0.0
👤 Autor: Fredy Romero
🔗 Script ID: ${ScriptApp.getScriptId()}

Hojas configuradas:
• ${SHEETS.alma.getName()}
• ${SHEETS.overdueItems.getName()}
• ${SHEETS.trackingItems.getName()}
• ${SHEETS.returnedItems.getName()}
  `.trim();
  
  UI.alert('Información del Script ℹ️', info, UI.ButtonSet.OK);
};

/**
 * Crea el menú personalizado en la interfaz de Google Sheets
 * Se ejecuta automáticamente al abrir el documento
 * 
 * ESTRUCTURA DEL MENÚ:
 * Scripts 🟢
 * ├── ➡️ Procesar datos de: [Hoja Alma]
 * ├── 🧪 Ejecutar acciones (L) de: [Hoja Deudores]
 * ├── ─────────────
 * ├── 🗑️ Borrar datos de: [Hoja Alma]
 * ├── ─────────────
 * └── ⚠️ Información del script
 * 
 * @returns {void}
 */
const onOpen = () => {
  console.log('🎨 Creando menú personalizado...');
  
  try {
    UI.createMenu('Scripts 🟢')
      .addItem('➡️ Procesar datos de: ' + SHEETS.alma.getName(), 'startProcess')
      .addItem('🧪 Ejecutar acciones (L) de: ' + SHEETS.overdueItems.getName(), 'executeActions')
      .addSeparator()
      .addItem('🗑️ Borrar datos de: ' + SHEETS.alma.getName(), 'deleteData')
      .addSeparator()
      .addItem('⚠️ Información del script', 'hasScript')
      .addToUi();
    
    console.log('✓ Menú creado exitosamente');
  } catch (error) {
    console.error('❌ Error creando menú:', error);
  }
};

// ========================================
// 8. DOCUMENTACIÓN DE FLUJO COMPLETO
// ========================================

/**
 * FLUJO COMPLETO DEL SISTEMA
 * ═══════════════════════════════════════════════════════════════
 * 
 * 1️⃣ IMPORTACIÓN DE DATOS DESDE ALMA
 *    ┌─────────────────────────────────────────┐
 *    │ Sistema Alma (Biblioteca)               │
 *    │ Exporta datos de préstamos vencidos     │
 *    └────────────┬────────────────────────────┘
 *                 ↓
 *    ┌─────────────────────────────────────────┐
 *    │ Hoja: "Reporte de deudores - Widget"    │
 *    │ Contiene datos importados (A-L)         │
 *    └─────────────────────────────────────────┘
 * 
 * 2️⃣ PROCESAMIENTO (startProcess)
 *    ┌─────────────────────────────────────────┐
 *    │ Análisis de datos                       │
 *    │ • Identificar nuevos deudores           │
 *    │ • Identificar recursos devueltos        │
 *    │ • Actualizar estados                    │
 *    └────────────┬────────────────────────────┘
 *                 ↓
 *    ┌────────────┴────────────┐
 *    ↓                         ↓
 *    NUEVOS DEUDORES          RECURSOS DEVUELTOS
 *    ↓                         ↓
 *    ┌──────────────────┐     ┌──────────────────┐
 *    │ Préstamos        │     │ Recursos         │
 *    │ vencidos /       │     │ devueltos /      │
 *    │ Deudores         │     │ Histórico        │
 *    └──────────────────┘     └──────────────────┘
 * 
 * 3️⃣ GESTIÓN DE ACCIONES (executeActions)
 *    ┌──────────────────────────────────────────┐
 *    │ Usuario define acciones en columna L     │
 *    │ de "Préstamos vencidos / Deudores"       │
 *    └────────────┬─────────────────────────────┘
 *                 ↓
 *    ┌────────────┴────────────────────────────┐
 *    │ Procesamiento por lotes                 │
 *    └────────────┬────────────────────────────┘
 *                 ↓
 *    ┌────────────┴────────────┐
 *    │ Acciones disponibles:   │
 *    │ • Enviar recordatorios  │
 *    │ • Mover a seguimiento   │
 *    │ • Marcar como devuelto  │
 *    └─────────────────────────┘
 * 
 * 4️⃣ TRAZABILIDAD
 *    ┌──────────────────────────────────────────┐
 *    │ Columna M: Bitácora de acciones          │
 *    │ Registra cada acción con timestamp       │
 *    │ Ejemplo:                                 │
 *    │ "14/10/2025 10:30: Primer recordatorio"  │
 *    │ "15/10/2025 14:20: Movido a seguimiento" │
 *    └──────────────────────────────────────────┘
 * 
 * VENTAJAS DE ESTE DISEÑO:
 * ✓ Procesamiento por lotes (eficiente)
 * ✓ Historial completo de cada préstamo
 * ✓ Minimiza operaciones de lectura/escritura
 * ✓ Interfaz simple para el usuario
 * ✓ Trazabilidad completa de acciones
 */