/**
 * CONSTANTES Y CONFIGURACIÓN
 */
const UI = SpreadsheetApp.getUi();
const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEETS = {
  alma: SS.getSheetByName("Reporte de deudores - Widget"),
  prestamosVencidos: SS.getSheetByName("Préstamos vencidos / Deudores"),
  seguimientoPrestamos: SS.getSheetByName("Seguimiento de préstamos"),
  recursosDevueltos: SS.getSheetByName("Recursos devueltos / Histórico")
};
const AUTHORIZED_USER = "fromeror@continental.edu.pe";

// **********************************************
// FUNCIONES PRINCIPALES
// **********************************************

/**
 * Prepara la hoja para nuevos datos limpiando contenido previo
 */
const resetSheetForNewData = () => {
  if (!SHEETS.alma) {
    UI.alert('Error', 'Hoja "Reporte de deudores - Widget" no encontrada', UI.ButtonSet.OK);
    return;
  }

  const lastRow = SHEETS.alma.getLastRow();
  if (lastRow < 2) {
    UI.alert('Info', 'La hoja ya está vacía', UI.ButtonSet.OK);
    return;
  }

  SHEETS.alma.getRange(`A2:M${lastRow}`).clearContent();
  SpreadsheetApp.flush();
  UI.alert('Éxito', `Se limpiaron ${lastRow - 1} filas`, UI.ButtonSet.OK);
};

/**
 * Procesa los datos optimizando las operaciones de lectura/escritura
 */
const startProcess = () => {
  // Validación inicial de hojas
  if (!SHEETS.alma || !SHEETS.prestamosVencidos || !SHEETS.recursosDevueltos) {
    const missingSheets = [];
    if (!SHEETS.alma) missingSheets.push("Reporte de deudores - Widget");
    if (!SHEETS.prestamosVencidos) missingSheets.push("Préstamos vencidos / Deudores");
    if (!SHEETS.recursosDevueltos) missingSheets.push("Recursos devueltos / Histórico");

    UI.alert(
      "Error de configuración",
      `No se encontraron las siguientes hojas:\n\n- ${missingSheets.join("\n- ")}\n\nVerifica los nombres de las hojas.`,
      UI.ButtonSet.OK
    );
    return;
  }

  try {
    console.time("Procesamiento datos");

    if (SHEETS.alma.getRange('A2').getValue() === "") {
      UI.alert("Información", "No hay datos para procesar.", UI.ButtonSet.OK);
      return;
    }

    // Carga de datos en memoria
    const [almaHeaders, ...almaData] = SHEETS.alma.getDataRange().getValues();
    const [prestamosHeaders, ...prestamosData] = SHEETS.prestamosVencidos.getDataRange().getValues();

    // Crear índices para búsquedas rápidas
    const prestamoIndex = new Set(
      prestamosData.map(row => `${row[2]}__${row[8]}__${row[9]}__${row[10]}`)
    );

    // Procesamiento de datos
    const newDebtors = [];
    const updates = [];
    const returnedItems = [];
    let rowsToDelete = [];

    almaData.forEach((row, i) => {
      const recordKey = `${row[2]}__${row[8]}__${row[9]}__${row[10]}`;
      const isRegistered = prestamoIndex.has(recordKey);

      updates.push({
        row: i + 2,
        value: isRegistered ? "YA REGISTRADO" : "NUEVO DEUDOR"
      });

      if (!isRegistered) newDebtors.push(row.slice(0, 12));
    });

    const almaIndex = new Set(
      almaData.map(row => `${row[2]}__${row[8]}__${row[9]}__${row[10]}`)
    );

    prestamosData.forEach((row, i) => {
      const recordKey = `${row[2]}__${row[8]}__${row[9]}__${row[10]}`;
      if (!almaIndex.has(recordKey)) {
        returnedItems.push([...row.slice(0, 12), "Sí", new Date(), "Devuelto por el usuario"]);
        rowsToDelete.push(i + 2);
      }
    });

    if (updates.length) {
      const sortedUpdates = updates.sort((a, b) => a.row - b.row);
      const firstRow = sortedUpdates[0].row;
      const lastRow = sortedUpdates[sortedUpdates.length - 1].row;
      const rowCount = lastRow - firstRow + 1;

      const outputValues = new Array(rowCount).fill([""]);

      sortedUpdates.forEach(update => {
        outputValues[update.row - firstRow] = [update.value];
      });

      SHEETS.alma.getRange(firstRow, 13, rowCount, 1).setValues(outputValues);
    }

    if (newDebtors.length) {
      SHEETS.prestamosVencidos.getRange(
        SHEETS.prestamosVencidos.getLastRow() + 1, 1,
        newDebtors.length, newDebtors[0].length
      ).setValues(newDebtors);
    }

    if (returnedItems.length) {
      SHEETS.recursosDevueltos.getRange(
        SHEETS.recursosDevueltos.getLastRow() + 1, 1,
        returnedItems.length, returnedItems[0].length
      ).setValues(returnedItems);

      rowsToDelete.sort((a, b) => b - a).forEach(row => {
        SHEETS.prestamosVencidos.deleteRow(row);
      });
    }

    // Resultados
    console.timeEnd("Procesamiento datos");
    const summary = `
    Total registros previos: ${updates.filter(u => u.value === "YA REGISTRADO").length}
    Total nuevos deudores: ${newDebtors.length}
    Total ítems devueltos: ${returnedItems.length}
    `;

    console.log(summary);
    UI.alert("Resumen del Proceso", summary, UI.ButtonSet.OK);

  } catch (error) {
    console.error("Error en startProcess:", error);
    UI.alert(
      "Error en el proceso",
      `Ocurrió un error inesperado:\n\n${error.message}`,
      UI.ButtonSet.OK
    );
  }
};

// **********************************************
// FUNCIONES DE ACCIONES
// **********************************************

/**
 * Mueve registros a Recursos devueltos/Histórico (por lotes)
 */
const moverARecursosDevueltos = (rowsWithNumbers) => {
  try {
    if (!SHEETS.prestamosVencidos || !SHEETS.recursosDevueltos) {
      UI.alert("Error", "No se encontraron las hojas requeridas", UI.ButtonSet.OK);
      return false;
    }

    const rowsData = rowsWithNumbers.map(row => row.slice(0, -1));
    const rowNumbers = rowsWithNumbers.map(row => row[row.length - 1]);

    const valuesToCopy = rowsData.map(row => {
      const baseData = row.slice(0, 12);
      return [
        ...baseData,
        "Sí",
        new Date(),
        "Proceso automático"
      ];
    });

    const lastRow = SHEETS.recursosDevueltos.getLastRow();
    SHEETS.recursosDevueltos.getRange(lastRow + 1, 1, valuesToCopy.length, 15)
      .setValues(valuesToCopy);

    rowNumbers.sort((a, b) => b - a).forEach(rowNum => {
      SHEETS.prestamosVencidos.deleteRow(rowNum);
    });

    return true;

  } catch (error) {
    console.error("Error en moverARecursosDevueltos:", error);
    UI.alert(
      "Error",
      `Error moviendo registros a Recursos devueltos:\n${error.message}`,
      UI.ButtonSet.OK
    );
    return false;
  }
};

/**
 * Mueve registros a Seguimiento de préstamos (por lotes)
 */
const moverASeguimientoPrestamos = (rowsData) => {
  try {
    if (!SHEETS.prestamosVencidos || !SHEETS.seguimientoPrestamos) {
      UI.alert("Error", "No se encontraron las hojas requeridas", UI.ButtonSet.OK);
      return false;
    }

    const valuesToCopy = rowsData.map(row => row.slice(0, 12));

    const lastRow = SHEETS.seguimientoPrestamos.getLastRow();
    SHEETS.seguimientoPrestamos.getRange(lastRow + 1, 1, valuesToCopy.length, 12)
      .setValues(valuesToCopy);

    return true;

  } catch (error) {
    console.error("Error en moverASeguimientoPrestamos:", error);
    UI.alert(
      "Error",
      `Error moviendo registros a Seguimiento de préstamos:\n${error.message}`,
      UI.ButtonSet.OK
    );
    return false;
  }
};

// TODO: Implementar acciones

/**
 * Envía correo de primer recordatorio
 */
const enviarPrimerRecordatorio = (rows) => {
  // Implementación pendiente
};

/**
 * Envía correo de segundo recordatorio
 */
const enviarSegundoRecordatorio = (rows) => {
  // Implementación pendiente
};

/**
 * Envía correo de aviso de recarga
 */
const enviarAvisoRecarga = (rows) => {
  // Implementación pendiente
};

/**
 * Envía correo de confirmación de recarga
 */
const enviarConfirmacionRecarga = (rows) => {
  // Implementación pendiente
};

/**
 * Ejecuta acciones basadas en los valores de la columna N (14)
 */
const executeActions = () => {
  if (!SHEETS.prestamosVencidos) {
    UI.alert("Error", "No se encontró la hoja 'Préstamos vencidos / Deudores'", UI.ButtonSet.OK);
    return;
  }

  const data = SHEETS.prestamosVencidos.getDataRange().getValues();
  const headers = data.shift();

  const ACTION_MAP = {
    "Enviar correo: Primer recordatorio": enviarPrimerRecordatorio,
    "Enviar correo: Segundo recordatorio": enviarSegundoRecordatorio,
    "Enviar correo: Aviso de recarga": enviarAvisoRecarga,
    "Enviar correo: Confirmación de la recarga": enviarConfirmacionRecarga,
    "Mover a: Recursos devueltos / Histórico": moverARecursosDevueltos,
    "Mover a: Seguimiento de préstamos": moverASeguimientoPrestamos
  };

  const actionsBatch = {
    "Enviar correo: Primer recordatorio": [],
    "Enviar correo: Segundo recordatorio": [],
    "Enviar correo: Aviso de recarga": [],
    "Enviar correo: Confirmación de la recarga": [],
    "Mover a: Recursos devueltos / Histórico": [],
    "Mover a: Seguimiento de préstamos": []
  };

  data.forEach((row, index) => {
    const rowNumber = index + 2;
    const actionValue = row[13];

    if (actionValue && ACTION_MAP[actionValue]) {
      actionsBatch[actionValue].push({
        data: row,
        rowNumber: rowNumber
      });
    }
  });

  let processedCount = 0;

  if (actionsBatch["Mover a: Recursos devueltos / Histórico"].length > 0) {
    const batch = actionsBatch["Mover a: Recursos devueltos / Histórico"];
    const rowsToProcess = batch.map(item => [...item.data, item.rowNumber]);

    if (moverARecursosDevueltos(rowsToProcess)) {
      processedCount += batch.length;
      console.log(`Movidos ${batch.length} registros a Recursos devueltos`);
    }
  }

  if (actionsBatch["Mover a: Seguimiento de préstamos"].length > 0) {
    const batch = actionsBatch["Mover a: Seguimiento de préstamos"];
    const rowsToProcess = batch.map(item => item.data);

    if (moverASeguimientoPrestamos(rowsToProcess)) {
      processedCount += batch.length;
      console.log(`Movidos ${batch.length} registros a Seguimiento de préstamos`);
    }
  }

  ["Enviar correo: Primer recordatorio", "Enviar correo: Segundo recordatorio", "Enviar correo: Recarga de pensión"].forEach(action => {
    if (actionsBatch[action].length > 0) {
      const batch = actionsBatch[action];
      batch.forEach(item => {
        ACTION_MAP[action](item.data, item.rowNumber);
      });
      processedCount += batch.length;
      console.log(`Procesados ${batch.length} ${action}`);
    }
  });

  const summary = `Proceso completado:\n\n` +
    `- Movidos a Recursos devueltos: ${actionsBatch["Mover a: Recursos devueltos / Histórico"].length}\n` +
    `- Movidos a Seguimiento: ${actionsBatch["Mover a: Seguimiento de préstamos"].length}\n` +
    `- Correos enviados: ${
      actionsBatch["Enviar correo: Primer recordatorio"].length + 
      actionsBatch["Enviar correo: Segundo recordatorio"].length + 
      actionsBatch["Enviar correo: Aviso de recarga"].length +
      actionsBatch["Enviar correo: Confirmación de la recarga"].length
    }\n` +
    `\nTotal acciones: ${processedCount}`;

  UI.alert("Resumen de ejecución", summary, UI.ButtonSet.OK);
};

// **********************************************
// MENÚ
// **********************************************

/**
 * Crea el menú personalizado
 */
const onOpen = () => {
  const menu = UI.createMenu('Scripts 🟢')
    .addItem('🔄 Procesar reporte de Alma', 'startProcess')
    .addItem('⚡ Ejecutar acciones por ítem', 'executeActions')
    .addSeparator()
    .addItem('🗑️ Limpiar información', 'resetSheetForNewData');

  if (Session.getActiveUser().getEmail() === AUTHORIZED_USER) {
    menu
      .addSeparator()
      .addSubMenu(UI.createMenu('⚙️ Avanzado')
        .addItem('Mover a: Seguimiento de préstamos', 'moverASeguimientoPrestamos')
        .addItem('Mover a: Recursos devueltos', 'moverARecursosDevueltos'));
  }

  menu.addToUi();
};