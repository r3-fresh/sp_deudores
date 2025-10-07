/**
 * CONSTANTES Y CONFIGURACIÓN
 */
const UI = SpreadsheetApp.getUi();
const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEETS = {
  alma: SS.getSheetById(563966915),
  overdueItems: SS.getSheetById(491373272),
  trackingItems: SS.getSheetById(687630222),
  returnedItems: SS.getSheetById(1634827826),
};

// **********************************************
// FUNCIONES PRINCIPALES
// **********************************************

/**
 * Prepara la hoja para nuevos datos limpiando contenido previo
 */
const deleteData = () => {
  if (!SHEETS.alma) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `No se encontró la hoja ${SHEETS.alma.getSheetName()}.`,
      "Error en la configuración ❌",
      5
    );
    return;
  }

  const lastRow = SHEETS.alma.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `La hoja ${SHEETS.alma.getSheetName()} ya se encuentra vacía.`,
      "Información ⚠️",
      5
    );
    return;
  }

  SHEETS.alma.getRange(`A2:L${lastRow}`).clearContent();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Se limpiaron ${lastRow - 1} filas de la hoja ${SHEETS.alma.getSheetName()}.`,
    "Éxito ✅",
    5
  );
};

/**
 * Procesa los datos optimizando las operaciones de lectura/escritura
 */
const startProcess = () => {
  // Validación inicial de hojas
  if (!SHEETS.alma || !SHEETS.overdueItems || !SHEETS.returnedItems) {
    const missingSheets = [];
    if (!SHEETS.alma) missingSheets.push(SHEETS.alma.getSheetName());
    if (!SHEETS.overdueItems) missingSheets.push(SHEETS.overdueItems.getSheetName());
    if (!SHEETS.returnedItems) missingSheets.push(SHEETS.returnedItems.getSheetName());

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `No se encontraron las siguientes hojas: - ${missingSheets.join("\n- ")}. Verifica los IDs de las hojas.`,
      "Error en la configuración ❌",
      5
    );
    return;
  }

  try {
    console.time("Procesamiento datos");

    if (SHEETS.alma.getRange('A2').getValue() === "") {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `No hay datos para procesar en ${SHEETS.alma.getSheetName()}.`,
        "Error en los datos ❌",
        5
      );
      return;
    }

    // Carga de datos en memoria
    const [almaHeaders, ...almaData] = SHEETS.alma.getDataRange().getValues();
    const [overdueHeaders, ...overdueData] = SHEETS.overdueItems.getDataRange().getValues();

    // Crear índices para búsquedas rápidas
    const overdueIndex = new Set(
      overdueData.map(row => `${row[2]}__${row[8]}__${row[9]}__${row[10]}`)
    );

    // Procesamiento de datos
    const newDebtors = [];
    const updates = [];
    const returnedItems = [];
    let rowsToDelete = [];

    almaData.forEach((row, i) => {
      const recordKey = `${row[2]}__${row[8]}__${row[9]}__${row[10]}`;
      const isRegistered = overdueIndex.has(recordKey);

      updates.push({
        row: i + 2,
        value: isRegistered ? "YA REGISTRADO" : "NUEVO DEUDOR"
      });

      if (!isRegistered) newDebtors.push(row.slice(0, 11));
    });

    const almaIndex = new Set(
      almaData.map(row => `${row[2]}__${row[8]}__${row[9]}__${row[10]}`)
    );

    overdueData.forEach((row, i) => {
      const recordKey = `${row[2]}__${row[8]}__${row[9]}__${row[10]}`;
      if (!almaIndex.has(recordKey)) {
        const fecha = row[9].split('/');
        const mesNumero = fecha[1];
        const meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
        const mesTexto = meses[parseInt(mesNumero) - 1];
        returnedItems.push([...row.slice(0, 11), new Date(), "Devuelto por el usuario", mesTexto, mesNumero]);
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

      SHEETS.alma.getRange(firstRow, 12, rowCount, 1).setValues(outputValues);
    }

    if (newDebtors.length) {
      SHEETS.overdueItems.getRange(
        SHEETS.overdueItems.getLastRow() + 1, 1,
        newDebtors.length, newDebtors[0].length
      ).setValues(newDebtors);
    }

    if (returnedItems.length) {
      SHEETS.returnedItems.getRange(
        SHEETS.returnedItems.getLastRow() + 1, 1,
        returnedItems.length, returnedItems[0].length
      ).setValues(returnedItems);

      rowsToDelete.sort((a, b) => b - a).forEach(row => {
        SHEETS.overdueItems.deleteRow(row);
      });
    }

    // Resultados
    console.timeEnd("Procesamiento datos");
    const summary = `
    Total registros previos: ${updates.filter(u => u.value === "YA REGISTRADO").length} // 
    Total nuevos deudores: ${newDebtors.length} // 
    Total ítems devueltos: ${returnedItems.length}
    `;

    console.log(summary);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      summary,
      "Resumen del proceso ✅",
      15
    );

  } catch (error) {
    console.error("Error en startProcess:", error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Ocurrió un error inesperado: ${error.message}`,
      "Error en el proceso ❌",
      5
    );
  }
};

// **********************************************
// FUNCIONES DE ACCIONES
// **********************************************

/**
 * Mueve registros a Recursos devueltos/Histórico (por lotes)
 */
const moveToReturnedItems = (rowsWithNumbers) => {
  try {
    if (!SHEETS.overdueItems || !SHEETS.returnedItems) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "No se encontraron las hojas requeridas",
        "Error ❌",
        5
      );
      return false;
    }

    const rowsData = rowsWithNumbers.map(row => row.slice(0, -1));
    const rowNumbers = rowsWithNumbers.map(row => row[row.length - 1]);

    const valuesToCopy = rowsData.map(row => {
      const baseData = row.slice(0, 11);
      const fecha = row[9].split('/');
      const mesNumero = fecha[1];
      const meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
      const mesTexto = meses[parseInt(mesNumero) - 1];
      return [
        ...baseData,
        new Date(),
        "Proceso automático",
        mesTexto,
        mesNumero
      ];
    });

    const lastRow = SHEETS.returnedItems.getLastRow();
    SHEETS.returnedItems.getRange(lastRow + 1, 1, valuesToCopy.length, 15)
      .setValues(valuesToCopy);

    rowNumbers.sort((a, b) => b - a).forEach(rowNum => {
      SHEETS.overdueItems.deleteRow(rowNum);
    });

    return true;

  } catch (error) {
    console.error("Error en moveToReturnedItems:", error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Error moviendo registros a ${SHEETS.returnedItems.getSheetName()}: ${error.message}`,
      "Error ❌",
      5
    );
    return false;
  }
};

/**
 * Mueve registros a Seguimiento de préstamos (por lotes)
 */
const moveToTrackingItems = (rowsData) => {
  try {
    if (!SHEETS.overdueItems || !SHEETS.trackingItems) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "No se encontraron las hojas requeridas",
        "Error ❌",
        5
      );
      return false;
    }

    const valuesToCopy = rowsData.map(row => row.slice(0, 11));

    const lastRow = SHEETS.trackingItems.getLastRow();
    SHEETS.trackingItems.getRange(lastRow + 1, 1, valuesToCopy.length, 11)
      .setValues(valuesToCopy);

    return true;

  } catch (error) {
    console.error("Error en moveToTrackingItems:", error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Error moviendo registros a ${SHEETS.trackingItems.getSheetName()}: ${error.message}`,
      "Error ❌",
      5
    );
    return false;
  }
};

// TODO: Implementar acciones

/**
 * Envía correo de primer recordatorio
 */
const sendFirstReminder = (rows) => {
  UI.alert("sendFirstReminder" + JSON.stringify(rows))
};

/**
 * Envía correo de segundo recordatorio
 */
const sendSecondReminder = (rows) => {
  UI.alert("sendSecondReminder" + JSON.stringify(rows))
};

/**
 * Envía correo de aviso de recarga
 */
const sendRechargeNotice = (rows) => {
  UI.alert("sendRechargeNotice" + JSON.stringify(rows))
};

/**
 * Envía correo de confirmación de recarga
 */
const sendRechargeConfirmation = (rows) => {
  UI.alert("sendRechargeConfirmation" + JSON.stringify(rows))
};

/**
 * Ejecuta acciones basadas en los valores de la columna N (14)
 */
const executeActions = () => {
  if (!SHEETS.overdueItems) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "No se encontró la hoja 'Préstamos vencidos / Deudores'",
      "Error ❌",
      5
    );
    return;
  }

  const data = SHEETS.overdueItems.getDataRange().getValues();
  const headers = data.shift();

  const ACTION_MAP = {
    "✉️ Primer recordatorio": sendFirstReminder,
    "✉️ Segundo recordatorio": sendSecondReminder,
    "✉️ Aviso de recarga": sendRechargeNotice,
    "✉️ Confirmación de la recarga": sendRechargeConfirmation,
    "Ítem devuelto/encontrado": moveToReturnedItems,
    "Dar seguimiento al ítem": moveToTrackingItems
  };

  const actionsBatch = {
    "✉️ Primer recordatorio": [],
    "✉️ Segundo recordatorio": [],
    "✉️ Aviso de recarga": [],
    "✉️ Confirmación de la recarga": [],
    "Ítem devuelto/encontrado": [],
    "Dar seguimiento al ítem": []
  };

  data.forEach((row, index) => {
    const rowNumber = index + 2;
    const actionValue = row[11];

    if (actionValue && ACTION_MAP[actionValue]) {
      actionsBatch[actionValue].push({
        data: row,
        rowNumber: rowNumber
      });
    }
  });

  if (actionsBatch["Ítem devuelto/encontrado"].length > 0) {
    const batch = actionsBatch["Ítem devuelto/encontrado"];
    const rowsToProcess = batch.map(item => [...item.data, item.rowNumber]);

    if (moveToReturnedItems(rowsToProcess)) {
      console.log(`Movidos ${batch.length} registros a Recursos devueltos`);
    }
  }

  if (actionsBatch["Dar seguimiento al ítem"].length > 0) {
    const batch = actionsBatch["Dar seguimiento al ítem"];
    const rowsToProcess = batch.map(item => item.data);

    if (moveToTrackingItems(rowsToProcess)) {
      console.log(`Movidos ${batch.length} registros a Seguimiento de préstamos`);
    }
  }

  ["✉️ Primer recordatorio", "✉️ Segundo recordatorio", "✉️ Aviso de recarga", "✉️ Confirmación de la recarga"].forEach(action => {
    if (actionsBatch[action].length > 0) {
      const batch = actionsBatch[action];
      batch.forEach(item => {
        ACTION_MAP[action](item.data, item.rowNumber);
      });
      console.log(`Procesados ${batch.length} ${action}`);
    }
  });

  const summary = `Proceso completado:\n\n` +
    `- Movidos a Recursos devueltos: ${actionsBatch["Ítem devuelto/encontrado"].length} // ` +
    `- Movidos a Seguimiento: ${actionsBatch["Dar seguimiento al ítem"].length} // ` +
    `- Correos enviados: ${actionsBatch["✉️ Primer recordatorio"].length +
    actionsBatch["✉️ Segundo recordatorio"].length +
    actionsBatch["✉️ Aviso de recarga"].length +
    actionsBatch["✉️ Confirmación de la recarga"].length
    }`;

  SpreadsheetApp.getActiveSpreadsheet().toast(
    summary,
    "Resumen de ejecución ✅",
    15
  );
};

const hasScript = () => {
  UI.alert("Información ⚠️", "Script: SP | Reporte de deudores", UI.ButtonSet.OK);
}

// **********************************************
// MENÚ
// **********************************************

/**
 * Crea el menú personalizado
 */
const onOpen = () => {
  UI.createMenu('Scripts 🟢')
    .addItem('➡️ Procesar datos de: ' + SHEETS.alma.getSheetName(), 'startProcess')
    .addItem('🧪 Ejecutar acciones (L) de: ' + SHEETS.overdueItems.getSheetName(), 'executeActions')
    .addSeparator()
    .addItem('🗑️ Borrar datos de: ' + SHEETS.alma.getSheetName(), 'deleteData') // ✅
    .addSeparator()
    .addItem('⚠️ Información del script', 'hasScript') // ✅
    .addToUi();
};