// ========================================
// PROCESAMIENTO DE DATOS
// ========================================

/**
 * Limpia la hoja de Alma preparándola para nuevos datos
 */
const deleteData = () => {
    if (!validateSheet(SHEETS.alma, "Reporte de deudores - Widget")) {
        return;
    }

    const lastRow = SHEETS.alma.getLastRow();

    if (lastRow < 2) {
        showToast("La hoja ya se encuentra vacía", "Información", 5, "ℹ️");
        return;
    }

    SHEETS.alma.getRange(`A2:L${lastRow}`).clearContent();
    showToast(`Se limpiaron ${lastRow - 1} filas`, "Limpieza exitosa", 5, "✅");
};

/**
 * Procesa datos de Alma e identifica nuevos deudores y devoluciones
 * 
 * FLUJO:
 * 1. Validación de hojas requeridas
 * 2. Carga en memoria con índices Set para búsqueda O(1)
 * 3. Identificación de nuevos deudores
 * 4. Identificación de recursos devueltos
 * 5. Escritura batch de cambios
 * 6. Reporte de resultados
 */
const startProcess = () => {
    // Validación de hojas requeridas
    const requiredSheets = [
        { sheet: SHEETS.alma, name: "Reporte de deudores - Widget" },
        { sheet: SHEETS.overdueItems, name: "Préstamos vencidos / Deudores" },
        { sheet: SHEETS.returnedItems, name: "Recursos devueltos / Histórico" },
    ];

    const missingSheets = requiredSheets
        .filter((s) => !s.sheet)
        .map((s) => s.name);

    if (missingSheets.length > 0) {
        showToast(
            `Hojas faltantes:\n- ${missingSheets.join("\n- ")}`,
            "Error de configuración",
            8,
            "❌"
        );
        return;
    }

    if (SHEETS.alma.getRange("A2").getValue() === "") {
        showToast("No hay datos para procesar", "Error", 5, "❌");
        return;
    }

    try {
        // Carga en memoria
        const almaFullData = SHEETS.alma.getDataRange().getValues();
        const almaHeaders = almaFullData[0];
        const almaData = almaFullData.slice(1);

        const overdueFullData = SHEETS.overdueItems.getDataRange().getValues();
        const overdueHeaders = overdueFullData[0];
        const overdueData = overdueFullData.slice(1);

        // Crear índices para búsqueda rápida O(1)
        const overdueIndex = new Set(
            overdueData.map((row) => generateRecordKey(row))
        );

        // Identificar nuevos deudores
        const newDebtors = [];
        const updates = [];

        almaData.forEach((row, index) => {
            const recordKey = generateRecordKey(row);
            const isRegistered = overdueIndex.has(recordKey);

            updates.push({
                row: index + 2,
                value: isRegistered ? STATUS.REGISTERED : STATUS.NEW,
            });

            if (!isRegistered) {
                // Copiar las primeras 11 columnas (Campus hasta Fecha de Vencimiento)
                // Las columnas 11-16 se dejan vacías para overdueItems
                newDebtors.push(row.slice(0, 11));
            }
        });

        // Identificar devoluciones
        const almaIndex = new Set(almaData.map((row) => generateRecordKey(row)));
        const returnedItems = [];
        const rowsToDelete = [];

        overdueData.forEach((row, index) => {
            const recordKey = generateRecordKey(row);

            if (!almaIndex.has(recordKey)) {
                const currentLog = row[COLUMNS.LOG] || "";
                const timestamp = new Date().toLocaleString("es-PE", {
                    timeZone: "America/Lima",
                    year: "numeric",
                    month: "2-digit",
                    day: "2-digit",
                    hour: "2-digit",
                    minute: "2-digit",
                });

                const actionMessage = currentLog
                    ? `${currentLog}\n${timestamp}: Devuelto por el usuario`
                    : `${timestamp}: Devuelto por el usuario`;

                // Copiar todas las columnas de overdueItems (17 columnas)
                // Luego agregar columnas específicas de returnedItems
                const returnedRow = [
                    ...row.slice(0, 11),      // Campus hasta Fecha de Vencimiento
                    new Date(),                // Fecha de devolución
                    actionMessage,             // Bitácora actualizada
                    row[COLUMNS.RECHARGE_DATE] || "",    // Fecha de recargo
                    row[COLUMNS.WITHDRAWAL_DATE] || "",  // Fecha de retiro
                    row[COLUMNS.COST] || "",             // Costo
                    row[COLUMNS.OBSERVATIONS] || "",     // Observaciones
                    "",                        // Estado (vacío inicialmente)
                    "",                        // Consulta de pago a caja
                    "",                        // ¿Realizó el pago?
                ];

                returnedItems.push(returnedRow);
                rowsToDelete.push(index + 2);
            }
        });

        // Escritura batch
        if (updates.length > 0) {
            const sortedUpdates = updates.sort((a, b) => a.row - b.row);
            const firstRow = sortedUpdates[0].row;
            const lastRow = sortedUpdates[sortedUpdates.length - 1].row;
            const rowCount = lastRow - firstRow + 1;

            const outputValues = new Array(rowCount).fill([""]);
            sortedUpdates.forEach((update) => {
                outputValues[update.row - firstRow] = [update.value];
            });

            // Escribir todos los estados en columna "Estado" (índice 11)
            SHEETS.alma
                .getRange(firstRow, COLUMNS.ALMA_STATUS + 1, rowCount, 1)
                .setValues(outputValues);
        }

        if (newDebtors.length > 0) {
            const lastRow = SHEETS.overdueItems.getLastRow();
            SHEETS.overdueItems
                .getRange(lastRow + 1, 1, newDebtors.length, newDebtors[0].length)
                .setValues(newDebtors);
        }

        if (returnedItems.length > 0) {
            const lastRow = SHEETS.returnedItems.getLastRow();
            SHEETS.returnedItems
                .getRange(lastRow + 1, 1, returnedItems.length, returnedItems[0].length)
                .setValues(returnedItems);

            rowsToDelete
                .sort((a, b) => b - a)
                .forEach((row) => {
                    SHEETS.overdueItems.deleteRow(row);
                });
        }

        // Reporte final
        const registeredCount = updates.filter(
            (u) => u.value === STATUS.REGISTERED
        ).length;

        const summary = [
            `Registros previos: ${registeredCount}`,
            `Nuevos deudores: ${newDebtors.length}`,
            `Ítems devueltos: ${returnedItems.length}`,
        ].join(" // ");

        showToast(summary, "Proceso completado", 15, "✅");
    } catch (error) {
        console.error("❌ Error en startProcess:", error);
        console.error("Stack:", error.stack);

        showToast(
            `Error inesperado: ${error.message}`,
            "Error en proceso",
            8,
            "❌"
        );
    }
};
