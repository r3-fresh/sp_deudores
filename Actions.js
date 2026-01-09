// ========================================
// FUNCIONES DE ACCIONES POR LOTES
// ========================================

/**
 * Mueve múltiples registros a "Recursos devueltos / Histórico"
 * @param {Array<Array>} rowsWithNumbers - Array de filas con números de fila al final
 * @returns {boolean} true si tuvo éxito
 */
const moveToReturnedItems = (rowsWithNumbers) => {
    try {
        if (
            !validateSheet(SHEETS.overdueItems, "Préstamos vencidos") ||
            !validateSheet(SHEETS.returnedItems, "Recursos devueltos")
        ) {
            return false;
        }

        const rowsData = rowsWithNumbers.map((row) => row.slice(0, -1));
        const rowNumbers = rowsWithNumbers.map((row) => row[row.length - 1]);

        const valuesToCopy = rowsData.map((row, index) => {
            const baseData = row.slice(0, COLUMNS.ACTION);
            const rowNumber = rowNumbers[index];

            const logInfo = SHEETS.overdueItems
                .getRange(rowNumber, COLUMNS.LOG + 1)
                .getValue();

            const actionMessage = logInfo
                ? `${logInfo}\n${new Date().toLocaleString()}: Ítem devuelto por ejecución de acciones`
                : `${new Date().toLocaleString()}: Ítem devuelto por ejecución de acciones`;

            return [
                ...baseData,
                new Date(),
                actionMessage,
            ];
        });

        const lastRow = SHEETS.returnedItems.getLastRow();
        SHEETS.returnedItems
            .getRange(lastRow + 1, 1, valuesToCopy.length, valuesToCopy[0].length)
            .setValues(valuesToCopy);

        rowNumbers
            .sort((a, b) => b - a)
            .forEach((rowNum) => {
                SHEETS.overdueItems.deleteRow(rowNum);
            });

        return true;
    } catch (error) {
        console.error("❌ Error en moveToReturnedItems:", error);
        showToast(`Error moviendo registros: ${error.message}`, "Error", 5, "❌");
        return false;
    }
};

/**
 * Mueve múltiples registros a "Seguimiento de préstamos"
 * NO elimina las filas originales, solo limpia la acción
 * 
 * @param {Array<Array>} rowsWithNumbers - Array de filas con números
 * @returns {boolean} true si tuvo éxito
 */
const moveToTrackingItems = (rowsWithNumbers) => {
    try {
        if (
            !validateSheet(SHEETS.overdueItems, "Préstamos vencidos") ||
            !validateSheet(SHEETS.trackingItems, "Seguimiento de préstamos")
        ) {
            return false;
        }

        const rowsData = rowsWithNumbers.map((row) => row.slice(0, -1));
        const rowNumbers = rowsWithNumbers.map((row) => row[row.length - 1]);

        const valuesToCopy = rowsData.map((row, index) => {
            const baseData = row.slice(0, COLUMNS.ACTION);
            const rowNumber = rowNumbers[index];

            const logInfo = SHEETS.overdueItems
                .getRange(rowNumber, COLUMNS.LOG + 1)
                .getValue();

            const actionMessage = logInfo
                ? `${logInfo}\n${new Date().toLocaleString()}: Ítem movido a Seguimiento`
                : `${new Date().toLocaleString()}: Ítem movido a Seguimiento`;

            SHEETS.overdueItems
                .getRange(rowNumber, COLUMNS.ACTION + 1)
                .clearContent();

            return [
                ...baseData,
                new Date(),
                actionMessage,
            ];
        });

        const lastRow = SHEETS.trackingItems.getLastRow();
        SHEETS.trackingItems
            .getRange(lastRow + 1, 1, valuesToCopy.length, valuesToCopy[0].length)
            .setValues(valuesToCopy);

        return true;
    } catch (error) {
        console.error("❌ Error en moveToTrackingItems:", error);
        showToast(
            `Error moviendo a seguimiento: ${error.message}`,
            "Error",
            5,
            "❌"
        );
        return false;
    }
};

/**
 * Ejecuta todas las acciones pendientes en la hoja "Préstamos vencidos"
 * 
 * FLUJO:
 * 1. Leer hoja y clasificar por tipo de acción
 * 2. Procesar movimientos en batch
 * 3. Procesar correos individuales
 * 4. Mostrar resumen de acciones ejecutadas
 */
const executeActions = () => {
    if (!validateSheet(SHEETS.overdueItems, "Préstamos vencidos / Deudores")) {
        return;
    }

    try {
        const fullData = SHEETS.overdueItems.getDataRange().getValues();
        const headers = fullData[0];
        const data = fullData.slice(1);

        const ACTION_MAP = {
            [ACTIONS.FIRST_REMINDER]: sendFirstReminder,
            [ACTIONS.SECOND_REMINDER]: sendSecondReminder,
            [ACTIONS.RECHARGE_NOTICE]: sendRechargeNotice,
            [ACTIONS.RECHARGE_CONFIRMATION]: sendRechargeConfirmation,
            [ACTIONS.MOVE_TO_RETURNED]: moveToReturnedItems,
            [ACTIONS.MOVE_TO_TRACKING]: moveToTrackingItems,
        };

        const actionsBatch = {
            [ACTIONS.FIRST_REMINDER]: [],
            [ACTIONS.SECOND_REMINDER]: [],
            [ACTIONS.RECHARGE_NOTICE]: [],
            [ACTIONS.RECHARGE_CONFIRMATION]: [],
            [ACTIONS.MOVE_TO_RETURNED]: [],
            [ACTIONS.MOVE_TO_TRACKING]: [],
        };

        // Clasificar cada fila según su acción
        data.forEach((row, index) => {
            const rowNumber = index + 2;
            const actionValue = row[COLUMNS.ACTION];

            if (actionValue && ACTION_MAP[actionValue]) {
                actionsBatch[actionValue].push({
                    data: row,
                    rowNumber: rowNumber,
                });
            }
        });

        const totalActions = Object.values(actionsBatch).reduce(
            (sum, batch) => sum + batch.length,
            0
        );

        if (totalActions === 0) {
            showToast(
                "No hay acciones pendientes para ejecutar",
                "Información",
                5,
                "ℹ️"
            );
            return;
        }

        // Procesar movimientos a Recursos devueltos (batch)
        if (actionsBatch[ACTIONS.MOVE_TO_RETURNED].length > 0) {
            const batch = actionsBatch[ACTIONS.MOVE_TO_RETURNED];
            const rowsToProcess = batch.map((item) => [...item.data, item.rowNumber]);
            moveToReturnedItems(rowsToProcess);
        }

        // Procesar movimientos a Seguimiento (batch)
        if (actionsBatch[ACTIONS.MOVE_TO_TRACKING].length > 0) {
            const batch = actionsBatch[ACTIONS.MOVE_TO_TRACKING];
            const rowsToProcess = batch.map((item) => [...item.data, item.rowNumber]);
            moveToTrackingItems(rowsToProcess);
        }

        // Procesar envíos de correo (individual)
        const emailActions = [
            ACTIONS.FIRST_REMINDER,
            ACTIONS.SECOND_REMINDER,
            ACTIONS.RECHARGE_NOTICE,
            ACTIONS.RECHARGE_CONFIRMATION,
        ];

        emailActions.forEach((action) => {
            if (actionsBatch[action].length > 0) {
                const batch = actionsBatch[action];
                batch.forEach((item) => {
                    try {
                        ACTION_MAP[action](item.data, item.rowNumber);
                    } catch (error) {
                        console.error(`❌ Error procesando fila ${item.rowNumber}:`, error);
                    }
                });
            }
        });

        const emailCount =
            actionsBatch[ACTIONS.FIRST_REMINDER].length +
            actionsBatch[ACTIONS.SECOND_REMINDER].length +
            actionsBatch[ACTIONS.RECHARGE_NOTICE].length +
            actionsBatch[ACTIONS.RECHARGE_CONFIRMATION].length;

        const summary = [
            `Ítems devueltos: ${actionsBatch[ACTIONS.MOVE_TO_RETURNED].length}`,
            `Ítems en seguimiento: ${actionsBatch[ACTIONS.MOVE_TO_TRACKING].length}`,
            `Correos enviados: ${emailCount}`,
        ].join(" // ");

        showToast(summary, "Acciones ejecutadas", 15, "✅");
    } catch (error) {
        console.error("❌ Error en executeActions:", error);
        console.error("Stack:", error.stack);

        showToast(`Error ejecutando acciones: ${error.message}`, "Error", 8, "❌");
    }
};
