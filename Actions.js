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

        const timestamp = new Date().toLocaleString("es-PE", {
            timeZone: "America/Lima",
            year: "numeric",
            month: "2-digit",
            day: "2-digit",
            hour: "2-digit",
            minute: "2-digit",
        });

        const valuesToCopy = rowsData.map((row, index) => {
            const rowNumber = rowNumbers[index];

            const logInfo = SHEETS.overdueItems
                .getRange(rowNumber, COLUMNS.LOG + 1)
                .getValue();

            const actionMessage = logInfo
                ? `${logInfo}\n${timestamp}: Ítem devuelto por ejecución de acciones`
                : `${timestamp}: Ítem devuelto por ejecución de acciones`;

            // Estructura para returnedItems (20 columnas):
            // 0-10: Campus hasta Fecha de Vencimiento
            // 11: Fecha de devolución
            // 12: Bitácora de acciones
            // 13-16: Fecha recargo, Fecha retiro, Costo, Observaciones
            // 17-19: Estado, Consulta pago caja, ¿Realizó pago?

            return [
                ...row.slice(0, 11),                    // Campus hasta Fecha de Vencimiento
                new Date(),                             // Fecha de devolución
                actionMessage,                          // Bitácora actualizada
                row[COLUMNS.RECHARGE_DATE] || "",      // Fecha de recargo
                row[COLUMNS.WITHDRAWAL_DATE] || "",    // Fecha de retiro
                row[COLUMNS.COST] || "",               // Costo
                row[COLUMNS.OBSERVATIONS] || "",       // Observaciones
                "",                                     // Estado
                "",                                     // Consulta de pago a caja
                "",                                     // ¿Realizó el pago?
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

        const timestamp = new Date().toLocaleString("es-PE", {
            timeZone: "America/Lima",
            year: "numeric",
            month: "2-digit",
            day: "2-digit",
            hour: "2-digit",
            minute: "2-digit",
        });

        const valuesToCopy = rowsData.map((row, index) => {
            const rowNumber = rowNumbers[index];

            const logInfo = SHEETS.overdueItems
                .getRange(rowNumber, COLUMNS.LOG + 1)
                .getValue();

            const actionMessage = logInfo
                ? `${logInfo}\n${timestamp}: Ítem movido a Seguimiento`
                : `${timestamp}: Ítem movido a Seguimiento`;

            // Estructura para trackingItems (20 columnas):
            // 0-10: Campus hasta Fecha de Vencimiento
            // 11: Fecha de seguimiento
            // 12: Bitácora de acciones
            // 13-16: Fecha recargo, Fecha retiro, Costo, Observaciones
            // 17-19: Estado, Consulta pago caja, ¿Realizó pago?

            return [
                ...row.slice(0, 11),                    // Campus hasta Fecha de Vencimiento
                new Date(),                             // Fecha de seguimiento
                actionMessage,                          // Bitácora actualizada
                row[COLUMNS.RECHARGE_DATE] || "",      // Fecha de recargo
                row[COLUMNS.WITHDRAWAL_DATE] || "",    // Fecha de retiro
                row[COLUMNS.COST] || "",               // Costo
                row[COLUMNS.OBSERVATIONS] || "",       // Observaciones
                "",                                     // Estado
                "",                                     // Consulta de pago a caja
                "",                                     // ¿Realizó el pago?
            ];
        });

        // Insertar en seguimiento (1 operación)
        const lastRow = SHEETS.trackingItems.getLastRow();
        SHEETS.trackingItems
            .getRange(lastRow + 1, 1, valuesToCopy.length, valuesToCopy[0].length)
            .setValues(valuesToCopy);

        // Limpiar acciones de las filas en overdueItems
        // IMPORTANTE: Hacer esto DESPUÉS de copiar los datos
        rowNumbers.forEach((rowNum) => {
            SHEETS.overdueItems.getRange(rowNum, COLUMNS.ACTION + 1).clearContent();
        });

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
 * Agrupa items por ID Usuario para envío consolidado de correos
 * @param {Array} batch - Array de items con misma acción
 * @returns {Map} Map de userId -> array de items
 */
const groupByUserId = (batch) => {
    const grouped = new Map();

    batch.forEach(item => {
        const userId = item.data[COLUMNS.USER_ID];
        if (!grouped.has(userId)) {
            grouped.set(userId, []);
        }
        grouped.get(userId).push(item);
    });

    return grouped;
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

        // ===================================
        // ORDEN CRÍTICO DE EJECUCIÓN
        // ===================================
        // 1. Correos (no modifican filas)
        // 2. Movimientos a Seguimiento (no eliminan filas)
        // 3. Movimientos a Devueltos (ELIMINA filas - debe ser ÚLTIMO)
        //
        // Razón: Cuando se eliminan filas, los índices de fila cambian.
        // Si procesamos devoluciones primero, las referencias de fila
        // de las demás acciones quedarán obsoletas.

        // PASO 1: Procesar envíos de correo agrupados por (usuario + acción)
        const emailActions = [
            ACTIONS.FIRST_REMINDER,
            ACTIONS.SECOND_REMINDER,
            ACTIONS.RECHARGE_NOTICE,
            ACTIONS.RECHARGE_CONFIRMATION,
        ];

        emailActions.forEach((action) => {
            if (actionsBatch[action].length > 0) {
                const batch = actionsBatch[action];
                const groupedByUser = groupByUserId(batch); // Ya filtrado por acción

                // Procesar cada usuario (enviar un solo correo por usuario)
                groupedByUser.forEach((userItems, userId) => {
                    try {
                        const dataItems = userItems.map(item => item.data);
                        const rowNumbers = userItems.map(item => item.rowNumber);

                        // Llamar función de envío con arrays
                        ACTION_MAP[action](dataItems, rowNumbers);
                    } catch (error) {
                        console.error(`❌ Error procesando usuario ${userId}:`, error);
                    }
                });
            }
        });

        // PASO 2: Procesar movimientos a Seguimiento (no elimina filas)
        if (actionsBatch[ACTIONS.MOVE_TO_TRACKING].length > 0) {
            const batch = actionsBatch[ACTIONS.MOVE_TO_TRACKING];
            const rowsToProcess = batch.map((item) => [...item.data, item.rowNumber]);
            moveToTrackingItems(rowsToProcess);
        }

        // PASO 3: Procesar movimientos a Recursos devueltos (ELIMINA filas - debe ser último)
        if (actionsBatch[ACTIONS.MOVE_TO_RETURNED].length > 0) {
            const batch = actionsBatch[ACTIONS.MOVE_TO_RETURNED];
            const rowsToProcess = batch.map((item) => [...item.data, item.rowNumber]);
            moveToReturnedItems(rowsToProcess);
        }

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
