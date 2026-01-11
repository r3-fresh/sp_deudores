// ========================================
// INTERFAZ DE USUARIO
// ========================================

/**
 * Muestra información sobre el script actual
 */
const hasScript = () => {
    const info = `
📄 Script: SP | Reporte de deudores`.trim();

    UI.alert("Información del Script ℹ️", info, UI.ButtonSet.OK);
};

/**
 * Crea el menú personalizado en la interfaz de Google Sheets
 * Se ejecuta automáticamente al abrir el documento
 */
const onOpen = () => {
    const email = Session.getActiveUser().getEmail();
    if (email == "bibliotecariovirtual@continental.edu.pe") {
        try {
            UI.createMenu("Scripts 🟢")
                .addItem("➡️ Procesar datos de: " + SHEETS.alma.getName(), "startProcess")
                .addItem(
                    "🧪 Ejecutar acciones (L) de: " + SHEETS.overdueItems.getName(),
                    "executeActions"
                )
                .addSeparator()
                .addItem("🗑️ Borrar datos de: " + SHEETS.alma.getName(), "deleteData")
                .addSeparator()
                .addItem("⚠️ Información del script", "hasScript")
                .addToUi();
        } catch (error) {
            console.error("❌ Error creando menú:", error);
        }
    }
};
