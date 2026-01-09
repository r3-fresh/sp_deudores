/**
 * ========================================
 * SISTEMA DE GESTIÓN DE DEUDORES
 * ========================================
 *
 * Archivo principal - Punto de entrada del sistema
 *
 * Este sistema automatiza el proceso de gestión de deudores
 * utilizando Google Apps Script con integración a Google Sheets.
 *
 * ESTRUCTURA MODULAR:
 * - Config.js: Constantes y configuración
 * - Utils.js: Funciones auxiliares
 * - DataProcessor.js: Procesamiento de datos de Alma
 * - Emails.js: Envío de recordatorios por correo
 * - Actions.js: Ejecución de acciones por lotes
 * - UI.js: Interfaz de usuario y menús
 *
 * FLUJO PRINCIPAL:
 * 1. Importar datos desde Alma
 * 2. Procesar nuevos deudores y devoluciones (startProcess)
 * 3. Definir acciones en columna L
 * 4. Ejecutar acciones (executeActions)
 *
 * @author Fredy Romero <romeroespinoza.fp@gmail.com>
 * @version 2.0.0 - Refactorizado para mejor mantenibilidad
 * @license MIT
 */

/**
 * NOTAS IMPORTANTES:
 *
 * 1. Todos los archivos .gs comparten el mismo namespace global en Apps Script
 * 2. Las funciones exportadas deben ser llamadas directamente desde los menús
 * 3. Los límites de cuota de GmailApp aplican (500 correos/día para cuentas gratuitas)
 * 4. El sistema usa procesamiento batch para optimizar operaciones de Sheets
 *
 * HOJAS REQUERIDAS:
 * - Reporte de deudores - Widget (ID: 563966915)
 * - Préstamos vencidos / Deudores (ID: 491373272)
 * - Seguimiento de préstamos (ID: 687630222)
 * - Recursos devueltos / Histórico (ID: 1634827826)
 */

// Este archivo sirve como punto de entrada y documentación principal
// Todas las funciones están definidas en sus respectivos módulos
