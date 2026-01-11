// ========================================
// ENVÍO DE CORREOS ELECTRÓNICOS
// ========================================

/**
 * Envía un correo electrónico usando plantilla HTML
 * @param {string} to - Dirección de correo del destinatario
 * @param {string} subject - Asunto del correo
 * @param {string} htmlBody - Cuerpo del correo en HTML
 * @returns {boolean} true si el envío fue exitoso
 */
const sendEmail = (to, subject, htmlBody) => {
  try {
    if (!to || to.trim() === "") {
      console.error("Email destinatario vacío");
      return false;
    }

    GmailApp.sendEmail(to, subject, "", {
      htmlBody: htmlBody,
      name: "Hub de Información - UC Continental",
    });

    console.log(`✅ Email enviado a ${to}: ${subject}`);
    return true;
  } catch (error) {
    console.error(`❌ Error enviando email a ${to}:`, error);
    return false;
  }
};

/**
 * Obtiene el nombre del mes en español
 * @param {Date} date - Fecha
 * @returns {string} Nombre del mes
 */
const getMonthName = (date) => {
  const months = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
  ];
  return months[date.getMonth()];
};

/**
 * Formatea una fecha en formato dd/mm/yyyy
 * @param {Date|string} date - Fecha a formatear
 * @returns {string} Fecha formateada
 */
const formatDate = (date) => {
  if (!date) return "";
  const d = typeof date === 'string' ? new Date(date) : date;
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
};

/**
 * Formatea lista de libros como HTML
 * @param {string} bookTitle - Título del libro
 * @returns {string} HTML con el libro
 */
const formatBookList = (bookTitle) => {
  return `<li>${bookTitle}</li>`;
};

/**
 * Envía primer recordatorio al deudor
 * @param {Array} data - Datos del registro
 * @param {number} rowNumber - Número de fila
 */
const sendFirstReminder = (data, rowNumber) => {
  const email = data[COLUMNS.EMAIL];
  const nombre = data[COLUMNS.FULL_NAME];
  const titulo = data[COLUMNS.TITLE];
  const fechaVencimiento = formatDate(data[COLUMNS.DUE_DATE]);
  const mes = getMonthName(new Date(data[COLUMNS.DUE_DATE]));

  // Cargar plantilla HTML
  const template = HtmlService.createTemplateFromFile('templates/emailFirstReminder');
  template.NOMBRE = nombre;
  template.MES = mes;
  template.FECHA_VENCIMIENTO = fechaVencimiento;
  template.LIBROS = formatBookList(titulo);
  template.URL_IMAGEN_BUZON = "https://hubinformacion.continental.edu.pe/web/wp-content/uploads/2026/01/buzones-hyo.png";

  const subject = "Hub Huancayo | 🚨 ¡Atención! Tienes un libro pendiente para devolver ⚠️ 1er recordatorio";
  const htmlBody = template.evaluate().getContent();

  if (sendEmail(email, subject, htmlBody)) {
    const currentLog = SHEETS.overdueItems
      .getRange(rowNumber, COLUMNS.LOG + 1)
      .getValue();
    updateActionLog(rowNumber, "✉️ Primer recordatorio enviado", currentLog);
  } else {
    showToast(
      `No se pudo enviar correo a ${nombre}`,
      "Error de envío",
      5,
      "❌"
    );
  }
};

/**
 * Envía segundo recordatorio al deudor (tono más urgente)
 * @param {Array} data - Datos del registro
 * @param {number} rowNumber - Número de fila
 */
const sendSecondReminder = (data, rowNumber) => {
  const email = data[COLUMNS.EMAIL];
  const nombre = data[COLUMNS.FULL_NAME];
  const titulo = data[COLUMNS.TITLE];
  const fechaVencimiento = formatDate(data[COLUMNS.DUE_DATE]);

  // Cargar plantilla HTML
  const template = HtmlService.createTemplateFromFile('templates/emailSecondReminder');
  template.NOMBRE = nombre;
  template.FECHA_VENCIMIENTO = fechaVencimiento;
  template.LIBROS = formatBookList(titulo);
  template.URL_IMAGEN_BUZON = ""; // Usuario agregará el enlace

  const subject = "Hub Huancayo | 🚨 ¡Atención! Aún tienes un libro pendiente por devolver ⚠️ 2do recordatorio";
  const htmlBody = template.evaluate().getContent();

  if (sendEmail(email, subject, htmlBody)) {
    const currentLog = SHEETS.overdueItems
      .getRange(rowNumber, COLUMNS.LOG + 1)
      .getValue();
    updateActionLog(rowNumber, "⚠️ Segundo recordatorio enviado", currentLog);
  }
};

/**
 * Envía aviso de recarga (penalización)
 * @param {Array} data - Datos del registro
 * @param {number} rowNumber - Número de fila
 */
const sendRechargeNotice = (data, rowNumber) => {
  const email = data[COLUMNS.EMAIL];
  const nombre = data[COLUMNS.FULL_NAME];
  const titulo = data[COLUMNS.TITLE];
  const fechaVencimiento = formatDate(data[COLUMNS.DUE_DATE]);
  const costo = data[COLUMNS.COST] || "S/ 0.00"; // Obtener costo o valor por defecto

  // Calcular fecha límite (por ejemplo, 3 días después de hoy)
  const fechaLimite = new Date();
  fechaLimite.setDate(fechaLimite.getDate() + 3);

  // Cargar plantilla HTML
  const template = HtmlService.createTemplateFromFile('templates/emailRechargeNotice');
  template.NOMBRE = nombre;
  template.FECHA_VENCIMIENTO = fechaVencimiento;
  template.FECHA_LIMITE = formatDate(fechaLimite);
  template.LIBROS = formatBookList(titulo);
  template.MONTO = costo;

  const subject = "Hub Huancayo | 🚨 Aviso de recarga por devolución pendiente de libro";
  const htmlBody = template.evaluate().getContent();

  if (sendEmail(email, subject, htmlBody)) {
    const currentLog = SHEETS.overdueItems
      .getRange(rowNumber, COLUMNS.LOG + 1)
      .getValue();
    updateActionLog(rowNumber, "💳 Aviso de recarga enviado", currentLog);
  }
};

/**
 * Envía confirmación de pago de recarga
 * @param {Array} data - Datos del registro
 * @param {number} rowNumber - Número de fila
 */
const sendRechargeConfirmation = (data, rowNumber) => {
  const email = data[COLUMNS.EMAIL];
  const nombre = data[COLUMNS.FULL_NAME];
  const titulo = data[COLUMNS.TITLE];
  const costo = data[COLUMNS.COST] || "S/ 0.00";

  // Cargar plantilla HTML
  const template = HtmlService.createTemplateFromFile('templates/emailRechargeConfirmation');
  template.NOMBRE = nombre;
  template.LIBROS = formatBookList(titulo);
  template.MONTO = costo;
  template.URL_IMAGEN_BUZON = "https://hubinformacion.continental.edu.pe/web/wp-content/uploads/2026/01/buzones-hyo.png";

  const subject = "Hub Huancayo | 🚨 Confirmación de recargo por devolución pendiente";
  const htmlBody = template.evaluate().getContent();

  if (sendEmail(email, subject, htmlBody)) {
    const currentLog = SHEETS.overdueItems
      .getRange(rowNumber, COLUMNS.LOG + 1)
      .getValue();
    updateActionLog(rowNumber, "✅ Confirmación de recarga enviada", currentLog);
  }
};
