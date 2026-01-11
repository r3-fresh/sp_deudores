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

    return true;
  } catch (error) {
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

  let d;

  // Si ya es un objeto Date válido
  if (date instanceof Date) {
    d = date;
  }
  // Si es un string en formato dd/mm/yyyy
  else if (typeof date === 'string' && date.includes('/')) {
    const parts = date.split('/');
    if (parts.length === 3) {
      // Formato dd/mm/yyyy -> convertir a Date(year, month-1, day)
      d = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
    } else {
      d = new Date(date);
    }
  }
  // Cualquier otro tipo (número, string sin /, etc.)
  else {
    d = new Date(date);
  }

  // Validar que la fecha sea válida
  if (isNaN(d.getTime())) {
    console.error("Fecha inválida:", date);
    return "";
  }

  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
};

/**
 * Formatea lista de libros como HTML
 * @param {Array<string>|string} books - Título(s) del/los libro(s)
 * @returns {string} HTML con los libros
 */
const formatBookList = (books) => {
  if (typeof books === 'string') {
    return `<li>${books}</li>`;
  }
  // Si es un array, generar múltiples <li>
  return books.map(book => `<li>${book}</li>`).join('\n');
};

/**
 * Envía primer recordatorio al deudor
 * @param {Array<Array>} dataItems - Array de registros del mismo usuario
 * @param {Array<number>} rowNumbers - Array de números de fila
 */
const sendFirstReminder = (dataItems, rowNumbers) => {
  const firstItem = dataItems[0];
  const email = firstItem[COLUMNS.EMAIL];
  const nombre = firstItem[COLUMNS.FULL_NAME];

  // Combinar títulos de todos los libros
  const titulos = dataItems.map(item => item[COLUMNS.TITLE]);

  // Fecha de vencimiento = 1 día después de hoy
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const fechaVencimiento = formatDate(tomorrow);
  const mes = getMonthName(tomorrow);

  // Cargar y procesar plantilla HTML
  const template = HtmlService.createTemplateFromFile('templates/emailFirstReminder');
  template.NOMBRE = nombre;
  template.CAMPUS = EMAIL_CONFIG.CAMPUS_NAME;
  template.MES = mes;
  template.FECHA_VENCIMIENTO = fechaVencimiento;
  template.LIBROS = formatBookList(titulos);
  template.URL_IMAGEN_BUZON = EMAIL_CONFIG.MAILBOX_IMAGE_URL;

  const subject = `Hub ${EMAIL_CONFIG.CAMPUS_NAME} | ⚠️ ¡Atención! Tienes un libro pendiente para devolver ⚠️ 1er recordatorio`;

  const htmlBody = template.evaluate().getContent();

  if (sendEmail(email, subject, htmlBody)) {
    // Actualizar log de TODOS los registros
    rowNumbers.forEach((rowNumber) => {
      const currentLog = SHEETS.overdueItems.getRange(rowNumber, COLUMNS.LOG + 1).getValue();
      updateActionLog(rowNumber, "✉️ Primer recordatorio enviado", currentLog);
    });
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
 * @param {Array<Array>} dataItems - Array de registros del mismo usuario
 * @param {Array<number>} rowNumbers - Array de números de fila
 */
const sendSecondReminder = (dataItems, rowNumbers) => {
  const firstItem = dataItems[0];
  const email = firstItem[COLUMNS.EMAIL];
  const nombre = firstItem[COLUMNS.FULL_NAME];

  // Combinar títulos de todos los libros
  const titulos = dataItems.map(item => item[COLUMNS.TITLE]);

  // Fecha de vencimiento = 1 día después de hoy
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const fechaVencimiento = formatDate(tomorrow);

  // Cargar y procesar plantilla HTML
  const template = HtmlService.createTemplateFromFile('templates/emailSecondReminder');
  template.NOMBRE = nombre;
  template.CAMPUS = EMAIL_CONFIG.CAMPUS_NAME;
  template.FECHA_VENCIMIENTO = fechaVencimiento;
  template.LIBROS = formatBookList(titulos);
  template.URL_IMAGEN_BUZON = EMAIL_CONFIG.MAILBOX_IMAGE_URL;

  const subject = `Hub ${EMAIL_CONFIG.CAMPUS_NAME} | ⚠️ ¡Atención! Aún tienes un libro pendiente por devolver ⚠️ 2do recordatorio`;

  const htmlBody = template.evaluate().getContent();

  if (sendEmail(email, subject, htmlBody)) {
    // Actualizar log de TODOS los registros
    rowNumbers.forEach((rowNumber) => {
      const currentLog = SHEETS.overdueItems.getRange(rowNumber, COLUMNS.LOG + 1).getValue();
      updateActionLog(rowNumber, "⚠️ Segundo recordatorio enviado", currentLog);
    });
  }
};

/**
 * Envía aviso de recarga (penalización)
 * @param {Array<Array>} dataItems - Array de registros del mismo usuario
 * @param {Array<number>} rowNumbers - Array de números de fila
 */
const sendRechargeNotice = (dataItems, rowNumbers) => {
  const firstItem = dataItems[0];
  const email = firstItem[COLUMNS.EMAIL];
  const nombre = firstItem[COLUMNS.FULL_NAME];

  // Combinar títulos de todos los libros
  const titulos = dataItems.map(item => item[COLUMNS.TITLE]);

  // Usar fecha de vencimiento del primer registro
  const fechaVencimiento = formatDate(firstItem[COLUMNS.DUE_DATE]);

  // Usar fecha límite del primer registro
  const fechaLimiteValue = firstItem[COLUMNS.RECHARGE_DATE] || "";
  const fechaLimite = formatDate(fechaLimiteValue);

  // Sumar costos (parsear texto a número)
  const costos = dataItems.map(item => {
    const costoStr = item[COLUMNS.COST] || "0.00";
    return parseFloat(costoStr.replace(/[^\d.]/g, '')) || 0;
  });
  const totalCosto = costos.reduce((sum, c) => sum + c, 0);

  // Cargar y procesar plantilla HTML
  const template = HtmlService.createTemplateFromFile('templates/emailRechargeNotice');
  template.NOMBRE = nombre;
  template.CAMPUS = EMAIL_CONFIG.CAMPUS_NAME;
  template.FECHA_VENCIMIENTO = fechaVencimiento;
  template.FECHA_LIMITE = fechaLimite;
  template.LIBROS = formatBookList(titulos);
  template.MONTO = `S/ ${totalCosto.toFixed(2)}`;
  template.URL_IMAGEN_BUZON = EMAIL_CONFIG.MAILBOX_IMAGE_URL;

  const subject = `Hub ${EMAIL_CONFIG.CAMPUS_NAME} | ⚠️ Aviso de recarga por devolución pendiente de libro`;

  const htmlBody = template.evaluate().getContent();

  if (sendEmail(email, subject, htmlBody)) {
    // Actualizar log de TODOS los registros
    rowNumbers.forEach((rowNumber) => {
      const currentLog = SHEETS.overdueItems.getRange(rowNumber, COLUMNS.LOG + 1).getValue();
      updateActionLog(rowNumber, "💳 Aviso de recarga enviado", currentLog);
    });
  }
};

/**
 * Envía confirmación de pago de recarga
 * @param {Array<Array>} dataItems - Array de registros del mismo usuario
 * @param {Array<number>} rowNumbers - Array de números de fila
 */
const sendRechargeConfirmation = (dataItems, rowNumbers) => {
  const firstItem = dataItems[0];
  const email = firstItem[COLUMNS.EMAIL];
  const nombre = firstItem[COLUMNS.FULL_NAME];

  // Combinar títulos de todos los libros
  const titulos = dataItems.map(item => item[COLUMNS.TITLE]);

  // Sumar costos (parsear texto a número)
  const costos = dataItems.map(item => {
    const costoStr = item[COLUMNS.COST] || "0.00";
    return parseFloat(costoStr.replace(/[^\d.]/g, '')) || 0;
  });
  const totalCosto = costos.reduce((sum, c) => sum + c, 0);

  // Cargar y procesar plantilla HTML
  const template = HtmlService.createTemplateFromFile('templates/emailRechargeConfirmation');
  template.NOMBRE = nombre;
  template.CAMPUS = EMAIL_CONFIG.CAMPUS_NAME;
  template.LIBROS = formatBookList(titulos);
  template.MONTO = `S/ ${totalCosto.toFixed(2)}`;
  template.URL_IMAGEN_BUZON = EMAIL_CONFIG.MAILBOX_IMAGE_URL;

  const subject = `Hub ${EMAIL_CONFIG.CAMPUS_NAME} | ⚠️ Confirmación de recargo por devolución pendiente`;

  const htmlBody = template.evaluate().getContent();

  if (sendEmail(email, subject, htmlBody)) {
    // Actualizar log de TODOS los registros
    rowNumbers.forEach((rowNumber) => {
      const currentLog = SHEETS.overdueItems.getRange(rowNumber, COLUMNS.LOG + 1).getValue();
      updateActionLog(rowNumber, "✅ Confirmación de recarga enviada", currentLog);
    });
  }
};
