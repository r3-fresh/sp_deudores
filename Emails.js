// ========================================
// ENVÍO DE CORREOS ELECTRÓNICOS
// ========================================

/**
 * Envía un correo electrónico genérico
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
 * Crea el cuerpo HTML básico para un recordatorio
 * @param {Object} data - Datos del préstamo y deudor
 * @returns {string} HTML del correo
 */
const createReminderEmailBody = (data) => {
    const nombre = data[COLUMNS.NAME];
    const apellido = data[COLUMNS.LASTNAME];
    const titulo = data[COLUMNS.TITLE];
    const biblioteca = data[COLUMNS.LIBRARY];
    const fechaVencimiento = data[COLUMNS.DUE_DATE];

    return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
        .container { max-width: 600px; margin: 0 auto; padding: 20px; }
        .header { background-color: #5A00AA; color: white; padding: 20px; text-align: center; }
        .content { background-color: #f9f9f9; padding: 20px; }
        .footer { text-align: center; margin-top: 20px; font-size: 12px; color: #666; }
        .highlight { background-color: #fff3cd; padding: 10px; border-left: 4px solid #ffc107; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h2>📚 Hub de Información - UC Continental</h2>
        </div>
        <div class="content">
          <p>Hola <strong>${nombre} ${apellido}</strong>,</p>
          <p>Te recordamos que tienes un recurso pendiente de devolución:</p>
          
          <div class="highlight">
            <p><strong>📖 Recurso:</strong> ${titulo}</p>
            <p><strong>📍 Biblioteca:</strong> ${biblioteca}</p>
            <p><strong>📅 Fecha de vencimiento:</strong> ${fechaVencimiento}</p>
          </div>
          
          <p>Por favor, realiza la devolución a la brevedad posible para evitar sanciones.</p>
          <p><strong>Importante:</strong> Si ya devolviste este recurso, ignora este mensaje.</p>
        </div>
        <div class="footer">
          <p>Hub de Información - Universidad Continental</p>
          <p>Este es un mensaje automático, por favor no responder.</p>
        </div>
      </div>
    </body>
    </html>
  `;
};

/**
 * Envía primer recordatorio al deudor
 * @param {Array} data - Datos del registro
 * @param {number} rowNumber - Número de fila
 */
const sendFirstReminder = (data, rowNumber) => {
    const email = data[COLUMNS.EMAIL];
    const nombre = data[COLUMNS.NAME];
    const subject = "📚 Recordatorio: Devolución de recurso pendiente";
    const body = createReminderEmailBody(data);

    if (sendEmail(email, subject, body)) {
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
 * Envía segundo recordatorio al deudor (tono más firme)
 * @param {Array} data - Datos del registro
 * @param {number} rowNumber - Número de fila
 */
const sendSecondReminder = (data, rowNumber) => {
    const email = data[COLUMNS.EMAIL];
    const nombre = data[COLUMNS.NAME];
    const apellido = data[COLUMNS.LASTNAME];
    const titulo = data[COLUMNS.TITLE];
    const fechaVencimiento = data[COLUMNS.DUE_DATE];

    const subject = "⚠️ URGENTE: Segundo recordatorio - Devolución pendiente";
    const body = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
        .container { max-width: 600px; margin: 0 auto; padding: 20px; }
        .header { background-color: #dc3545; color: white; padding: 20px; text-align: center; }
        .content { background-color: #f9f9f9; padding: 20px; }
        .warning { background-color: #f8d7da; padding: 15px; border-left: 4px solid #dc3545; margin: 15px 0; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h2>⚠️ SEGUNDO RECORDATORIO</h2>
        </div>
        <div class="content">
          <p>Estimado/a <strong>${nombre} ${apellido}</strong>,</p>
          <p>Este es nuestro <strong>segundo recordatorio</strong> sobre el siguiente recurso pendiente:</p>
          
          <div class="warning">
            <p><strong>📖 Recurso:</strong> ${titulo}</p>
            <p><strong>📅 Venció el:</strong> ${fechaVencimiento}</p>
          </div>
          
          <p><strong>Es necesario que realices la devolución de inmediato</strong> para evitar sanciones académicas.</p>
          <p>Si tienes algún inconveniente, por favor comunícate con nosotros.</p>
        </div>
      </div>
    </body>
    </html>
  `;

    if (sendEmail(email, subject, body)) {
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
    const subject = "💳 Aviso de recarga por mora en devolución";
    const body = createReminderEmailBody(data); // TODO: Personalizar para recarga

    if (sendEmail(email, subject, body)) {
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
    const subject = "✅ Confirmación de pago de recarga";
    const body = createReminderEmailBody(data); // TODO: Personalizar para confirmación

    if (sendEmail(email, subject, body)) {
        const currentLog = SHEETS.overdueItems
            .getRange(rowNumber, COLUMNS.LOG + 1)
            .getValue();
        updateActionLog(rowNumber, "✅ Confirmación de recarga enviada", currentLog);
    }
};
