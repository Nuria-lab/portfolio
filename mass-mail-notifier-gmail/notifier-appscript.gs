function enviarCorreos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getActiveSheet();
  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0];

  const iNombre = encabezados.indexOf("First Name [Required]");
  const iApellido = encabezados.indexOf("Last Name [Required]");
  const iCuenta = encabezados.indexOf("Email Address [Required]");
  const iContrasena = encabezados.indexOf("Password [Required]");
  const iAlternativo = encabezados.indexOf("Recovery Email");
  const iNotificado = encabezados.findIndex(h => h.toString().toLowerCase().trim() === "notificado");
  const iFechaEnvio = encabezados.findIndex(h => h.toString().toLowerCase().trim() === "fechaenvio");

  if (iNotificado === -1 || iFechaEnvio === -1) {
    Logger.log("❌ No se encontraron las columnas 'notificado' o 'fechaEnvio'.");
    SpreadsheetApp.getUi().alert("❌ No se encontraron las columnas 'notificado' o 'fechaEnvio'.");
    return;
  }

  let enviados = 0;
  let errores = 0;

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const valNotificado = fila[iNotificado];
    const yaNotificado = (typeof valNotificado === "string" && valNotificado.toLowerCase() === "ok");
    const correo = fila[iAlternativo];

    if (!yaNotificado && correo && correo.includes("@")) {
      const nombre = fila[iNombre];
      const apellido = fila[iApellido];
      const cuenta = fila[iCuenta];
      const contrasena = fila[iContrasena];

      const asunto = "Cuenta institucional  - no responder";
      const mensaje = `
Hola ${nombre} ${apellido},

Se ha creado tu cuenta institucional de alumno/a.

 Cuenta: ${cuenta}
 Contraseña: ${contrasena}

Ingresá en https://accounts.google.com con estos datos.

Vas a tener que cambiar la contraseña al primer ingreso y por favor agregá un mail alternativo para recuperación de contraseña.

Saludos,
Equipo TIC
⚠️ Este es un mensaje automático enviado desde la cuenta de soporte. Por favor, no respondas a este correo.⚠️
      `.trim();

      try {
        GmailApp.sendEmail(correo, asunto, mensaje);
        hoja.getRange(i + 1, iNotificado + 1).setValue("OK");
        hoja.getRange(i + 1, iFechaEnvio + 1).setValue(new Date());
        Logger.log(`✅ Notificación enviada a: ${correo}`);
        enviados++;
      } catch (e) {
        hoja.getRange(i + 1, iNotificado + 1).setValue("ERROR");
        Logger.log(`❌ Error al enviar a ${correo}: ${e.message}`);
        errores++;
      }
    }
  }

  SpreadsheetApp.getUi().alert(`Notificaciones enviadas: ${enviados}\nErrores: ${errores}`);
  Logger.log("Script finalizado.");
}
