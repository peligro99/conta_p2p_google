function procesarCorreosAirtm() {
  const SHEET_ID = '1T18Xsdjp8CGMEiKB0nwlR6IwC3ccrerv196C96_PjkY';
  const SHEET_NAME = 'air_usd_c_p';
  const etiquetaProcesado = 'P2P';
  const remitente = 'noreply@airtm.com';

  const hoja = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);

  // Filtro mejorado
  const query = `from:${remitente} subject:completado ("retiro" OR "agregar" OR "agregaste" OR "agregando") -label:${etiquetaProcesado}`;
  const hilos = GmailApp.search(query);

  if (hilos.length === 0) {
    Logger.log("No hay correos nuevos para procesar.");
    return;
  }

  const etiqueta = GmailApp.createLabel(etiquetaProcesado);

  hilos.forEach(hilo => {
    const mensajes = hilo.getMessages();

    mensajes.forEach(msg => {

      const asunto = msg.getSubject().toLowerCase();
      const cuerpo = msg.getPlainBody() || msg.getBody();

      // ===== EXTRACCIÓN DE DATOS =====
      const metodo = extraerDato(cuerpo, /Método de pago[\s\S]*?([A-Za-z0-9\s]+)\s*(Estado|ID)/i);
      const estado = extraerDato(cuerpo, /Estado[\s\S]*?([A-Za-z]+)\s*(ID|Fondos)/i);
      const idConfirmacion = extraerDato(cuerpo, /ID de confirmación[\s\S]*?([0-9A-Z]+)/i);
      const fondosEnviados = extraerDato(cuerpo, /Fondos enviados[\s\S]*?(\$[0-9.,\s]+USDC)/i);
      const fondosRecibidos = extraerDato(cuerpo, /Fondos recibidos[\s\S]*?(\$[0-9.,\s]+[A-Z]{3})/i);
      const tipoCambio = extraerDato(cuerpo, /Tipo de cambio neto[\s\S]*?(\$[0-9.,\sA-Z=]+)/i);

      // Fecha robusta
      const fechaEnvio = extraerDato(
        cuerpo,
        /Fecha de envío[\s\S]*?([\d]{1,2}\s+de\s+[a-zA-Z]+\s+de\s+\d{4}\s+[0-9:apm\s\-]+)/i
      );

      const idTransaccion = extraerDato(cuerpo, /ID de la transacción[\s\S]*?([A-Z0-9]+)/i);

      if (!metodo && !fondosEnviados && !fondosRecibidos) {
        Logger.log("Correo sin datos reconocibles, omitido.");
        hilo.addLabel(etiqueta);
        return;
      }

      // ===== INSERTAR NUEVA FILA =====
      const nuevaFila = hoja.getLastRow() + 1;

      hoja.getRange(nuevaFila, 1).setValue(new Date());
      hoja.getRange(nuevaFila, 2).setValue(asunto);
      hoja.getRange(nuevaFila, 3).setValue(metodo);
      hoja.getRange(nuevaFila, 4).setValue(estado);
      hoja.getRange(nuevaFila, 5).setValue(idConfirmacion);
      hoja.getRange(nuevaFila, 6).setValue(limpiarValor(fondosEnviados));
      hoja.getRange(nuevaFila, 7).setValue(limpiarValor(fondosRecibidos));
      hoja.getRange(nuevaFila, 8).setValue(tipoCambio);
      hoja.getRange(nuevaFila, 9).setValue(fechaEnvio); // NO fórmula
      hoja.getRange(nuevaFila, 10).setValue(idTransaccion);

      // ===== PROCESO SEGÚN TIPO =====
      if (asunto.includes("retiro")) {
        const valor = extraerNumero(fondosRecibidos);
        hoja.getRange(nuevaFila, 4).setValue(valor); // Columna D

        arrastrarFormulas(hoja, nuevaFila, [5, 6, 8, 10, 11, 12]);
      }

      if (asunto.includes("agregar")) {
        const valor = extraerNumeroAntesDeUSDC(cuerpo);
        hoja.getRange(nuevaFila, 7).setValue(valor); // Columna G

        arrastrarFormulas(hoja, nuevaFila, [5, 6, 8, 10, 11, 12]);
      }

      // ===== EXPORTAR PDF =====
      try {
        const nombrePDF = `${fechaEnvio}_${asunto}_${idTransaccion}`.replace(/[^\w\s.-]/g, "_");
        guardarCorreoComoPDF(msg, nombrePDF);
        Logger.log("PDF guardado: " + nombrePDF);
      } catch (e) {
        Logger.log("Error exportando PDF: " + e);
      }

      // ===== ETIQUETAR COMO PROCESADO =====
      hilo.addLabel(etiqueta);

    });
  });
}

//
// ===== FUNCIONES AUXILIARES =====
//

function extraerDato(texto, regex) {
  const match = texto.match(regex);
  return match ? match[1].trim() : "";
}

function limpiarValor(valor) {
  if (!valor) return "";
  return valor.replace(/[^\d.,]/g, "").trim();
}

function extraerNumero(texto) {
  if (!texto) return "";
  const match = texto.match(/([0-9.,]+)/);
  if (!match) return "";
  let num = match[1].replace(/\./g, "").replace(",", ".");
  return parseFloat(num);
}

function extraerNumeroAntesDeUSDC(texto) {
  if (!texto) return "";
  const match = texto.match(/([0-9.,]+)\s*USDC/i);
  if (!match) return "";
  let num = match[1].replace(/\./g, "").replace(",", ".");
  return parseFloat(num);
}

function arrastrarFormulas(hoja, filaDestino, columnas) {
  const filaOrigen = filaDestino - 1;
  columnas.forEach(col => {
    const f = hoja.getRange(filaOrigen, col).getFormula();
    if (f) hoja.getRange(filaDestino, col).setFormula(f);
  });
}

//
// ===== EXPORTAR PDF =====
//

function guardarCorreoComoPDF(msg, nombrePDF) {
  const carpetaDestinoId = "1aSve6SdjM5dp044CPWP8mgcMFSmOUTMa";
  const carpeta = DriveApp.getFolderById(carpetaDestinoId);

  // Construimos HTML para exportar el correo
  const html = `
    <html>
      <body style="font-family: Arial; font-size: 12px;">
        <h2>${msg.getSubject()}</h2>
        <p><b>Fecha:</b> ${msg.getDate()}</p>
        <hr>
        ${msg.getBody()}
      </body>
    </html>
  `;

  // Convertimos el HTML a un blob
  const blob = Utilities.newBlob(html, "text/html", "temp.html");

  // Convertimos ese blob a PDF
  const pdf = blob.getAs("application/pdf").setName(nombrePDF + ".pdf");

  // Guardamos el PDF en Drive
  carpeta.createFile(pdf);
}
