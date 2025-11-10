function procesarCorreosAirtm() {
  // === CONFIGURACIÓN ===
  const SHEET_ID = '1T18Xsdjp8CGMEiKB0nwlR6IwC3ccrerv196C96_PjkY';      // ← reemplaza con el ID de tu hoja
  const SHEET_NAME = 'air_usd_c_p';  // ← reemplaza con el nombre de la hoja
  const etiquetaProcesado = 'P2P';
  const remitente = 'noreply@airtm.com';
  
  // === CONEXIÓN A SHEET ===
  const hoja = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);

  // === BÚSQUEDA DE CORREOS ===
  const query = `from:${remitente} subject:(retiro completado OR agregar completado) -label:${etiquetaProcesado}`;
  const hilos = GmailApp.search(query);

  if (hilos.length === 0) {
    Logger.log("No hay correos nuevos para procesar.");
    return;
  }

  hilos.forEach(hilo => {
    const mensajes = hilo.getMessages();
    mensajes.forEach(msg => {
      const asunto = msg.getSubject().toLowerCase();
      const cuerpo = msg.getPlainBody();
      
      // Extraer los campos relevantes del cuerpo del mensaje
      const metodo = extraerDato(cuerpo, "Método de pago");
      const estado = extraerDato(cuerpo, "Estado");
      const idConfirmacion = extraerDato(cuerpo, "ID de confirmación");
      const fondosEnviados = extraerDato(cuerpo, "Fondos enviados");
      const fondosRecibidos = extraerDato(cuerpo, "Fondos recibidos");
      const tipoCambio = extraerDato(cuerpo, "Tipo de cambio neto");
      const fechaEnvio = extraerDato(cuerpo, "Fecha de envío");
      const idTransaccion = extraerDato(cuerpo, "ID de la transacción");

      // Insertar una nueva fila con los datos
      const ultimaFila = hoja.getLastRow() + 1;
      hoja.insertRowAfter(hoja.getLastRow());
      hoja.getRange(ultimaFila + 1, 1).setValue(new Date());
      hoja.getRange(ultimaFila + 1, 2).setValue(asunto);
      hoja.getRange(ultimaFila + 1, 3).setValue(metodo);
      hoja.getRange(ultimaFila + 1, 4).setValue(estado);
      hoja.getRange(ultimaFila + 1, 5).setValue(idConfirmacion);
      hoja.getRange(ultimaFila + 1, 6).setValue(fondosEnviados);
      hoja.getRange(ultimaFila + 1, 7).setValue(fondosRecibidos);
      hoja.getRange(ultimaFila + 1, 8).setValue(tipoCambio);
      hoja.getRange(ultimaFila + 1, 9).setValue(fechaEnvio);
      hoja.getRange(ultimaFila + 1, 10).setValue(idTransaccion);

      // === PROCESO SEGÚN TIPO ===
      if (asunto.includes("retiro")) {
        const valor = extraerNumero(fondosRecibidos);
        hoja.getRange(ultimaFila + 1, 4).setValue(valor); // columna D
        arrastrarFormulas(hoja, ultimaFila + 1, [5, 6, 8, 9, 10, 11, 12]); // E, F, H, I, J, K, L
      } 
      else if (asunto.includes("agregar")) {
        const valor = extraerNumeroAntesDeUSDC(cuerpo);
        hoja.getRange(ultimaFila + 1, 7).setValue(valor); // columna G
        arrastrarFormulas(hoja, ultimaFila + 1, [5, 6, 8, 9, 10, 11, 12]); // E, F, H, I, J, K, L
      }

      // Etiquetar correo como procesado
      const etiqueta = GmailApp.createLabel(etiquetaProcesado);
      hilo.addLabel(etiqueta);
    });
  });
}

/**
 * Extrae el valor de un campo en el cuerpo del mensaje.
 */
function extraerDato(texto, campo) {
  const regex = new RegExp(`${campo}\\s*([\\s\\S]*?)\\n`, "i");
  const match = texto.match(regex);
  return match ? match[1].trim() : "";
}

/**
 * Extrae valor numérico de un texto como "$16.20 USD"
 */
function extraerNumero(texto) {
  const match = texto.match(/([0-9.,]+)/);
  return match ? parseFloat(match[1].replace(',', '.')) : "";
}

/**
 * Busca un número antes de la palabra USDC
 */
function extraerNumeroAntesDeUSDC(texto) {
  const match = texto.match(/([0-9.,]+)\s*USDC/i);
  return match ? parseFloat(match[1].replace(',', '.')) : "";
}

/**
 * Copia las fórmulas desde la fila anterior en las columnas dadas.
 */
function arrastrarFormulas(hoja, filaDestino, columnas) {
  const filaOrigen = filaDestino - 1;
  columnas.forEach(col => {
    const formula = hoja.getRange(filaOrigen, col).getFormula();
    if (formula) {
      hoja.getRange(filaDestino, col).setFormula(formula);
    }
  });
}
function mostrarCuerpoCorreoAirtm() {
  const remitente = 'noreply@airtm.com';
  const query = `from:${remitente} subject:(retiro completado OR agregar completado)`;
  
  const hilos = GmailApp.search(query);
  if (hilos.length === 0) {
    Logger.log("No se encontraron correos que coincidan con el filtro.");
    return;
  }

  // Tomamos solo el más reciente
  const mensaje = hilos[0].getMessages().pop();
  
  const asunto = mensaje.getSubject();
  const cuerpoPlano = mensaje.getPlainBody();
  const cuerpoHTML = mensaje.getBody();

  Logger.log("==== ASUNTO ====");
  Logger.log(asunto);

  Logger.log("==== CUERPO PLANO ====");
  Logger.log(cuerpoPlano);

  Logger.log("==== CUERPO HTML ====");
  Logger.log(cuerpoHTML);
}
