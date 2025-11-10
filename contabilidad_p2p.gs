function procesarCorreosAirtm() {
 const SHEET_ID = '1T18Xsdjp8CGMEiKB0nwlR6IwC3ccrerv196C96_PjkY';      // ← reemplaza con el ID de tu hoja
  const SHEET_NAME = 'air_usd_c_p';  // ← reemplaza con el nombre de la hoja
  const etiquetaProcesado = 'P2P';
  const remitente = 'noreply@airtm.com';

  const hoja = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  // Buscamos todos los correos con retiro o agregar (sin distinguir mayúsculas)
  const query = `from:${remitente} subject:completado -label:${etiquetaProcesado}`;
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

      // --- Extraer los datos del cuerpo ---
      const metodo = extraerDato(cuerpo, /Método de pago[\s\S]*?([A-Za-z0-9\s]+)\s*(Estado|ID)/i);
      const estado = extraerDato(cuerpo, /Estado[\s\S]*?([A-Za-z]+)\s*(ID|Fondos)/i);
      const idConfirmacion = extraerDato(cuerpo, /ID de confirmación[\s\S]*?([0-9A-Z]+)/i);
      const fondosEnviados = extraerDato(cuerpo, /Fondos enviados[\s\S]*?(\$[0-9.,\s]+USDC)/i);
      const fondosRecibidos = extraerDato(cuerpo, /Fondos recibidos[\s\S]*?(\$[0-9.,\s]+[A-Z]{3})/i);
      const tipoCambio = extraerDato(cuerpo, /Tipo de cambio neto[\s\S]*?(\$[0-9.,\sA-Z=]+)/i);
      const fechaEnvio = extraerDato(cuerpo, /Fecha de envío[\s\S]*?([0-9a-zA-Z\s:pm\-]+)/i);
      const idTransaccion = extraerDato(cuerpo, /ID de la transacción[\s\S]*?([A-Z0-9]+)/i);

      // Evitar filas vacías
      if (!metodo && !fondosRecibidos && !fondosEnviados) {
        Logger.log("Correo sin datos reconocibles, se omite.");
        hilo.addLabel(etiqueta);
        return;
      }

      // --- Insertar una fila nueva con los valores capturados ---
      const nuevaFila = hoja.getLastRow() + 1;
      hoja.getRange(nuevaFila + 1, 1).setValue(new Date());
      hoja.getRange(nuevaFila + 1, 2).setValue(asunto);
      hoja.getRange(nuevaFila + 1, 3).setValue(metodo);
      hoja.getRange(nuevaFila + 1, 4).setValue(estado);
      hoja.getRange(nuevaFila + 1, 5).setValue(idConfirmacion);
      hoja.getRange(nuevaFila + 1, 6).setValue(limpiarValor(fondosEnviados));
      hoja.getRange(nuevaFila + 1, 7).setValue(limpiarValor(fondosRecibidos));
      hoja.getRange(nuevaFila + 1, 8).setValue(tipoCambio);
      hoja.getRange(nuevaFila + 1, 9).setValue(fechaEnvio);
      hoja.getRange(nuevaFila + 1, 10).setValue(idTransaccion);

      // --- Procesar según tipo de correo ---
      if (asunto.includes("retiro")) {
        const valor = extraerNumero(fondosRecibidos);
        hoja.getRange(nuevaFila + 1, 4).setValue(valor); // Columna D
        arrastrarFormulas(hoja, nuevaFila + 1, [5, 6, 8, 9, 10, 11, 12]);
      } else if (asunto.includes("agregar")) {
        const valor = extraerNumeroAntesDeUSDC(cuerpo);
        hoja.getRange(nuevaFila + 1, 7).setValue(valor); // Columna G
        arrastrarFormulas(hoja, nuevaFila + 1, [5, 6, 8, 9, 10, 11, 12]);
      }

      // --- Etiquetar como procesado ---
      hilo.addLabel(etiqueta);
    });
  });
}

/**
 * Función genérica para extraer texto con RegEx
 */
function extraerDato(texto, regex) {
  const match = texto.match(regex);
  return match ? match[1].trim() : "";
}

/**
 * Limpia texto de símbolos $, USD, COP, USDC dejando solo el número
 */
function limpiarValor(valor) {
  if (!valor) return "";
  return valor.replace(/[^\d.,]/g, "").trim();
}

/**
 * Convierte "$189,424 COP" en 189424.00
 */
function extraerNumero(texto) {
  if (!texto) return "";
  const match = texto.match(/([0-9.,]+)/);
  if (!match) return "";
  let num = match[1].replace(/\./g, "").replace(",", ".");
  return parseFloat(num);
}

/**
 * Extrae número antes de "USDC" (para agregar)
 */
function extraerNumeroAntesDeUSDC(texto) {
  if (!texto) return "";
  const match = texto.match(/([0-9.,]+)\s*USDC/i);
  if (!match) return "";
  let num = match[1].replace(/\./g, "").replace(",", ".");
  return parseFloat(num);
}

/**
 * Copia las fórmulas desde la fila anterior en las columnas indicadas
 */
function arrastrarFormulas(hoja, filaDestino, columnas) {
  const filaOrigen = filaDestino - 1;
  columnas.forEach(col => {
    const formula = hoja.getRange(filaOrigen, col).getFormula();
    if (formula) hoja.getRange(filaDestino, col).setFormula(formula);
  });
}
