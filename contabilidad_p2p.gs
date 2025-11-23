function extraerPrimeroQueCoincida(texto, listaRegex) {
  for (const rgx of listaRegex) {
    const m = texto.match(rgx);
    if (m && m[1]) return m[1].trim();
  }
  return "";
}
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
      let fondosEnviados = extraerPrimeroQueCoincida(cuerpo, [
  /Fondos enviados[\s\S]*?(\$?[0-9.,\s]+[A-Z]{3})/i,
  /Fondos a enviar[\s\S]*?(\$?[0-9.,\s]+[A-Z]{3})/i
]);

let fondosRecibidos = extraerPrimeroQueCoincida(cuerpo, [
  /Fondos recibidos[\s\S]*?(\$?[0-9.,\s]+[A-Z]{3})/i,
  /Fondos a recibir[\s\S]*?(\$?[0-9.,\s]+[A-Z]{3})/i
]);
      const tipoCambio = extraerDato(cuerpo, /Tipo de cambio neto[\s\S]*?(\$[0-9.,\sA-Z=]+)/i);

      // Fecha robusta
      const fechaEnvio = extraerDato(
        cuerpo,
        /Fecha de envío[\s\S]*?([\d]{1,2}\s+de\s+[a-zA-Z]+\s+de\s+\d{4}\s+[0-9:apm\s\-]+)/i
      );

      const idTransaccion = extraerDato(cuerpo, /ID de la transacción[\s\S]*?([A-Z0-9]+)/i);

     if (!metodo && !fondosEnviados && !fondosRecibidos) {
  Logger.log(`Correo sin datos reconocibles, omitido. ASUNTO: "${asunto}"`);
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
      const fechaFormateada = formatearFechaAAAAMMDD(fechaEnvio);
      const asuntoLimpio = limpiarAsuntoParaPDF(asunto);
      const nombrePDF = `${fechaFormateada}_${asuntoLimpio}_${idTransaccion}.pdf`;
        guardarCorreoComoPDFConImagenes(msg, nombrePDF);
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

function guardarCorreoComoPDFConImagenes(msg, nombrePDF) {
  const carpetaDestinoId = "1aSve6SdjM5dp044CPWP8mgcMFSmOUTMa";
  const carpeta = DriveApp.getFolderById(carpetaDestinoId);

  // Obtener HTML original del mensaje (si no hay, usamos body plano)
  let html = "";
  try {
    html = msg.getBody() || msg.getPlainBody() || "";
  } catch (e) {
    html = msg.getPlainBody() || "";
  }

  // 1) Reemplazar imágenes inline (cid:) usando attachments inline
  try {
    const atts = msg.getAttachments({ includeInlineImages: true });
    if (atts && atts.length > 0) {
      atts.forEach(att => {
        try {
          const ct = att.getContentType();
          if (!ct || !ct.match(/^image\//i)) return; // solo imágenes

          // Intentar obtener contentId si existe (varía según implementación)
          let cid = "";
          try { cid = att.getContentId(); } catch (e) { cid = ""; }

          // Crear data URI
          const bytes = att.getBytes();
          const b64 = Utilities.base64Encode(bytes);
          const dataUri = "data:" + ct + ";base64," + b64;

          // Reemplazar referencias por CID y por nombre en el HTML
          if (cid) {
            // src="cid:XYZ" o src='cid:XYZ'
            const reCid = new RegExp('(["\'])cid:' + escapeRegExp(cid) + '\\1', 'gi');
            html = html.replace(reCid, '"' + dataUri + '"');
            // también sin comillas
            html = html.replace(new RegExp('cid:' + escapeRegExp(cid), 'gi'), dataUri);
          }

          // Reemplazar por nombre de archivo también (por si lo referencia así)
          const name = att.getName();
          if (name) {
            const reName = new RegExp('(["\'])([^"\']*' + escapeRegExp(name) + ')[\\"\\\']', 'gi');
            html = html.replace(reName, '"' + dataUri + '"');
            html = html.replace(new RegExp(escapeRegExp(name), 'gi'), dataUri);
          }

        } catch (e) {
          // ignorar una attachment fallida y seguir
          Logger.log("Error procesando attachment inline: " + e);
        }
      });
    }
  } catch (e) {
    Logger.log("No se pudieron obtener attachments inline: " + e);
  }

  // 2) Buscar <img src="http..."> y tratar de descargar y reemplazar por data URI
  try {
    // encontrar todas las URLs de imagen en src="" o src=''
    const imgUrlRegex = /<img[^>]+src=(?:'|")([^'">]+)(?:'|")[^>]*>/gi;
    let match;
    const urlsProcesadas = {};
    while ((match = imgUrlRegex.exec(html)) !== null) {
      const src = match[1];
      // ignorar data: ya embebidas
      if (!src || src.trim().toLowerCase().startsWith('data:')) continue;
      // evitar procesar la misma url muchas veces
      if (urlsProcesadas[src]) continue;
      urlsProcesadas[src] = true;

      try {
        // Intento de descarga (timeout corto)
        const resp = UrlFetchApp.fetch(src, { muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: true, timeout: 10000 });
        if (resp.getResponseCode() === 200) {
          const contentType = resp.getHeaders()['Content-Type'] || resp.getHeaders()['content-type'] || '';
          if (contentType && contentType.toLowerCase().indexOf('image/') === 0) {
            const bytes = resp.getContent();
            const b64 = Utilities.base64Encode(bytes);
            const dataUri = 'data:' + contentType.split(';')[0] + ';base64,' + b64;
            // Reemplazar todas las ocurrencias de la URL por dataUri
            const escUrl = escapeRegExp(src);
            html = html.replace(new RegExp(escUrl, 'g'), dataUri);
          } else {
            // no es imagen, lo dejamos
            Logger.log('URL no es imagen o tipo desconocido: ' + src);
          }
        } else {
          Logger.log('No se pudo descargar imagen (' + resp.getResponseCode() + '): ' + src);
        }
      } catch (e) {
        Logger.log('Error descargando imagen externa: ' + src + " -> " + e);
      }
    }
  } catch (e) {
    Logger.log("Error procesando imágenes externas: " + e);
  }

  // 3) Generar PDF desde HTML ya con imágenes embebidas
  try {
    // Si el HTML viene sin <html> completo, envolvemos
    if (!/^<!doctype/i.test(html) && html.indexOf('<html') === -1) {
      html = '<!doctype html><html><head><meta charset="utf-8"></head><body>' + html + '</body></html>';
    }

    const blob = Utilities.newBlob(html, 'text/html', 'temp.html');
    const pdf = blob.getAs('application/pdf').setName(nombrePDF + '.pdf');
    carpeta.createFile(pdf);
  } catch (e) {
    Logger.log("Error generando PDF: " + e);
    throw e;
  }
}


// Helper: escapar texto para usar en RegExp
function escapeRegExp(string) {
  return String(string).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ==== FECHA FORMATO YYYYMMDD ====
function formatearFechaAAAAMMDD(fechaTexto) {
  try {
    const fecha = new Date(fechaTexto.replace(" de ", " ").replace("de ", ""));
    const yyyy = fecha.getFullYear();
    const mm = ("0" + (fecha.getMonth() + 1)).slice(-2);
    const dd = ("0" + fecha.getDate()).slice(-2);
    return `${yyyy}${mm}${dd}`;
  } catch (e) {
    return "00000000";
  }
}

//
// ==== LIMPIAR ASUNTO PARA NOMBRE DEL PDF ====
function limpiarAsuntoParaPDF(asunto) {
  return asunto
    .toLowerCase()
    .replace("retiro", "")
    .replace("agregar", "")
    .replace("completado", "")
    .trim()
    .replace(/\s+/g, " ");
}
