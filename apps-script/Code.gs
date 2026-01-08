/**
 * ========= CONFIG =========
 * 1) TEMPLATE_DOC_ID: Ya está con tu ID real.
 * 2) OUTPUT_FOLDER_ID: Pega el ID de la carpeta de Drive donde se guardarán las cotizaciones.
 *    (Si no pones carpeta, igual funciona, pero se guardará en "Mi unidad" raíz.)
 */
const TEMPLATE_DOC_ID = "1Z7QHT5F80E-RPEzi_iH7gdvQ1tC0ZU9GsDrk4HJXHMo";
const OUTPUT_FOLDER_ID = ""; // CARPETA "".

/**
 * Nombre de la hoja donde está tu tabla.
 * Si tu hoja se llama distinto, cámbialo (ej: "Cotizaciones").
 */
const SHEET_NAME = "Hoja 1";

/**
 * ========= MENÚ =========
 * Te crea un menú en el Sheets para generar cotizaciones fácil.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📄 Cotizaciones")
    .addItem("Generar cotización (fila activa)", "generarCotizacionDesdeFilaActiva")
    .addItem("Generar cotizaciones (marcadas TRUE)", "generarCotizacionesMarcadas")
    .addToUi();
}

/**
 * Genera la cotización tomando la fila donde está tu celda seleccionada.
 * Ideal para asignar a un botón.
 */
function generarCotizacionDesdeFilaActiva() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("No existe la hoja: " + SHEET_NAME);

  const row = sheet.getActiveCell().getRow();
  if (row === 1) {
    SpreadsheetApp.getUi().alert("Selecciona una fila de cotización (no el encabezado).");
    return;
  }

  generarCotizacionParaFila(sheet, row);
}

/**
 * Recorre todas las filas y genera cotización donde generar_contrato sea TRUE.
 */
function generarCotizacionesMarcadas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("No existe la hoja: " + SHEET_NAME);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colIndex = indexHeaders(headers);

  if (!colIndex.generar_contrato) {
    throw new Error("No existe la columna 'generar_contrato'. Revisa el encabezado exacto.");
  }

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  values.forEach((r, i) => {
    const rowNumber = i + 2;
    const flag = r[colIndex.generar_contrato - 1];
    if (flag === true || String(flag).toUpperCase().trim() === "TRUE") {
      generarCotizacionParaFila(sheet, rowNumber);
    }
  });
}

/**
 * ========= CORE =========
 * Crea copia del Docs plantilla, reemplaza {{variables}}, y escribe el link en la hoja.
 */
function generarCotizacionParaFila(sheet, row) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colIndex = indexHeaders(headers);

  const rowValues = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const data = rowToObject(headers, rowValues);

  // Validaciones mínimas
  if (!data.id_cotizacion) {
    SpreadsheetApp.getUi().alert(`La fila ${row} no tiene id_cotizacion. Completa ese campo.`);
    return;
  }

  // Fecha de generación: si está vacía, poner hoy
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const fechaGen = (data.fecha_generacion && String(data.fecha_generacion).trim() !== "")
    ? String(data.fecha_generacion)
    : Utilities.formatDate(now, tz, "yyyy-MM-dd");

  data.fecha_generacion = fechaGen;

  // ===== Crear copia del documento =====
  const templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);

  let copy;
  const docName = `Cotizacion_${data.id_cotizacion}_${(data.empresa_cliente || "").toString().trim()}`.trim();

  if (OUTPUT_FOLDER_ID && OUTPUT_FOLDER_ID.trim() !== "") {
    const folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID.trim());
    copy = templateFile.makeCopy(docName, folder);
  } else {
    copy = templateFile.makeCopy(docName);
  }

  const docId = copy.getId();
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  // ===== Reemplazar variables tipo {{campo}} =====
  Object.keys(data).forEach(key => {
    const value = (data[key] === null || data[key] === undefined) ? "" : String(data[key]);
    body.replaceText(`\\{\\{${escapeRegExp(key)}\\}\\}`, value);
  });

  doc.saveAndClose();

  // Link del documento generado
  const docUrl = `https://docs.google.com/document/d/${docId}/edit`;

  // ===== Escribir resultados en la hoja =====
  if (colIndex.fecha_generacion) {
    sheet.getRange(row, colIndex.fecha_generacion).setValue(fechaGen);
  }
  if (colIndex.link_cotizacion) {
    sheet.getRange(row, colIndex.link_cotizacion).setValue(docUrl);
  }

  // (Opcional) Desmarcar generar_contrato
  if (colIndex.generar_contrato) {
    sheet.getRange(row, colIndex.generar_contrato).setValue(false);
  }
}

/**
 * ========= HELPERS =========
 */
function indexHeaders(headers) {
  const idx = {};
  headers.forEach((h, i) => {
    const key = String(h).trim();
    if (key) idx[key] = i + 1; // columnas 1-based
  });
  return idx;
}

function rowToObject(headers, values) {
  const obj = {};
  headers.forEach((h, i) => {
    const key = String(h).trim();
    if (!key) return;
    obj[key] = values[i];
  });
  return obj;
}

function escapeRegExp(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
