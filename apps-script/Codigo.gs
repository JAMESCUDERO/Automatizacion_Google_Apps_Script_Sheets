/**
 * CONFIGURACIÓN CENTRAL
 * Permite modificar reglas sin tocar la lógica principal
 */
const CONFIG = {
  REQUESTS_SHEET: "SOLICITUDES",
  CONTROL_SHEET: "LISTAS",

  STATUS_COLUMN: 9,           // Columna I
  TIMESTAMP_OFFSET: 1,        // Columna J
  MIN_DATA_ROW: 2,

  CONTEXT_MIN_ROW: 820,
  CONTEXT_COLUMN: 3,          // Columna C
  CONTEXT_TARGET_CELL: "N21",

  ASSIGNMENT_SOURCE_CELL: "O21"
};

/**
 * Trigger principal que orquesta la automatización
 * Se ejecuta automáticamente en cada edición del Spreadsheet
 */
function onEdit(e) {
  registerStatusTimestamp(e);
  syncActiveValueToControlSheet(e);
}

/**
 * Registra automáticamente una marca de tiempo
 * cuando se edita una celda específica de la hoja SOLICITUDES
 */
function registerStatusTimestamp(e) {
  const sheet = e.range.getSheet();

  if (sheet.getName() !== CONFIG.REQUESTS_SHEET) return;

  const row = e.range.getRow();
  const column = e.range.getColumn();

  if (row < CONFIG.MIN_DATA_ROW) return;
  if (column !== CONFIG.STATUS_COLUMN) return;

  const timestampCell = e.range.offset(0, CONFIG.TIMESTAMP_OFFSET);
  timestampCell.setValue(new Date());
}

/**
 * Sincroniza el valor de la celda activa hacia la hoja de control
 * cuando se cumplen condiciones de posición y contexto
 */
function syncActiveValueToControlSheet(e) {
  const sheet = e.range.getSheet();
  const spreadsheet = sheet.getParent();

  const row = e.range.getRow();
  const column = e.range.getColumn();
  const value = e.range.getValue();

  if (row <= CONFIG.CONTEXT_MIN_ROW) return;
  if (column !== CONFIG.CONTEXT_COLUMN) return;

  const controlSheet = spreadsheet.getSheetByName(CONFIG.CONTROL_SHEET);
  if (!controlSheet) return;

  controlSheet
    .getRange(CONFIG.CONTEXT_TARGET_CELL)
    .setValue(value);
}

/**
 * Asigna automáticamente un responsable a una solicitud
 * copiando el valor calculado en la hoja de control
 */
function assignResponsibleFromRulesEngine() {
  const spreadsheet = SpreadsheetApp.getActive();

  const controlSheet = spreadsheet.getSheetByName(CONFIG.CONTROL_SHEET);
  const requestsSheet = spreadsheet.getSheetByName(CONFIG.REQUESTS_SHEET);

  if (!controlSheet || !requestsSheet) return;

  const sourceValue = controlSheet
    .getRange(CONFIG.ASSIGNMENT_SOURCE_CELL)
    .getValue();

  const targetCell = requestsSheet.getActiveCell();
  targetCell.setValue(sourceValue);
}
