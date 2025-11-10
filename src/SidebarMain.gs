/**
 * SidebarMain.gs
 * Controlador principal del panel lateral de Invitaciones
 * Autor: Liber Q. (EKA MINING)
 */

// Permite configurar de forma centralizada la hoja principal utilizada por el
// panel. Si otra parte del proyecto ya definió `SHEET_NAME`, la respetamos.
var SHEET_NAME =
  typeof this.SHEET_NAME === 'string' && this.SHEET_NAME
    ? this.SHEET_NAME
    : 'Invitaciones';
this.SHEET_NAME = SHEET_NAME;

// Estructura base de secciones y campos que el panel renderiza. Si en otra
// parte del código ya se definió `FIELD_SECTIONS` (por ejemplo, desde un
// archivo de configuración más detallado), la reutilizamos. En caso contrario,
// inicializamos con un arreglo vacío para evitar referencias no definidas en
// las plantillas HTML.
var FIELD_SECTIONS = Array.isArray(this.FIELD_SECTIONS) ? this.FIELD_SECTIONS : [];
this.FIELD_SECTIONS = FIELD_SECTIONS;

/* ============================
   📌 Abrir el panel lateral
============================ */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📂 Invitaciones')
    .addItem('Abrir panel', 'showSidebar')
    .addSeparator()
    .addItem('Sincronizar con Supabase', 'syncAllToSupabase')
    .addToUi();
}

/**
 * Muestra el panel lateral principal
 */
function showSidebar() {
  const template = HtmlService.createTemplateFromFile('Sidebar.html');
  const templateData = buildSidebarTemplateData();

  template.sheetName = templateData.sheetName;
  template.dictionarySheetName = templateData.dictionarySheetName;
  template.permissionsSheetName = templateData.permissionsSheetName;
  template.columnPermissionSheetName = templateData.columnPermissionSheetName;
  template.fieldSectionsJson = JSON.stringify(templateData.fieldSections);

  const html = template
    .evaluate()
    .setTitle('Gestión de Invitaciones')
    .setWidth(480);
  SpreadsheetApp.getUi().showSidebar(html);
}

function buildSidebarTemplateData() {
  const ss = SpreadsheetApp.getActive();
  const configuredSheetName = SHEET_NAME || '';
  const fallbackSheetName = ss && ss.getActiveSheet() ? ss.getActiveSheet().getName() : '';

  const sheetName = configuredSheetName || fallbackSheetName || 'Hoja activa';
  const dictionaryName =
    typeof COLUMN_DICTIONARY_SHEET_NAME === 'string' && COLUMN_DICTIONARY_SHEET_NAME
      ? COLUMN_DICTIONARY_SHEET_NAME
      : 'Diccionario Invitaciones';
  const permissionsName =
    typeof PERMISSIONS_SHEET_NAME === 'string' && PERMISSIONS_SHEET_NAME
      ? PERMISSIONS_SHEET_NAME
      : 'Permisos Invitaciones';
  const columnPermissionsName =
    typeof COLUMN_PERMISSION_SHEET_NAME === 'string' && COLUMN_PERMISSION_SHEET_NAME
      ? COLUMN_PERMISSION_SHEET_NAME
      : permissionsName;
  const fieldSections = Array.isArray(FIELD_SECTIONS) ? FIELD_SECTIONS : [];

  return {
    sheetName: sheetName,
    dictionarySheetName: dictionaryName,
    permissionsSheetName: permissionsName,
    columnPermissionSheetName: columnPermissionsName,
    fieldSections: fieldSections,
  };
}

/* ============================
   📄 Render Helpers
============================ */

/**
 * Incluye los archivos parciales dentro del HTML del Sidebar
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================
   🔎 Datos activos
============================ */

/**
 * Retorna la invitación activa (fila seleccionada)
 */
function getActiveInvitationData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitaciones');
  const range = sheet.getActiveCell();
  const row = range.getRow();
  if (row <= 1) return { sections: [] };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const data = {};

  headers.forEach((h, i) => data[h] = values[i]);

  const sections = [
    {
      title: "Información General",
      fields: [
        { header: "ID Invitación", alias: "ID Invitación", valueType: "text" },
        { header: "Descripción", alias: "Descripción", valueType: "textarea" },
        { header: "Cliente", alias: "Cliente", valueType: "text" },
        { header: "Zona de Trabajo", alias: "Zona de Trabajo", valueType: "text" },
        { header: "Estado Cotización", alias: "Estado de Cotización", valueType: "select", options: ["Pendiente", "En Proceso", "Presentado", "Ganado", "Perdido"] },
      ]
    },
    {
      title: "Fechas y Montos",
      fields: [
        { header: "Fecha de Invitación", alias: "Fecha de Invitación", valueType: "date" },
        { header: "Fecha de Presentación", alias: "Fecha de Presentación", valueType: "date" },
        { header: "Monto ofertado", alias: "Monto ofertado", valueType: "number" },
        { header: "Moneda", alias: "Moneda", valueType: "select", options: ["USD", "PEN"] },
      ]
    }
  ];

  return { rowNumber: row, values: data, sections };
}

/* ============================
   💾 Guardar y actualizar
============================ */

/**
 * Guarda o actualiza la fila activa
 */
function saveInvitation(obj) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitaciones');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const range = sh.getActiveCell();
  const row = range.getRow();
  if (row <= 1) throw new Error("Selecciona una fila válida.");

  Object.keys(obj).forEach(k => {
    const idx = headers.indexOf(k);
    if (idx !== -1) sh.getRange(row, idx + 1).setValue(obj[k]);
  });
}

/**
 * Actualiza invitación existente
 */
function updateInvitation(obj) {
  saveInvitation(obj);
}

/**
 * Elimina invitación activa
 */
function deleteActiveInvitation(row) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitaciones');
  sh.deleteRow(row);
}

/* ============================
   🗂 Diccionario de columnas
============================ */
function openColumnDictionarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = 'Diccionario';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  ss.setActiveSheet(sh);
}

/* ============================
   ☁️ Sincronización Supabase
============================ */
function syncToSupabase() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitaciones');
    const data = sh.getDataRange().getValues();
    const headers = data.shift();
    const rows = data.map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });
    const payload = rows.filter(r => r["ID Invitación"]);
    if (!payload.length) throw new Error("No hay datos válidos para sincronizar.");

    const url = SUPABASE_URL + '/rest/v1/' + SUPABASE_TABLE + '?on_conflict=invitation_id';
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        apikey: SUPABASE_ANON_KEY,
        Authorization: 'Bearer ' + SUPABASE_ANON_KEY,
        Prefer: 'resolution=merge-duplicates',
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const code = res.getResponseCode();
    if (code >= 200 && code < 300) {
      return "Sincronizado correctamente con Supabase.";
    } else {
      throw new Error("Error Supabase: " + res.getContentText());
    }

  } catch (err) {
    throw new Error("Sincronización fallida: " + err.message);
  }
}
