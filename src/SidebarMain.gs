const SIDEBAR_TITLE = 'Gestión de Invitaciones';
const SHEET_NAME = 'Invitaciones';
const HEADER_ROW = 1;

const COLUMN_MAP = {
  invitation_id: 'ID Invitación',
  cotizacion: 'Cotización',
  descripcion: 'Descripción',
  tipo_servicio: 'Tipo de Servicio',
  estado_cotizacion: 'Estado de Cotización',
  fecha_registro: 'Fecha de Registro',
  cliente: 'Cliente',
  zona_trabajo: 'Zona de Trabajo',
  solicitante: 'Solicitante',
  correo_solicitante: 'Correo del Solicitante',
  telefono_solicitante: 'Teléfono del Solicitante',
  responsable_tecnico: 'Responsable Técnico',
  correo_responsable_tecnico: 'Correo del Resp. Téc.',
  telefono_responsable_tecnico: 'Teléfono del Resp. Téc.',
  responsable_economico: 'Responsable Económico',
  correo_responsable_economico: 'Correo del Resp. Eco.',
  telefono_responsable_economico: 'Teléfono del Resp. Eco.',
  fecha_invitacion: 'Fecha de Invitación',
  fecha_confirmacion: 'Fecha de Confirmación',
  fecha_visita_tecnica: 'Fecha de Visita Téc.',
  fecha_consultas: 'Fecha de Consultas',
  fecha_abs_consultas: 'Fecha de Abs. Consultas',
  fecha_presentacion: 'Fecha de Presentación',
  monto_ofertado: 'Monto ofertado',
  moneda: 'Moneda',
  estado_propuesta: 'Estado de Propuesta',
  orden_de_compra: 'Orden de Compra',
  fecha_orden_compra: 'Fecha de la Orden de Compra',
  link_carpeta_drive: 'Link Carpeta Drive',
  link_archivo_enviado: 'Link Archivo Enviado',
  fecha_envio_propuesta: 'Fecha de Envío de Propuesta',
  hora_envio_propuesta: 'Hora de Envío de Propuesta',
  dias_a_vencimiento: 'Días a Vencimiento',
  riesgo_d3: 'Riesgo D-3',
  riesgo_d1: 'Riesgo D-1',
  enviado_a_tiempo: 'Enviado a Tiempo',
  tiempo_respuesta_dias_hab: 'Tiempo de Respuesta (días hábiles)',
  requiere_visita: 'Requiere Visita Técnica',
  visita_ejecutada: 'Visita Ejecutada',
  semana_iso: 'Semana ISO',
  mes_anio: 'Mes-Año',
  moneda_normalizada_usd: 'Moneda Normalizada (USD)',
  tipo_cambio: 'Tipo de Cambio',
  monto_ofertado_usd: 'Monto Ofertado (USD)',
  estado_pipeline: 'Estado Pipeline',
  notas_kpi: 'Notas KPI'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PMO • EKA')
    .addItem('Abrir panel de invitaciones', 'showSidebar')
    .addItem('Asegurar columnas requeridas', 'ensureInvitationColumns')
    .addToUi();
}

function showSidebar() {
  const template = HtmlService.createTemplateFromFile('Sidebar');
  template.sheetName = getInvitationSheet().getName();
  const htmlOutput = template
    .evaluate()
    .setTitle(SIDEBAR_TITLE)
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function getInvitationSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('No se encontró la hoja "' + SHEET_NAME + '".');
  }
  return sheet;
}

function getHeaders() {
  const sheet = getInvitationSheet();
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    return [];
  }
  return sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0].map(String);
}

function getColumnIndex(alias) {
  const headerName = COLUMN_MAP[alias];
  if (!headerName) {
    throw new Error('Alias de columna desconocido: ' + alias);
  }
  const headers = getHeaders();
  const index = headers.indexOf(headerName);
  if (index === -1) {
    throw new Error('No se encontró la columna: ' + headerName);
  }
  return index + 1;
}

function getRowData(rowNumber) {
  if (rowNumber <= HEADER_ROW) {
    throw new Error('La fila activa debe estar debajo del encabezado.');
  }
  const sheet = getInvitationSheet();
  const lastColumn = sheet.getLastColumn();
  const values = sheet.getRange(rowNumber, 1, 1, lastColumn).getValues()[0];
  const result = { rowNumber };
  const headers = getHeaders();
  Object.keys(COLUMN_MAP).forEach(function (alias) {
    const headerName = COLUMN_MAP[alias];
    const columnIndex = headers.indexOf(headerName);
    if (columnIndex !== -1) {
      result[alias] = values[columnIndex];
    }
  });
  return result;
}

function getActiveInvitation() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAME) {
    return {
      rowNumber: null,
      sheetName: sheet.getName(),
      expectedSheetName: SHEET_NAME,
      data: null,
    };
  }
  const row = sheet.getActiveCell().getRow();
  if (row <= HEADER_ROW) {
    return {
      rowNumber: null,
      sheetName: SHEET_NAME,
      expectedSheetName: SHEET_NAME,
      data: null,
    };
  }
  return {
    rowNumber: row,
    sheetName: SHEET_NAME,
    expectedSheetName: SHEET_NAME,
    data: getRowData(row),
  };
}

function updateInvitationRow(rowNumber, updates) {
  if (!rowNumber || rowNumber <= HEADER_ROW) {
    throw new Error('Número de fila inválido para actualizar.');
  }
  const sheet = getInvitationSheet();
  const headers = getHeaders();
  Object.keys(updates || {}).forEach(function (alias) {
    if (!(alias in COLUMN_MAP)) {
      return;
    }
    const headerName = COLUMN_MAP[alias];
    const index = headers.indexOf(headerName);
    if (index === -1) {
      throw new Error('No se encontró la columna: ' + headerName);
    }
    const value = normalizeValueForSheet(alias, updates[alias]);
    sheet.getRange(rowNumber, index + 1).setValue(value);
  });
  return getRowData(rowNumber);
}

function normalizeValueForSheet(alias, value) {
  if (value === undefined || value === null) {
    return '';
  }
  switch (alias) {
    case 'dias_a_vencimiento':
    case 'tiempo_respuesta_dias_hab':
    case 'tipo_cambio':
    case 'monto_ofertado':
    case 'monto_ofertado_usd':
      if (value === '') {
        return '';
      }
      var numericValue = Number(value);
      if (Number.isFinite(numericValue)) {
        return numericValue;
      }
      return value;
    case 'riesgo_d3':
    case 'riesgo_d1':
    case 'enviado_a_tiempo':
    case 'requiere_visita':
    case 'visita_ejecutada':
      return value === true ? true : value === false ? false : '';
    default:
      return value;
  }
}

function syncActiveInvitation() {
  const active = getActiveInvitation();
  if (!active.rowNumber) {
    throw new Error('Selecciona una fila válida en la hoja "' + SHEET_NAME + '".');
  }
  const payload = buildInvitationPayload(active.rowNumber);
  if (!payload.invitation_id) {
    throw new Error('La fila seleccionada no tiene "ID Invitación".');
  }
  const response = supabaseUpsert([payload]);
  return {
    rowNumber: active.rowNumber,
    syncedRows: response.length,
  };
}

function listColumnMetadata() {
  return Object.keys(COLUMN_MAP).map(function (alias) {
    return {
      alias: alias,
      header: COLUMN_MAP[alias],
    };
  });
}
