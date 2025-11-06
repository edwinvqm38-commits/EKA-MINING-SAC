/***** CONFIG *****/
const SUPABASE_URL = 'https://<TU-PROYECTO>.supabase.co';
const SUPABASE_ANON_KEY = '<TU-ANON-KEY>'; // MVP. Luego migraremos a Edge Function + x-api-key.
const SUPABASE_TABLE = 'invitations';
const INVITATIONS_HEADER_ROW = 1;

/***** COLUMNAS *****/
const INVITATIONS_SHEET_NAME = typeof SHEET_NAME !== 'undefined' ? SHEET_NAME : 'Invitaciones'; // cambia si tu hoja tiene otro nombre
const REQUIRED_HEADERS = [
  'ID Invitación',
  'Cotización',
  'Descripción',
  'Tipo de Servicio',
  'Estado de Cotización',
  'Fecha de Registro',
  'Cliente',
  'Zona de Trabajo',
  'Solicitante',
  'Correo del Solicitante',
  'Teléfono del Solicitante',
  'Responsable Técnico',
  'Correo del Resp. Téc.',
  'Teléfono del Resp. Téc.',
  'Responsable Económico',
  'Correo del Resp. Eco.',
  'Teléfono del Resp. Eco.',
  'Fecha de Invitación',
  'Fecha de Confirmación',
  'Fecha de Visita Téc.',
  'Fecha de Consultas',
  'Fecha de Abs. Consultas',
  'Fecha de Presentación',
  'Monto ofertado',
  'Moneda',
  'Estado de Propuesta',
  'Orden de Compra',
  'Fecha de la Orden de Compra',
  'Link Carpeta Drive',
  'Link Archivo Enviado',
  'Fecha de Envío de Propuesta',
  'Hora de Envío de Propuesta',
  'Días a Vencimiento',
  'Riesgo D-3',
  'Riesgo D-1',
  'Enviado a Tiempo',
  'Tiempo de Respuesta (días hábiles)',
  'Requiere Visita Técnica',
  'Visita Ejecutada',
  'Semana ISO',
  'Mes-Año',
  'Moneda Normalizada (USD)',
  'Tipo de Cambio',
  'Monto Ofertado (USD)',
  'Estado Pipeline',
  'Notas KPI'
];

/***** UTIL *****/
function _sheet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(INVITATIONS_SHEET_NAME);
  if (!sheet) {
    throw new Error('No se encontró la hoja "' + INVITATIONS_SHEET_NAME + '".');
  }
  return sheet;
}

function _headers() {
  const sh = _sheet();
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) {
    return [];
  }
  return sh.getRange(INVITATIONS_HEADER_ROW, 1, 1, lastCol).getValues()[0].map(String);
}

function _colIndexByName(name) {
  const headers = _headers();
  const idx = headers.indexOf(name);
  if (idx === -1) throw new Error('No se encontró la columna: ' + name);
  return idx + 1;
}

function _val(row, colName) {
  const col = _colIndexByName(colName);
  return _sheet().getRange(row, col).getValue();
}

function _dateOrNull(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _numberOrNull(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function _boolOrNull(v) {
  if (v === true || v === false) return v;
  if (typeof v === 'string') {
    const s = v.trim().toLowerCase();
    if (['true', 'sí', 'si', 'yes', 'y', '1'].includes(s)) return true;
    if (['false', 'no', '0', 'n'].includes(s)) return false;
  }
  if (typeof v === 'number') return v !== 0;
  return null;
}

/***** MAPEADOR (Sheet → JSON Supabase) *****/
function buildInvitationPayload(row) {
  const get = function (name) {
    return _val(row, name);
  };

  return {
    invitation_id: get('ID Invitación') || null,
    cotizacion: get('Cotización') || null,
    descripcion: get('Descripción') || null,
    tipo_servicio: get('Tipo de Servicio') || null,
    estado_cotizacion: get('Estado de Cotización') || null,
    fecha_registro: _dateOrNull(get('Fecha de Registro')),
    cliente: get('Cliente') || null,
    zona_trabajo: get('Zona de Trabajo') || null,
    solicitante: get('Solicitante') || null,
    correo_solicitante: get('Correo del Solicitante') || null,
    telefono_solicitante: get('Teléfono del Solicitante') || null,
    responsable_tecnico: get('Responsable Técnico') || null,
    correo_responsable_tecnico: get('Correo del Resp. Téc.') || null,
    telefono_responsable_tecnico: get('Teléfono del Resp. Téc.') || null,
    responsable_economico: get('Responsable Económico') || null,
    correo_responsable_economico: get('Correo del Resp. Eco.') || null,
    telefono_responsable_economico: get('Teléfono del Resp. Eco.') || null,
    fecha_invitacion: _dateOrNull(get('Fecha de Invitación')),
    fecha_confirmacion: _dateOrNull(get('Fecha de Confirmación')),
    fecha_visita_tecnica: _dateOrNull(get('Fecha de Visita Téc.')),
    fecha_consultas: _dateOrNull(get('Fecha de Consultas')),
    fecha_abs_consultas: _dateOrNull(get('Fecha de Abs. Consultas')),
    fecha_presentacion: _dateOrNull(get('Fecha de Presentación')),
    monto_ofertado: _numberOrNull(get('Monto ofertado')),
    moneda: get('Moneda') || null,
    estado_propuesta: get('Estado de Propuesta') || null,
    orden_de_compra: get('Orden de Compra') || null,
    fecha_orden_compra: _dateOrNull(get('Fecha de la Orden de Compra')),
    link_carpeta_drive: get('Link Carpeta Drive') || null,
    link_archivo_enviado: get('Link Archivo Enviado') || null,
    fecha_envio_propuesta: _dateOrNull(get('Fecha de Envío de Propuesta')),
    hora_envio_propuesta: get('Hora de Envío de Propuesta') || null,
    dias_a_vencimiento: _numberOrNull(get('Días a Vencimiento')),
    riesgo_d3: _boolOrNull(get('Riesgo D-3')),
    riesgo_d1: _boolOrNull(get('Riesgo D-1')),
    enviado_a_tiempo: _boolOrNull(get('Enviado a Tiempo')),
    tiempo_respuesta_dias_hab: _numberOrNull(get('Tiempo de Respuesta (días hábiles)')),
    requiere_visita: _boolOrNull(get('Requiere Visita Técnica')),
    visita_ejecutada: _boolOrNull(get('Visita Ejecutada')),
    semana_iso: get('Semana ISO') || null,
    mes_anio: get('Mes-Año') || null,
    moneda_normalizada_usd: get('Moneda Normalizada (USD)') || null,
    tipo_cambio: _numberOrNull(get('Tipo de Cambio')),
    monto_ofertado_usd: _numberOrNull(get('Monto Ofertado (USD)')),
    estado_pipeline: get('Estado Pipeline') || null,
    notas_kpi: get('Notas KPI') || null,
  };
}

/***** POST/UPSERT a Supabase (REST) *****/
function supabaseUpsert(rowsPayload) {
  const url = SUPABASE_URL + '/rest/v1/' + SUPABASE_TABLE + '?on_conflict=invitation_id';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      apikey: SUPABASE_ANON_KEY,
      Authorization: 'Bearer ' + SUPABASE_ANON_KEY,
      Prefer: 'resolution=merge-duplicates',
    },
    payload: JSON.stringify(rowsPayload),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code >= 200 && code < 300) {
    const text = res.getContentText();
    return text ? JSON.parse(text) : rowsPayload;
  }
  throw new Error('Supabase error ' + code + ': ' + res.getContentText());
}

/***** ACCIONES *****/
function syncSelectedRowsToSupabase() {
  const sh = _sheet();
  const sel = sh.getActiveRange();
  const start = sel.getRow();
  const end = start + sel.getNumRows() - 1;

  const payload = [];
  for (let r = start; r <= end; r++) {
    if (r <= INVITATIONS_HEADER_ROW) continue;
    const idInv = _val(r, 'ID Invitación');
    if (!idInv) continue;
    payload.push(buildInvitationPayload(r));
  }
  if (payload.length === 0) {
    SpreadsheetApp.getUi().alert('No hay filas válidas seleccionadas (falta ID Invitación).');
    return;
  }
  const out = supabaseUpsert(payload);
  SpreadsheetApp.getUi().alert('Sincronizadas ' + out.length + ' fila(s) a Supabase.');
}

function syncAllToSupabase() {
  const sh = _sheet();
  const lastRow = sh.getLastRow();
  const payload = [];
  for (let r = INVITATIONS_HEADER_ROW + 1; r <= lastRow; r++) {
    const idInv = _val(r, 'ID Invitación');
    if (!idInv) continue;
    payload.push(buildInvitationPayload(r));
  }
  if (payload.length === 0) {
    SpreadsheetApp.getUi().alert('No hay datos con ID Invitación.');
    return;
  }
  const out = supabaseUpsert(payload);
  SpreadsheetApp.getUi().alert('Sincronizadas ' + out.length + ' fila(s) a Supabase.');
}

function ensureInvitationColumns() {
  const sh = _sheet();
  const headers = _headers();
  let appended = 0;
  REQUIRED_HEADERS.forEach(function (header) {
    if (headers.indexOf(header) === -1) {
      sh.getRange(INVITATIONS_HEADER_ROW, sh.getLastColumn() + 1).setValue(header);
      headers.push(header);
      appended++;
    }
  });
  const message = appended
    ? 'Se agregaron ' + appended + ' columna(s) nuevas.'
    : 'Todas las columnas requeridas ya existen.';
  SpreadsheetApp.getUi().alert(message);
}
