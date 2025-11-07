const COLUMN_DICTIONARY_SHEET_NAME = 'Diccionario Invitaciones';
const COLUMN_DICTIONARY_HEADERS = [
  'Sección',
  'Encabezado en hoja',
  'Alias (código)',
  'Descripción',
  'Ejemplo (referencial)',
  'Tipo de dato',
  'Notas (editable)',
];

const COLUMN_DICTIONARY_DESCRIPTIONS = {
  invitation_id: 'Identificador único de la invitación dentro del pipeline de PMO.',
  cotizacion: 'Número o código de cotización asignado por el cliente o por PMO.',
  descripcion: 'Resumen del requerimiento o alcance de la invitación.',
  tipo_servicio: 'Clasificación del servicio solicitado (cotización, licitación, parada, etc.).',
  estado_cotizacion: 'Etapa interna de preparación de la cotización.',
  estado_propuesta: 'Situación comercial de la propuesta enviada al cliente.',
  estado_pipeline:
    'Estado dentro del pipeline comercial de EKA (ej. Prospecto, Propuesta enviada, Negociación, Cerrado). Úsalo para dar trazabilidad comercial.',
  fecha_registro: 'Fecha en la que se registró la invitación en la hoja.',
  cliente: 'Cliente o razón social que emite la invitación.',
  zona_trabajo: 'Zona, unidad o proyecto donde se realizará el servicio.',
  solicitante: 'Contacto del cliente que solicita la cotización.',
  correo_solicitante: 'Correo electrónico del solicitante.',
  telefono_solicitante: 'Número telefónico del solicitante.',
  responsable_tecnico: 'Responsable técnico designado por EKA para preparar la respuesta.',
  correo_responsable_tecnico: 'Correo del responsable técnico.',
  telefono_responsable_tecnico: 'Teléfono del responsable técnico.',
  responsable_economico: 'Responsable económico/comercial asignado al proceso.',
  correo_responsable_economico: 'Correo del responsable económico.',
  telefono_responsable_economico: 'Teléfono del responsable económico.',
  fecha_invitacion: 'Fecha oficial de recepción de la invitación.',
  fecha_confirmacion: 'Fecha en la que se confirmó la participación.',
  fecha_visita_tecnica: 'Fecha programada para ejecutar la visita técnica.',
  fecha_consultas: 'Fecha límite para enviar consultas al cliente.',
  fecha_abs_consultas: 'Fecha en la que el cliente absuelve las consultas.',
  fecha_presentacion: 'Fecha comprometida para presentar la propuesta.',
  monto_ofertado: 'Monto económico ofertado en la moneda original.',
  moneda: 'Moneda original asociada al monto ofertado.',
  moneda_normalizada_usd:
    'Moneda destino usada para normalizar el monto (por defecto USD) y facilitar la comparación entre procesos.',
  orden_de_compra: 'Número o referencia de la orden de compra recibida.',
  fecha_orden_compra: 'Fecha de emisión o recepción de la orden de compra.',
  link_carpeta_drive: 'Enlace a la carpeta de Google Drive con el expediente de la invitación.',
  link_archivo_enviado: 'Enlace al archivo enviado al cliente (propuesta u otros adjuntos).',
  fecha_envio_propuesta: 'Fecha en la que se envió la propuesta al cliente.',
  hora_envio_propuesta: 'Hora de envío de la propuesta.',
  dias_a_vencimiento: 'Días restantes para el vencimiento o fecha objetivo.',
  riesgo_d3: 'Marca si la invitación tiene riesgo D-3.',
  riesgo_d1: 'Marca si la invitación tiene riesgo D-1.',
  enviado_a_tiempo: 'Indica si la propuesta se envió dentro del plazo establecido.',
  tiempo_respuesta_dias_hab: 'Número de días hábiles invertidos en responder la invitación.',
  requiere_visita: 'Define si la invitación requiere visita técnica obligatoria.',
  visita_ejecutada: 'Indica si la visita técnica se ejecutó.',
  semana_iso: 'Semana ISO correspondiente al registro (formato YYYY-Www).',
  mes_anio: 'Período mes-año utilizado para análisis de KPIs.',
  moneda_normalizada_usd: 'Moneda destino usada para normalizar el monto (por defecto USD).',
  tipo_cambio: 'Tipo de cambio aplicado para obtener la equivalencia en USD.',
  monto_ofertado_usd: 'Monto ofertado convertido a dólares estadounidenses.',
  notas_kpi: 'Notas, hallazgos o acuerdos relevantes para seguimiento de KPIs.',
};

const COLUMN_DICTIONARY_EXAMPLES = {
  cotizacion: 'COT-245-2024',
  descripcion: 'Servicio de mantenimiento integral de faja transportadora',
  tipo_servicio: 'Cotización Formal',
  estado_cotizacion: 'Elab. Prop. Téc.',
  estado_propuesta: 'Enviada',
  estado_pipeline: 'Negociación',
  fecha_registro: '2024-09-12',
  cliente: 'Minera Andina S.A.',
  zona_trabajo: 'Planta concentradora - Guardia A',
  solicitante: 'María López',
  correo_solicitante: 'maria.lopez@minera.com',
  telefono_solicitante: '+51 999 123 456',
  responsable_tecnico: 'Juan Torres',
  correo_responsable_tecnico: 'juan.torres@eka.com',
  telefono_responsable_tecnico: '+51 998 765 432',
  responsable_economico: 'Carla Gómez',
  correo_responsable_economico: 'carla.gomez@eka.com',
  telefono_responsable_economico: '+51 997 555 333',
  fecha_invitacion: '2024-09-10',
  fecha_confirmacion: '2024-09-13',
  fecha_visita_tecnica: '2024-09-18',
  fecha_consultas: '2024-09-20',
  fecha_abs_consultas: '2024-09-23',
  fecha_presentacion: '2024-09-27',
  monto_ofertado: '150000',
  moneda: 'PEN',
  tipo_cambio: '0.27',
  monto_ofertado_usd: '40500',
  semana_iso: '2024-W38',
  mes_anio: '2024-09',
  notas_kpi: 'Se requiere seguimiento con compras el 25/09.',
  link_carpeta_drive: 'https://drive.google.com/…',
  link_archivo_enviado: 'https://drive.google.com/…',
  dias_a_vencimiento: '5',
  enviado_a_tiempo: 'Sí',
  tiempo_respuesta_dias_hab: '7',
  requiere_visita: 'Sí',
  visita_ejecutada: 'No',
  moneda_normalizada_usd: 'USD',
  monto_ofertado_usd: '40500',
};

function ensureColumnDictionarySheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(COLUMN_DICTIONARY_SHEET_NAME);
  const rows = buildColumnDictionaryRows();
  const header = COLUMN_DICTIONARY_HEADERS;
  const notesIndex = header.length;
  const existingNotes = {};

  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const aliasValues = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
      const noteValues = sheet.getRange(2, notesIndex, lastRow - 1, 1).getValues();
      for (let i = 0; i < aliasValues.length; i++) {
        const alias = (aliasValues[i][0] || '').toString().trim();
        if (alias) {
          existingNotes[alias] = noteValues[i][0] || '';
        }
      }
    }
    sheet.clearContents();
  } else {
    sheet = ss.insertSheet(COLUMN_DICTIONARY_SHEET_NAME);
  }

  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  if (rows.length) {
    const data = rows.map(function (row) {
      const alias = row[2];
      const note = existingNotes[alias] || '';
      const base = row.slice(0, header.length - 1);
      base.push(note);
      return base;
    });
    sheet.getRange(2, 1, data.length, header.length).setValues(data);
  }

  sheet.autoResizeColumns(1, header.length - 1);
  return sheet;
}

function buildColumnDictionaryRows() {
  const rows = [];
  FIELD_SECTIONS.forEach(function (section) {
    const sectionTitle = section.title || section.id;
    section.fields.forEach(function (field) {
      const description = COLUMN_DICTIONARY_DESCRIPTIONS[field.alias] || '';
      const example = COLUMN_DICTIONARY_EXAMPLES[field.alias] || '';
      rows.push([
        sectionTitle,
        field.header,
        field.alias,
        description,
        example,
        describeFieldValueType(field.valueType),
        '',
      ]);
    });
  });
  return rows;
}

function describeFieldValueType(valueType) {
  switch (valueType) {
    case 'date':
      return 'Fecha';
    case 'time':
      return 'Hora';
    case 'number':
      return 'Número';
    case 'boolean':
      return 'Sí/No';
    case 'textarea':
      return 'Texto largo';
    case 'select':
      return 'Lista (desplegable)';
    case 'link':
      return 'Hipervínculo';
    case 'text':
    default:
      return 'Texto';
  }
}

function openColumnDictionarySheet() {
  const sheet = ensureColumnDictionarySheet();
  SpreadsheetApp.getActive().setActiveSheet(sheet);
  return {
    sheetName: sheet.getName(),
  };
}
