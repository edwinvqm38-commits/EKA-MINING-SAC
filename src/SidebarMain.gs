const SIDEBAR_TITLE = 'Gestión de Invitaciones';
const SHEET_NAME = 'Invitaciones';
const HEADER_ROW = 1;

const TIPO_SERVICIO_OPTIONS = [
  'Cotización Formal',
  'Licitación',
  'Parada de Planta',
  'Adicional',
  'Adenda',
];

const TIPO_SERVICIO_PALETTE = {
  'cotización formal': '#2563eb',
  licitación: '#7c3aed',
  'parada de planta': '#dc2626',
  adicional: '#f97316',
  adenda: '#059669',
};

const ESTADO_COTIZACION_OPTIONS = [
  'Rev. Bases',
  'Visita Técnica',
  'Consultas',
  'Abs. Consultas',
  'Elab. Prop. Téc.',
  'Elab. Prop. Eco.',
  'Finalizado',
  'Enviado',
  'No se Cotiza',
  'Observado',
];

const ESTADO_PROPUESTA_OPTIONS = [
  'En elaboración',
  'Enviada',
  'En evaluación',
  'Adjudicada',
  'No adjudicada',
  'Observada',
  'Cancelada',
];

const DYNAMIC_OPTION_ALIASES = [
  'tipo_servicio',
  'estado_cotizacion',
  'estado_propuesta',
  'cliente',
  'zona_trabajo',
  'solicitante',
  'responsable_tecnico',
  'responsable_economico',
];

const FIELD_SECTIONS = [
  {
    id: 'general',
    title: 'Datos generales',
    description: 'Información básica del proceso de cotización.',
    fields: [
      {
        alias: 'invitation_id',
        header: 'ID Invitación',
        label: 'ID Invitación',
        valueType: 'text',
        width: 6,
        placeholder: 'INV-0001',
        hidden: true,
      },
      {
        alias: 'cotizacion',
        header: 'Cotización',
        label: 'Cotización',
        valueType: 'text',
        width: 6,
      },
      {
        alias: 'descripcion',
        header: 'Descripción',
        label: 'Descripción',
        valueType: 'textarea',
        width: 12,
        rows: 3,
      },
      {
        alias: 'tipo_servicio',
        header: 'Tipo de Servicio',
        label: 'Tipo de Servicio',
        valueType: 'select',
        width: 6,
        options: TIPO_SERVICIO_OPTIONS,
        allowCustom: true,
        palette: TIPO_SERVICIO_PALETTE,
      },
      {
        alias: 'estado_cotizacion',
        header: 'Estado de Cotización',
        label: 'Estado de Cotización',
        valueType: 'select',
        width: 6,
        options: ESTADO_COTIZACION_OPTIONS,
        allowCustom: true,
      },
      {
        alias: 'estado_propuesta',
        header: 'Estado de Propuesta',
        label: 'Estado de Propuesta',
        valueType: 'select',
        width: 6,
        options: ESTADO_PROPUESTA_OPTIONS,
        allowCustom: true,
      },
      {
        alias: 'estado_pipeline',
        header: 'Estado Pipeline',
        label: 'Estado Pipeline',
        valueType: 'text',
        width: 6,
        helper: 'Consulta el "Diccionario Invitaciones" para las definiciones de pipeline.',
      },
      {
        alias: 'fecha_registro',
        header: 'Fecha de Registro',
        label: 'Fecha de Registro',
        valueType: 'date',
        width: 6,
      },
      {
        alias: 'cliente',
        header: 'Cliente',
        label: 'Cliente',
        valueType: 'select',
        width: 6,
        optionsSource: 'cliente',
        allowCustom: true,
      },
      {
        alias: 'zona_trabajo',
        header: 'Zona de Trabajo',
        label: 'Zona de Trabajo',
        valueType: 'select',
        width: 6,
        optionsSource: 'zona_trabajo',
        allowCustom: true,
      },
    ],
  },
  {
    id: 'contactos',
    title: 'Contactos clave',
    description: 'Personas involucradas en la invitación.',
    fields: [
      {
        alias: 'solicitante',
        header: 'Solicitante',
        label: 'Solicitante',
        valueType: 'select',
        width: 12,
        columnClass: 'col-12 col-lg-4',
        optionsSource: 'solicitante',
        allowCustom: true,
        helper: 'Selecciona un solicitante existente o agrega uno nuevo.',
        group: 'solicitante',
        groupLabel: 'Solicitante',
        groupDescription: 'Contacto principal que gestiona la invitación.',
      },
      {
        alias: 'correo_solicitante',
        header: 'Correo del Solicitante',
        label: 'Correo del Solicitante',
        valueType: 'text',
        width: 12,
        columnClass: 'col-12 col-lg-4',
        placeholder: 'correo@empresa.com',
        group: 'solicitante',
      },
      {
        alias: 'telefono_solicitante',
        header: 'Teléfono del Solicitante',
        label: 'Teléfono del Solicitante',
        valueType: 'text',
        width: 12,
        columnClass: 'col-12 col-lg-4',
        group: 'solicitante',
      },
      {
        alias: 'responsable_tecnico',
        header: 'Responsable Técnico',
        label: 'Responsable Técnico',
        valueType: 'select',
        width: 12,
        columnClass: 'col-12 col-lg-4',
        group: 'responsable_tecnico',
        groupLabel: 'Responsable técnico',
        groupDescription: 'Profesional a cargo de la coordinación técnica.',
        optionsSource: 'responsable_tecnico',
        allowCustom: true,
      },
      {
        alias: 'correo_responsable_tecnico',
        header: 'Correo del Resp. Téc.',
        label: 'Correo del Resp. Téc.',
        valueType: 'text',
        width: 12,
        columnClass: 'col-12 col-lg-4',
        group: 'responsable_tecnico',
      },
      {
        alias: 'telefono_responsable_tecnico',
        header: 'Teléfono del Resp. Téc.',
        label: 'Teléfono del Resp. Téc.',
        valueType: 'text',
        width: 12,
        columnClass: 'col-12 col-lg-4',
        group: 'responsable_tecnico',
      },
      {
        alias: 'responsable_economico',
        header: 'Responsable Económico',
        label: 'Responsable Económico',
        valueType: 'select',
        width: 12,
        columnClass: 'col-12 col-lg-4',
        group: 'responsable_economico',
        groupLabel: 'Responsable económico',
        groupDescription: 'Contacto financiero y comercial del proceso.',
        optionsSource: 'responsable_economico',
        allowCustom: true,
      },
      {
        alias: 'correo_responsable_economico',
        header: 'Correo del Resp. Eco.',
        label: 'Correo del Resp. Eco.',
        valueType: 'text',
        width: 12,
        columnClass: 'col-12 col-lg-4',
        group: 'responsable_economico',
      },
      {
        alias: 'telefono_responsable_economico',
        header: 'Teléfono del Resp. Eco.',
        label: 'Teléfono del Resp. Eco.',
        valueType: 'text',
        width: 12,
        columnClass: 'col-12 col-lg-4',
        group: 'responsable_economico',
      },
    ],
  },
  {
    id: 'cronograma',
    title: 'Cronograma y entregables',
    description: 'Fechas clave, entregables y documentación de soporte.',
    fields: [
      {
        alias: 'fecha_invitacion',
        header: 'Fecha de Invitación',
        label: 'Fecha de Invitación',
        valueType: 'date',
        width: 6,
      },
      {
        alias: 'fecha_confirmacion',
        header: 'Fecha de Confirmación',
        label: 'Fecha de Confirmación',
        valueType: 'date',
        width: 6,
      },
      {
        alias: 'fecha_visita_tecnica',
        header: 'Fecha de Visita Téc.',
        label: 'Fecha de Visita Téc.',
        valueType: 'date',
        width: 6,
      },
      {
        alias: 'fecha_consultas',
        header: 'Fecha de Consultas',
        label: 'Fecha de Consultas',
        valueType: 'date',
        width: 6,
      },
      {
        alias: 'fecha_abs_consultas',
        header: 'Fecha de Abs. Consultas',
        label: 'Fecha de Abs. Consultas',
        valueType: 'date',
        width: 6,
      },
      {
        alias: 'fecha_presentacion',
        header: 'Fecha de Presentación',
        label: 'Fecha de Presentación',
        valueType: 'date',
        width: 6,
      },
      {
        alias: 'fecha_envio_propuesta',
        header: 'Fecha de Envío de Propuesta',
        label: 'Fecha de Envío de Propuesta',
        valueType: 'date',
        width: 6,
      },
      {
        alias: 'hora_envio_propuesta',
        header: 'Hora de Envío de Propuesta',
        label: 'Hora de Envío de Propuesta',
        valueType: 'time',
        width: 6,
      },
      {
        alias: 'orden_de_compra',
        header: 'Orden de Compra',
        label: 'Orden de Compra',
        valueType: 'text',
        width: 6,
      },
      {
        alias: 'fecha_orden_compra',
        header: 'Fecha de la Orden de Compra',
        label: 'Fecha de la Orden de Compra',
        valueType: 'date',
        width: 6,
      },
      {
        alias: 'link_carpeta_drive',
        header: 'Link Carpeta Drive',
        label: 'Link Carpeta Drive',
        valueType: 'link',
        width: 12,
        helper: 'Guarda la carpeta como hipervínculo para acceder rápidamente.',
      },
      {
        alias: 'link_archivo_enviado',
        header: 'Link Archivo Enviado',
        label: 'Link Archivo Enviado',
        valueType: 'link',
        width: 12,
        helper: 'Adjunta el enlace del archivo enviado al cliente.',
      },
    ],
  },
  {
    id: 'kpi',
    title: 'KPIs y métricas',
    description: 'Seguimiento de indicadores clave y normalizaciones.',
    fields: [
      {
        alias: 'monto_ofertado',
        header: 'Monto ofertado',
        label: 'Monto ofertado',
        valueType: 'number',
        width: 6,
      },
      {
        alias: 'moneda',
        header: 'Moneda',
        label: 'Moneda',
        valueType: 'select',
        width: 6,
        options: [
          { value: 'PEN', label: 'Sol (PEN)' },
          { value: 'USD', label: 'Dólar (USD)' },
        ],
        allowCustom: true,
        helper: 'Selecciona la moneda de origen o agrega una nueva si es necesario.',
      },
      {
        alias: 'moneda_normalizada_usd',
        header: 'Moneda Normalizada (USD)',
        label: 'Moneda Normalizada (USD)',
        valueType: 'text',
        width: 6,
        readOnly: true,
        helper: 'Se actualiza automáticamente según la moneda seleccionada.',
      },
      {
        alias: 'tipo_cambio',
        header: 'Tipo de Cambio',
        label: 'Tipo de Cambio',
        valueType: 'number',
        width: 6,
      },
      {
        alias: 'monto_ofertado_usd',
        header: 'Monto Ofertado (USD)',
        label: 'Monto Ofertado (USD)',
        valueType: 'number',
        width: 6,
        readOnly: true,
        helper: 'Calculado automáticamente según moneda, monto y tipo de cambio.',
      },
      {
        alias: 'dias_a_vencimiento',
        header: 'Días a Vencimiento',
        label: 'Días a Vencimiento',
        valueType: 'number',
        width: 6,
      },
      {
        alias: 'riesgo_d3',
        header: 'Riesgo D-3',
        label: 'Riesgo D-3',
        valueType: 'boolean',
        width: 6,
      },
      {
        alias: 'riesgo_d1',
        header: 'Riesgo D-1',
        label: 'Riesgo D-1',
        valueType: 'boolean',
        width: 6,
      },
      {
        alias: 'enviado_a_tiempo',
        header: 'Enviado a Tiempo',
        label: 'Enviado a Tiempo',
        valueType: 'boolean',
        width: 6,
      },
      {
        alias: 'tiempo_respuesta_dias_hab',
        header: 'Tiempo de Respuesta (días hábiles)',
        label: 'Tiempo de Respuesta (días hábiles)',
        valueType: 'number',
        width: 6,
      },
      {
        alias: 'requiere_visita',
        header: 'Requiere Visita Técnica',
        label: 'Requiere Visita Técnica',
        valueType: 'boolean',
        width: 6,
      },
      {
        alias: 'visita_ejecutada',
        header: 'Visita Ejecutada',
        label: 'Visita Ejecutada',
        valueType: 'boolean',
        width: 6,
      },
      {
        alias: 'semana_iso',
        header: 'Semana ISO',
        label: 'Semana ISO',
        valueType: 'text',
        width: 6,
      },
      {
        alias: 'mes_anio',
        header: 'Mes-Año',
        label: 'Mes-Año',
        valueType: 'text',
        width: 6,
      },
      {
        alias: 'notas_kpi',
        header: 'Notas KPI',
        label: 'Notas KPI',
        valueType: 'textarea',
        width: 12,
        rows: 3,
      },
    ],
  },
];

const COLUMN_MAP = (function () {
  const map = {};
  FIELD_SECTIONS.forEach(function (section) {
    section.fields.forEach(function (field) {
      map[field.alias] = field.header;
    });
  });
  return map;
})();

const FIELD_LOOKUP = (function () {
  const lookup = {};
  FIELD_SECTIONS.forEach(function (section) {
    section.fields.forEach(function (field) {
      lookup[field.alias] = field;
    });
  });
  return lookup;
})();

function cloneFieldDefinition(field) {
  if (!field) {
    return field;
  }
  const copy = {};
  Object.keys(field).forEach(function (key) {
    const value = field[key];
    if (Array.isArray(value)) {
      copy[key] = value.slice();
    } else if (value && typeof value === 'object') {
      copy[key] = Object.assign({}, value);
    } else {
      copy[key] = value;
    }
  });
  return copy;
}

function cloneFieldSections() {
  return FIELD_SECTIONS.map(function (section) {
    const copy = Object.assign({}, section);
    copy.fields = section.fields.map(function (field) {
      return cloneFieldDefinition(field);
    });
    return copy;
  });
}

function getInvitationSheetConfig() {
  return {
    sheetName: SHEET_NAME,
    fieldSections: cloneFieldSections(),
  };
}

function buildFieldSections(config) {
  if (config && Array.isArray(config.fieldSections) && config.fieldSections.length) {
    return config.fieldSections;
  }
  return cloneFieldSections();
}

function buildFieldSectionsForVisibility(columnVisibility) {
  const clonedSections = cloneFieldSections();
  return clonedSections
    .map(function (section) {
      section.fields = section.fields.filter(function (field) {
        if (!field || !field.alias) {
          return false;
        }
        if (field.alias === 'invitation_id') {
          return true;
        }
        if (!columnVisibility) {
          return true;
        }
        return columnVisibility[field.alias] !== false;
      });
      return section;
    })
    .filter(function (section) {
      return Array.isArray(section.fields) && section.fields.length > 0;
    });
}

function filterOptionMapByVisibility(optionMap, columnVisibility) {
  if (!columnVisibility) {
    return optionMap;
  }
  const filtered = {};
  Object.keys(optionMap || {}).forEach(function (alias) {
    if (columnVisibility[alias] === false) {
      return;
    }
    filtered[alias] = optionMap[alias];
  });
  return filtered;
}

function filterContactDirectoryByVisibility(directory, columnVisibility) {
  if (!columnVisibility) {
    return directory;
  }
  const filtered = {};
  Object.keys(directory || {}).forEach(function (alias) {
    if (columnVisibility[alias] === false) {
      return;
    }
    filtered[alias] = directory[alias];
  });
  return filtered;
}

function filterUpdatesForColumnVisibility(updates, columnVisibility) {
  if (!columnVisibility) {
    return Object.assign({}, updates);
  }
  const filtered = {};
  Object.keys(updates || {}).forEach(function (alias) {
    if (alias !== 'invitation_id' && columnVisibility[alias] === false) {
      return;
    }
    filtered[alias] = updates[alias];
  });
  return filtered;
}

const BOOLEAN_TRUE_VALUES = ['true', 'sí', 'si', 'yes', 'y', '1', 'verdadero'];
const BOOLEAN_FALSE_VALUES = ['false', 'no', '0', 'n', 'falso'];

function onOpen() {
  ensurePermissionsSheet();
  ensureColumnPermissionsSheet();
  ensureColumnDictionarySheet();
  SpreadsheetApp.getUi()
    .createMenu('PMO • EKA')
    .addItem('Abrir panel de invitaciones', 'showSidebar')
    .addItem('Asegurar columnas requeridas', 'ensureInvitationColumns')
    .addToUi();
}

function showSidebar() {
  ensurePermissionsSheet();
  ensureColumnPermissionsSheet();
  ensureColumnDictionarySheet();

  const config = getInvitationSheetConfig();
  const sections = buildFieldSections(config);

  const template = HtmlService.createTemplateFromFile('Sidebar');
  template.sheetName = config.sheetName;
  template.fieldSectionsJson = JSON.stringify(sections);
  template.dictionarySheetName = COLUMN_DICTIONARY_SHEET_NAME;
  template.permissionsSheetName = PERMISSIONS_SHEET_NAME;
  template.columnPermissionSheetName = COLUMN_PERMISSIONS_SHEET_NAME;

  const htmlOutput = template
    .evaluate()
    .setTitle(SIDEBAR_TITLE)
    .setWidth(420);
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

function ensureInvitationIdForRow(rowNumber) {
  if (!rowNumber || rowNumber <= HEADER_ROW) {
    return '';
  }
  const headers = getHeaders();
  const headerName = COLUMN_MAP.invitation_id;
  const columnIndex = headerName ? headers.indexOf(headerName) : -1;
  if (columnIndex === -1) {
    return '';
  }
  const sheet = getInvitationSheet();
  const cell = sheet.getRange(rowNumber, columnIndex + 1);
  const currentValue = String(cell.getDisplayValue() || '').trim();
  if (currentValue) {
    return currentValue;
  }
  const generatedId = generateInvitationId(sheet);
  cell.setValue(generatedId);
  return generatedId;
}

function generateInvitationId(sheet) {
  const tz = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const props = PropertiesService.getDocumentProperties();
  let counter = Number(props.getProperty('INVITATION_SEQUENCE') || '0');
  let candidate;
  const existingIds = new Set();
  if (sheet) {
    const headers = getHeaders();
    const headerName = COLUMN_MAP.invitation_id;
    const columnIndex = headerName ? headers.indexOf(headerName) : -1;
    if (columnIndex !== -1) {
      const lastRow = sheet.getLastRow();
      if (lastRow > HEADER_ROW) {
        const range = sheet.getRange(HEADER_ROW + 1, columnIndex + 1, lastRow - HEADER_ROW, 1);
        const values = range.getDisplayValues();
        values.forEach(function (row) {
          const value = (row[0] || '').toString().trim();
          if (value) {
            existingIds.add(value);
          }
        });
      }
    }
  }
  do {
    counter += 1;
    candidate = 'INV-' + timestamp + '-' + String(counter).padStart(4, '0');
  } while (existingIds.has(candidate));
  props.setProperty('INVITATION_SEQUENCE', String(counter));
  return candidate;
}

function getRowData(rowNumber, columnVisibility) {
  if (rowNumber <= HEADER_ROW) {
    throw new Error('La fila activa debe estar debajo del encabezado.');
  }
  ensureInvitationIdForRow(rowNumber);
  const sheet = getInvitationSheet();
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    return { rowNumber: rowNumber };
  }
  const headers = getHeaders();
  const range = sheet.getRange(rowNumber, 1, 1, lastColumn);
  const values = range.getValues()[0];
  const displayValues = range.getDisplayValues()[0];
  const formulas = range.getFormulas()[0];
  const richTextValues = range.getRichTextValues()[0];
  const result = { rowNumber: rowNumber };

  Object.keys(COLUMN_MAP).forEach(function (alias) {
    if (alias !== 'invitation_id' && columnVisibility && columnVisibility[alias] === false) {
      return;
    }
    const headerName = COLUMN_MAP[alias];
    const columnIndex = headers.indexOf(headerName);
    if (columnIndex === -1) {
      return;
    }
    const def = FIELD_LOOKUP[alias];
    const rawValue = values[columnIndex];
    const displayValue = displayValues[columnIndex];
    const formula = formulas[columnIndex];
    const richText = richTextValues[columnIndex];
    result[alias] = formatCellForClient(def, rawValue, displayValue, formula, richText);
  });

  return result;
}

function getActiveInvitation(columnVisibility) {
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
    data: getRowData(row, columnVisibility),
  };
}

function buildDropdownOptions() {
  const options = {};
  DYNAMIC_OPTION_ALIASES.forEach(function (alias) {
    options[alias] = collectDistinctColumnValues(alias);
  });
  return options;
}

function collectDistinctColumnValues(alias) {
  const headerName = COLUMN_MAP[alias];
  if (!headerName) {
    return [];
  }
  const sheet = getInvitationSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) {
    return [];
  }
  const headers = getHeaders();
  const columnIndex = headers.indexOf(headerName);
  if (columnIndex === -1) {
    return [];
  }
  const rowCount = lastRow - HEADER_ROW;
  const range = sheet.getRange(HEADER_ROW + 1, columnIndex + 1, rowCount, 1);
  const values = range
    .getDisplayValues()
    .map(function (row) {
      const value = row[0];
      return value === null || value === undefined ? '' : String(value).trim();
    })
    .filter(function (value) {
      return value !== '';
    });
  const seen = {};
  const unique = [];
  values.forEach(function (value) {
    const key = value.toLowerCase();
    if (!seen[key]) {
      seen[key] = true;
      unique.push(value);
    }
  });
  unique.sort(function (a, b) {
    const aKey = a.toLowerCase();
    const bKey = b.toLowerCase();
    if (aKey < bKey) {
      return -1;
    }
    if (aKey > bKey) {
      return 1;
    }
    return 0;
  });
  return unique;
}

function buildContactDirectory() {
  const directory = {
    solicitante: {},
    responsable_tecnico: {},
    responsable_economico: {},
  };
  const sheet = getInvitationSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) {
    return directory;
  }
  const headers = getHeaders();
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    return directory;
  }
  const range = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastColumn);
  const values = range.getValues();
  const solicitanteIndex = headers.indexOf(COLUMN_MAP.solicitante);
  const solicitanteEmailIndex = headers.indexOf(COLUMN_MAP.correo_solicitante);
  const solicitantePhoneIndex = headers.indexOf(COLUMN_MAP.telefono_solicitante);
  const tecnicoIndex = headers.indexOf(COLUMN_MAP.responsable_tecnico);
  const tecnicoEmailIndex = headers.indexOf(COLUMN_MAP.correo_responsable_tecnico);
  const tecnicoPhoneIndex = headers.indexOf(COLUMN_MAP.telefono_responsable_tecnico);
  const economicoIndex = headers.indexOf(COLUMN_MAP.responsable_economico);
  const economicoEmailIndex = headers.indexOf(COLUMN_MAP.correo_responsable_economico);
  const economicoPhoneIndex = headers.indexOf(COLUMN_MAP.telefono_responsable_economico);

  values.forEach(function (row) {
    if (solicitanteIndex !== -1) {
      const rawName = row[solicitanteIndex];
      const name = rawName === null || rawName === undefined ? '' : String(rawName).trim();
      if (name) {
        const key = name.toLowerCase();
        if (!directory.solicitante[key]) {
          directory.solicitante[key] = {
            name: name,
            email:
              solicitanteEmailIndex !== -1 && row[solicitanteEmailIndex]
                ? String(row[solicitanteEmailIndex]).trim()
                : '',
            phone:
              solicitantePhoneIndex !== -1 && row[solicitantePhoneIndex]
                ? String(row[solicitantePhoneIndex]).trim()
                : '',
          };
        }
      }
    }

    if (tecnicoIndex !== -1) {
      const rawTecnico = row[tecnicoIndex];
      const tecnico = rawTecnico === null || rawTecnico === undefined ? '' : String(rawTecnico).trim();
      if (tecnico) {
        const key = tecnico.toLowerCase();
        if (!directory.responsable_tecnico[key]) {
          directory.responsable_tecnico[key] = {
            name: tecnico,
            email:
              tecnicoEmailIndex !== -1 && row[tecnicoEmailIndex]
                ? String(row[tecnicoEmailIndex]).trim()
                : '',
            phone:
              tecnicoPhoneIndex !== -1 && row[tecnicoPhoneIndex]
                ? String(row[tecnicoPhoneIndex]).trim()
                : '',
          };
        }
      }
    }

    if (economicoIndex !== -1) {
      const rawEconomico = row[economicoIndex];
      const economico = rawEconomico === null || rawEconomico === undefined ? '' : String(rawEconomico).trim();
      if (economico) {
        const key = economico.toLowerCase();
        if (!directory.responsable_economico[key]) {
          directory.responsable_economico[key] = {
            name: economico,
            email:
              economicoEmailIndex !== -1 && row[economicoEmailIndex]
                ? String(row[economicoEmailIndex]).trim()
                : '',
            phone:
              economicoPhoneIndex !== -1 && row[economicoPhoneIndex]
                ? String(row[economicoPhoneIndex]).trim()
                : '',
          };
        }
      }
    }
  });

  return directory;
}

function removeDropdownValue(request) {
  assertUserCanEditInvitations();
  const alias = request && request.alias;
  const value = request && request.value;
  if (!alias || !value) {
    throw new Error('Especifica la lista y el valor que deseas eliminar.');
  }
  if (!COLUMN_MAP[alias]) {
    throw new Error('No se reconoce el alias de columna: ' + alias);
  }
  const headers = getHeaders();
  const headerName = COLUMN_MAP[alias];
  const columnIndex = headers.indexOf(headerName);
  if (columnIndex === -1) {
    throw new Error('No se encontró la columna "' + headerName + '" en la hoja.');
  }
  const sheet = getInvitationSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) {
    return { removed: 0 };
  }
  const range = sheet.getRange(HEADER_ROW + 1, columnIndex + 1, lastRow - HEADER_ROW, 1);
  const values = range.getValues();
  const normalizedTarget = String(value).trim().toLowerCase();
  let removed = 0;
  const rowsToClear = [];
  const updatedValues = values.map(function (row, index) {
    const rawValue = row[0];
    const normalized = rawValue === null || rawValue === undefined ? '' : String(rawValue).trim().toLowerCase();
    if (normalized === normalizedTarget) {
      removed++;
      rowsToClear.push(index);
      return [''];
    }
    return [rawValue];
  });
  if (removed === 0) {
    return { removed: 0 };
  }
  range.setValues(updatedValues);

  clearAssociatedContacts(alias, rowsToClear);

  return { removed: removed };
}

function clearAssociatedContacts(alias, rowIndexes) {
  if (!rowIndexes || !rowIndexes.length) {
    return;
  }
  let relatedAliases = [];
  if (alias === 'solicitante') {
    relatedAliases = ['correo_solicitante', 'telefono_solicitante'];
  } else if (alias === 'responsable_tecnico') {
    relatedAliases = ['correo_responsable_tecnico', 'telefono_responsable_tecnico'];
  } else if (alias === 'responsable_economico') {
    relatedAliases = ['correo_responsable_economico', 'telefono_responsable_economico'];
  }
  if (!relatedAliases.length) {
    return;
  }
  const headers = getHeaders();
  const sheet = getInvitationSheet();
  rowIndexes.forEach(function (index) {
    const rowNumber = HEADER_ROW + 1 + index;
    relatedAliases.forEach(function (relatedAlias) {
      const headerName = COLUMN_MAP[relatedAlias];
      if (!headerName) {
        return;
      }
      const columnIndex = headers.indexOf(headerName);
      if (columnIndex === -1) {
        return;
      }
      sheet.getRange(rowNumber, columnIndex + 1).setValue('');
    });
  });
}

function applyKpiCalculations(rowNumber, updates) {
  const currencyRaw = getPendingValue(rowNumber, 'moneda', updates);
  const amountRaw = getPendingValue(rowNumber, 'monto_ofertado', updates);
  let tipoCambioRaw = getPendingValue(rowNumber, 'tipo_cambio', updates);
  const currency = currencyRaw === null || currencyRaw === undefined ? '' : String(currencyRaw).trim();
  const amount = toNumberOrNull(amountRaw);
  let tipoCambio = toNumberOrNull(tipoCambioRaw);
  let normalizedCurrency = '';
  let amountUsd = null;

  if (!currency) {
    updates.moneda_normalizada_usd = '';
    updates.monto_ofertado_usd = '';
    return;
  }

  const upperCurrency = currency.toUpperCase();
  const isUsd = ['USD', 'US$', 'DOLAR', 'DÓLAR', 'DOLARES', 'DÓLARES'].indexOf(upperCurrency) !== -1;
  const isPen = ['PEN', 'SOL', 'SOLES', 'S/.', 'S/'].indexOf(upperCurrency) !== -1;

  if (isUsd) {
    normalizedCurrency = 'USD';
    if (amount !== null) {
      amountUsd = amount;
    }
    if (tipoCambio === null || tipoCambio === 0) {
      tipoCambio = 1;
      updates.tipo_cambio = 1;
    }
  } else if (isPen) {
    normalizedCurrency = 'USD';
    if (amount !== null && tipoCambio !== null && tipoCambio !== 0) {
      amountUsd = amount * tipoCambio;
    }
  } else {
    normalizedCurrency = currency;
    if (amount !== null) {
      amountUsd = amount;
    }
  }

  updates.moneda_normalizada_usd = normalizedCurrency || '';
  if (amountUsd === null || amountUsd === undefined || amountUsd === '') {
    updates.monto_ofertado_usd = '';
  } else {
    updates.monto_ofertado_usd = amountUsd;
  }
}

function getPendingValue(rowNumber, alias, updates) {
  if (updates && Object.prototype.hasOwnProperty.call(updates, alias)) {
    return updates[alias];
  }
  return getCellValue(rowNumber, alias);
}

function getCellValue(rowNumber, alias) {
  if (!COLUMN_MAP[alias]) {
    return '';
  }
  const headers = getHeaders();
  const headerName = COLUMN_MAP[alias];
  const columnIndex = headers.indexOf(headerName);
  if (columnIndex === -1) {
    return '';
  }
  const sheet = getInvitationSheet();
  return sheet.getRange(rowNumber, columnIndex + 1).getValue();
}

function toNumberOrNull(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  if (typeof value === 'number' && !isNaN(value)) {
    return value;
  }
  const parsed = Number(value);
  return isNaN(parsed) ? null : parsed;
}

function updateInvitationRow(rowNumber, updates, columnVisibility) {
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
  return getRowData(rowNumber, columnVisibility);
}

function normalizeValueForSheet(alias, value) {
  const definition = FIELD_LOOKUP[alias];
  if (!definition) {
    return value === undefined || value === null ? '' : value;
  }
  if (value === undefined || value === null || value === '') {
    return '';
  }

  switch (definition.valueType) {
    case 'date':
      return coerceDateValue(value) || '';
    case 'time':
      return typeof value === 'string' ? value : '';
    case 'number':
      return coerceNumberValue(value);
    case 'boolean':
      return coerceBooleanValue(value);
    case 'link':
      return coerceLinkValue(value);
    case 'textarea':
    case 'text':
    default:
      return value;
  }
}

function coerceDateValue(value) {
  if (!value) {
    return null;
  }
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return isNaN(value.getTime()) ? null : value;
  }
  const parsed = new Date(value);
  if (isNaN(parsed.getTime())) {
    return null;
  }
  return parsed;
}

function coerceNumberValue(value) {
  if (value === '' || value === null || value === undefined) {
    return '';
  }
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : '';
}

function coerceBooleanValue(value) {
  if (value === true || value === false) {
    return value;
  }
  if (typeof value === 'number') {
    return value !== 0;
  }
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (BOOLEAN_TRUE_VALUES.indexOf(normalized) !== -1) {
      return true;
    }
    if (BOOLEAN_FALSE_VALUES.indexOf(normalized) !== -1) {
      return false;
    }
  }
  return '';
}

function coerceLinkValue(value) {
  if (typeof value === 'object' && value !== null) {
    const url = value.url ? String(value.url).trim() : '';
    const text = value.text ? String(value.text).trim() : '';
    if (!url && !text) {
      return '';
    }
    if (!url) {
      return text;
    }
    const label = text || url;
    return buildHyperlinkFormula(url, label);
  }
  if (typeof value === 'string') {
    return value;
  }
  return '';
}

function formatCellForClient(definition, rawValue, displayValue, formula, richText) {
  if (!definition) {
    return rawValue;
  }

  switch (definition.valueType) {
    case 'date':
      return formatDateValue(rawValue, displayValue);
    case 'time':
      return formatTimeValue(rawValue, displayValue);
    case 'number':
      return formatNumberValue(rawValue, displayValue);
    case 'boolean':
      return formatBooleanValue(rawValue);
    case 'link':
      return extractHyperlinkDescriptor(rawValue, displayValue, formula, richText);
    case 'textarea':
    case 'text':
    default:
      return rawValue === '' || rawValue === null || rawValue === undefined
        ? ''
        : String(rawValue);
  }
}

function formatDateValue(rawValue, displayValue) {
  const value = rawValue || displayValue;
  if (!value) {
    return '';
  }
  if (Object.prototype.toString.call(rawValue) === '[object Date]') {
    return Utilities.formatDate(rawValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const parsed = new Date(value);
  if (isNaN(parsed.getTime())) {
    return '';
  }
  return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatTimeValue(rawValue, displayValue) {
  if (!rawValue && !displayValue) {
    return '';
  }
  if (Object.prototype.toString.call(rawValue) === '[object Date]') {
    return Utilities.formatDate(rawValue, Session.getScriptTimeZone(), 'HH:mm');
  }
  if (typeof rawValue === 'string' && rawValue.match(/^\d{1,2}:\d{2}$/)) {
    return rawValue;
  }
  if (typeof displayValue === 'string' && displayValue.match(/^\d{1,2}:\d{2}$/)) {
    return displayValue;
  }
  return '';
}

function formatNumberValue(rawValue, displayValue) {
  if (rawValue === '' || rawValue === null || rawValue === undefined) {
    return '';
  }
  if (typeof rawValue === 'number') {
    return String(rawValue);
  }
  if (typeof rawValue === 'string' && rawValue.trim() !== '') {
    return rawValue;
  }
  if (typeof displayValue === 'string' && displayValue.trim() !== '') {
    return displayValue;
  }
  return '';
}

function formatBooleanValue(rawValue) {
  if (rawValue === true || rawValue === false) {
    return rawValue;
  }
  if (rawValue === '' || rawValue === null || rawValue === undefined) {
    return null;
  }
  if (typeof rawValue === 'number') {
    return rawValue !== 0;
  }
  if (typeof rawValue === 'string') {
    const normalized = rawValue.trim().toLowerCase();
    if (BOOLEAN_TRUE_VALUES.indexOf(normalized) !== -1) {
      return true;
    }
    if (BOOLEAN_FALSE_VALUES.indexOf(normalized) !== -1) {
      return false;
    }
  }
  return null;
}

function extractHyperlinkDescriptor(rawValue, displayValue, formula, richText) {
  let text = '';
  let url = '';

  if (formula) {
    const parsed = parseHyperlinkFormula(formula);
    if (parsed) {
      url = parsed.url || '';
      text = parsed.text || parsed.url || displayValue || rawValue || '';
      return {
        url: url,
        text: text,
        formula: formula,
      };
    }
  }

  const richUrl = extractUrlFromRichText(richText);
  if (richUrl) {
    url = richUrl;
    text = displayValue || rawValue || richUrl;
    return {
      url: url,
      text: text || richUrl,
      formula: buildHyperlinkFormula(url, text || richUrl),
    };
  }

  const candidate = displayValue || rawValue;
  if (candidate && typeof candidate === 'string' && candidate.trim().startsWith('http')) {
    url = candidate.trim();
    text = candidate.trim();
  } else {
    text = candidate ? String(candidate) : '';
  }

  return {
    url: url,
    text: text,
    formula: url ? buildHyperlinkFormula(url, text || url) : formula || '',
  };
}

function extractUrlFromRichText(richText) {
  if (!richText) {
    return '';
  }
  try {
    if (typeof richText.getLinkUrl === 'function') {
      const direct = richText.getLinkUrl();
      if (direct) {
        return direct;
      }
    }
    if (typeof richText.getRuns === 'function') {
      const runs = richText.getRuns();
      if (runs && runs.length) {
        for (var i = 0; i < runs.length; i++) {
          const runUrl = runs[i].getLinkUrl();
          if (runUrl) {
            return runUrl;
          }
        }
      }
    }
  } catch (err) {
    // Ignorar errores de compatibilidad en RichTextValue.
  }
  return '';
}

function syncInvitationRow(rowNumber) {
  assertUserCanSyncInvitations();
  if (!rowNumber || rowNumber <= HEADER_ROW) {
    throw new Error('Selecciona una fila válida en la hoja "' + SHEET_NAME + '".');
  }
  const payload = buildInvitationPayload(rowNumber);
  if (!payload.invitation_id) {
    throw new Error('La fila seleccionada no tiene "ID Invitación".');
  }
  const response = supabaseUpsert([payload]);
  return {
    rowNumber: rowNumber,
    syncedRows: response.length,
  };
}

function syncActiveInvitation() {
  const permissions = getCurrentUserPermissions();
  const columnVisibility = getColumnVisibilityForEmail(permissions.email);
  const active = getActiveInvitation(columnVisibility);
  if (!active.rowNumber) {
    throw new Error('Selecciona una fila válida en la hoja "' + SHEET_NAME + '".');
  }
  return syncInvitationRow(active.rowNumber);
}

function getSidebarContext() {
  ensurePermissionsSheet();
  ensureColumnPermissionsSheet();
  ensureColumnDictionarySheet();

  let permissions;
  try {
    permissions = getCurrentUserPermissions();
  } catch (error) {
    permissions = {
      email: '',
      phone: '',
      role: PERMISSION_ROLE_VIEWER,
      status: PERMISSIONS_STATUS_DISABLED,
      enabled: false,
      canEdit: false,
      canSync: false,
      canDelete: false,
      canManageAccess: false,
    };
  }

  let columnVisibility = null;
  try {
    columnVisibility = getColumnVisibilityForEmail(permissions.email);
  } catch (error) {
    columnVisibility = null;
  }

  let activeInvitation;
  try {
    activeInvitation = getActiveInvitation(columnVisibility);
  } catch (error) {
    const activeSheet = SpreadsheetApp.getActiveSheet();
    activeInvitation = {
      rowNumber: null,
      sheetName: activeSheet ? activeSheet.getName() : '',
      expectedSheetName: SHEET_NAME,
      data: null,
      error: error && error.message ? String(error.message) : 'Error al obtener la fila activa.',
    };
  }

  let dropdownOptions = {};
  try {
    const optionMap = buildDropdownOptions();
    dropdownOptions = filterOptionMapByVisibility(optionMap, columnVisibility);
  } catch (error) {
    dropdownOptions = {};
  }

  let contactDirectory = {};
  try {
    const directory = buildContactDirectory();
    contactDirectory = filterContactDirectoryByVisibility(directory, columnVisibility);
  } catch (error) {
    contactDirectory = {};
  }

  return {
    permissions: permissions,
    activeInvitation: activeInvitation,
    dropdownOptions: dropdownOptions,
    contactDirectory: contactDirectory,
  };
}

function saveInvitation(request) {
  assertUserCanEditInvitations();
  const permissions = getCurrentUserPermissions();
  const columnVisibility = getColumnVisibilityForEmail(permissions.email);
  const rowNumber = request && request.rowNumber;
  const updates = filterUpdatesForColumnVisibility((request && request.updates) || {}, columnVisibility);
  if (!rowNumber || rowNumber <= HEADER_ROW) {
    throw new Error('Selecciona una fila válida en la hoja "' + SHEET_NAME + '".');
  }
  ensureInvitationIdForRow(rowNumber);
  applyKpiCalculations(rowNumber, updates);
  const updated = updateInvitationRow(rowNumber, updates, columnVisibility);
  return {
    rowNumber: rowNumber,
    data: updated,
  };
}

function acceptInvitation(request) {
  assertUserCanSyncInvitations();
  const permissions = getCurrentUserPermissions();
  const columnVisibility = getColumnVisibilityForEmail(permissions.email);
  const filteredRequest = {
    rowNumber: request && request.rowNumber,
    updates: filterUpdatesForColumnVisibility((request && request.updates) || {}, columnVisibility),
  };
  const saved = saveInvitation(filteredRequest);
  const syncResult = syncInvitationRow(saved.rowNumber);
  return {
    rowNumber: saved.rowNumber,
    data: getRowData(saved.rowNumber, columnVisibility),
    syncedRows: syncResult.syncedRows,
  };
}

function deleteInvitation(request) {
  assertUserCanDeleteInvitations();
  const rowNumber = request && request.rowNumber;
  if (!rowNumber || rowNumber <= HEADER_ROW) {
    throw new Error('Selecciona una fila válida para eliminar.');
  }
  const sheet = getInvitationSheet();
  const lastRow = sheet.getLastRow();
  if (rowNumber > lastRow) {
    throw new Error('La fila ' + rowNumber + ' no existe en la hoja.');
  }
  sheet.deleteRow(rowNumber);
  return {
    deletedRow: rowNumber,
  };
}

function listColumnMetadata() {
  return FIELD_SECTIONS.reduce(function (acc, section) {
    section.fields.forEach(function (field) {
      acc.push({
        alias: field.alias,
        header: field.header,
        section: section.id,
      });
    });
    return acc;
  }, []);
}

function buildHyperlinkFormula(url, text) {
  const safeUrl = escapeFormulaValue(url);
  const safeText = escapeFormulaValue(text);
  return '=HYPERLINK("' + safeUrl + '","' + safeText + '")';
}

function parseHyperlinkFormula(formula) {
  if (!formula || typeof formula !== 'string') {
    return null;
  }
  const trimmed = formula.trim();
  const match = /^=HYPERLINK\(\s*"((?:[^"]|"")*)"\s*(?:,\s*"((?:[^"]|"")*)"\s*)?\)$/i.exec(trimmed);
  if (!match) {
    return null;
  }
  const url = match[1] ? match[1].replace(/""/g, '"') : '';
  const text = match[2] ? match[2].replace(/""/g, '"') : '';
  return { url: url, text: text };
}

function escapeFormulaValue(value) {
  return String(value || '').replace(/"/g, '""');
}
