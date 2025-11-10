const INVITATION_STATIC_OPTIONS = {
  tipo_servicio: {
    values: ['Cotización Formal', 'Licitación', 'Parada de Planta', 'Adicional', 'Adenda'],
    palette: {
      'cotización formal': '#2563eb',
      licitación: '#7c3aed',
      'parada de planta': '#dc2626',
      adicional: '#f97316',
      adenda: '#059669',
    },
  },
  estado_cotizacion: {
    values: [
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
    ],
  },
  estado_propuesta: {
    values: [
      'En elaboración',
      'Enviada',
      'En evaluación',
      'Adjudicada',
      'No adjudicada',
      'Observada',
      'Cancelada',
    ],
  },
};

const INVITATION_CONTACT_ROLES = [
  {
    key: 'solicitante',
    header: 'Solicitante',
    email: { alias: 'correo_solicitante', header: 'Correo del Solicitante' },
    phone: { alias: 'telefono_solicitante', header: 'Teléfono del Solicitante' },
    description: 'Contacto principal que gestiona la invitación.',
  },
  {
    key: 'responsable_tecnico',
    header: 'Responsable Técnico',
    email: { alias: 'correo_responsable_tecnico', header: 'Correo del Resp. Téc.' },
    phone: { alias: 'telefono_responsable_tecnico', header: 'Teléfono del Resp. Téc.' },
    description: 'Profesional a cargo de la coordinación técnica.',
  },
  {
    key: 'responsable_economico',
    header: 'Responsable Económico',
    email: { alias: 'correo_responsable_economico', header: 'Correo del Resp. Eco.' },
    phone: { alias: 'telefono_responsable_economico', header: 'Teléfono del Resp. Eco.' },
    description: 'Contacto financiero y comercial del proceso.',
  },
];

const INVITATION_FIELD_ENTRIES = (function () {
  const entries = [
    ['invitation_id', 'ID Invitación', { placeholder: 'INV-0001', hidden: true }],
    ['cotizacion', 'Cotización'],
    ['descripcion', 'Descripción', { valueType: 'textarea', width: 12, rows: 3 }],
    [
      'tipo_servicio',
      'Tipo de Servicio',
      {
        valueType: 'select',
        width: 6,
        options: INVITATION_STATIC_OPTIONS.tipo_servicio.values,
        allowCustom: true,
        palette: INVITATION_STATIC_OPTIONS.tipo_servicio.palette,
      },
    ],
    [
      'estado_cotizacion',
      'Estado de Cotización',
      { valueType: 'select', options: INVITATION_STATIC_OPTIONS.estado_cotizacion.values, allowCustom: true },
    ],
    [
      'estado_propuesta',
      'Estado de Propuesta',
      { valueType: 'select', options: INVITATION_STATIC_OPTIONS.estado_propuesta.values, allowCustom: true },
    ],
    [
      'estado_pipeline',
      'Estado Pipeline',
      { helper: 'Consulta el "Diccionario Invitaciones" para las definiciones de pipeline.' },
    ],
    ['fecha_registro', 'Fecha de Registro', { valueType: 'date' }],
    ['cliente', 'Cliente', { valueType: 'select', optionsSource: 'cliente', allowCustom: true }],
    ['zona_trabajo', 'Zona de Trabajo', { valueType: 'select', optionsSource: 'zona_trabajo', allowCustom: true }],
    ['fecha_invitacion', 'Fecha de Invitación', { valueType: 'date' }],
    ['fecha_confirmacion', 'Fecha de Confirmación', { valueType: 'date' }],
    ['fecha_visita_tecnica', 'Fecha de Visita Téc.', { valueType: 'date' }],
    ['fecha_consultas', 'Fecha de Consultas', { valueType: 'date' }],
    ['fecha_abs_consultas', 'Fecha de Abs. Consultas', { valueType: 'date' }],
    ['fecha_presentacion', 'Fecha de Presentación', { valueType: 'date' }],
    ['fecha_envio_propuesta', 'Fecha de Envío de Propuesta', { valueType: 'date' }],
    ['hora_envio_propuesta', 'Hora de Envío de Propuesta', { valueType: 'time' }],
    ['orden_de_compra', 'Orden de Compra'],
    ['fecha_orden_compra', 'Fecha de la Orden de Compra', { valueType: 'date' }],
    [
      'link_carpeta_drive',
      'Link Carpeta Drive',
      { valueType: 'link', width: 12, helper: 'Guarda la carpeta como hipervínculo para acceder rápidamente.' },
    ],
    [
      'link_archivo_enviado',
      'Link Archivo Enviado',
      { valueType: 'link', width: 12, helper: 'Adjunta el enlace del archivo enviado al cliente.' },
    ],
    ['monto_ofertado', 'Monto ofertado', { valueType: 'number' }],
    [
      'moneda',
      'Moneda',
      {
        valueType: 'select',
        options: [
          { value: 'PEN', label: 'Sol (PEN)' },
          { value: 'USD', label: 'Dólar (USD)' },
        ],
        allowCustom: true,
        helper: 'Selecciona la moneda de origen o agrega una nueva si es necesario.',
      },
    ],
    [
      'moneda_normalizada_usd',
      'Moneda Normalizada (USD)',
      { readOnly: true, helper: 'Se actualiza automáticamente según la moneda seleccionada.' },
    ],
    ['tipo_cambio', 'Tipo de Cambio', { valueType: 'number' }],
    [
      'monto_ofertado_usd',
      'Monto Ofertado (USD)',
      {
        valueType: 'number',
        readOnly: true,
        helper: 'Calculado automáticamente según moneda, monto y tipo de cambio.',
      },
    ],
    ['dias_a_vencimiento', 'Días a Vencimiento', { valueType: 'number' }],
    ['riesgo_d3', 'Riesgo D-3', { valueType: 'boolean' }],
    ['riesgo_d1', 'Riesgo D-1', { valueType: 'boolean' }],
    ['enviado_a_tiempo', 'Enviado a Tiempo', { valueType: 'boolean' }],
    ['tiempo_respuesta_dias_hab', 'Tiempo de Respuesta (días hábiles)', { valueType: 'number' }],
    ['requiere_visita', 'Requiere Visita Técnica', { valueType: 'boolean' }],
    ['visita_ejecutada', 'Visita Ejecutada', { valueType: 'boolean' }],
    ['semana_iso', 'Semana ISO'],
    ['mes_anio', 'Mes-Año'],
    ['notas_kpi', 'Notas KPI', { valueType: 'textarea', width: 12, rows: 3 }],
  ];

  INVITATION_CONTACT_ROLES.forEach(function (role) {
    entries.push([
      role.key,
      role.header,
      {
        valueType: 'select',
        width: 12,
        columnClass: 'col-12 col-lg-4',
        optionsSource: role.key,
        allowCustom: true,
        helper: 'Selecciona un ' + role.header.toLowerCase() + ' existente o agrega uno nuevo.',
        group: role.key,
        groupLabel: role.header,
        groupDescription: role.description,
      },
    ]);
    entries.push([
      role.email.alias,
      role.email.header,
      { width: 12, columnClass: 'col-12 col-lg-4', group: role.key },
    ]);
    entries.push([
      role.phone.alias,
      role.phone.header,
      { width: 12, columnClass: 'col-12 col-lg-4', group: role.key },
    ]);
  });

  return entries;
})();

const INVITATION_FIELD_MAP = buildInvitationFieldMap(INVITATION_FIELD_ENTRIES);
const INVITATION_COLUMN_MAP = buildInvitationColumnMap(INVITATION_FIELD_MAP);

const INVITATION_SECTION_LAYOUT = [
  [
    'general',
    'Datos generales',
    'Información básica del proceso de cotización.',
    ['invitation_id', 'cotizacion', 'descripcion', 'tipo_servicio', 'estado_cotizacion', 'estado_propuesta', 'estado_pipeline', 'fecha_registro', 'cliente', 'zona_trabajo'],
  ],
  [
    'contactos',
    'Contactos clave',
    'Personas involucradas en la invitación.',
    flattenContactAliases(),
  ],
  [
    'cronograma',
    'Cronograma y entregables',
    'Fechas clave, entregables y documentación de soporte.',
    [
      'fecha_invitacion',
      'fecha_confirmacion',
      'fecha_visita_tecnica',
      'fecha_consultas',
      'fecha_abs_consultas',
      'fecha_presentacion',
      'fecha_envio_propuesta',
      'hora_envio_propuesta',
      'orden_de_compra',
      'fecha_orden_compra',
      'link_carpeta_drive',
      'link_archivo_enviado',
    ],
  ],
  [
    'kpi',
    'KPIs y métricas',
    'Seguimiento de indicadores clave y normalizaciones.',
    [
      'monto_ofertado',
      'moneda',
      'moneda_normalizada_usd',
      'tipo_cambio',
      'monto_ofertado_usd',
      'dias_a_vencimiento',
      'riesgo_d3',
      'riesgo_d1',
      'enviado_a_tiempo',
      'tiempo_respuesta_dias_hab',
      'requiere_visita',
      'visita_ejecutada',
      'semana_iso',
      'mes_anio',
      'notas_kpi',
    ],
  ],
];

function buildInvitationFieldMap(entries) {
  const map = {};
  entries.forEach(function (entry) {
    const alias = entry[0];
    const header = entry[1];
    const config = entry[2] || {};
    const field = Object.assign({ alias: alias, header: header, label: header, valueType: 'text', width: 6 }, config);
    if (!field.label) {
      field.label = header;
    }
    map[alias] = Object.freeze(field);
  });
  return Object.freeze(map);
}

function buildInvitationColumnMap(fieldMap) {
  const map = {};
  Object.keys(fieldMap).forEach(function (alias) {
    map[alias] = fieldMap[alias].header;
  });
  return Object.freeze(map);
}

function flattenContactAliases() {
  const aliases = [];
  INVITATION_CONTACT_ROLES.forEach(function (role) {
    aliases.push(role.key, role.email.alias, role.phone.alias);
  });
  return aliases;
}

function cloneInvitationField(alias) {
  const base = INVITATION_FIELD_MAP[alias];
  if (!base) {
    return null;
  }
  return JSON.parse(JSON.stringify(base));
}

function getInvitationFieldDefinition(alias) {
  const cloned = cloneInvitationField(alias);
  if (!cloned) {
    throw new Error('No se encontró la definición del campo: ' + alias);
  }
  return cloned;
}

function getInvitationFieldSections() {
  return INVITATION_SECTION_LAYOUT.map(function (section) {
    return {
      id: section[0],
      title: section[1],
      description: section[2],
      fields: section[3]
        .map(function (alias) {
          return cloneInvitationField(alias);
        })
        .filter(function (field) {
          return field !== null;
        }),
    };
  });
}

function getInvitationColumnMap() {
  return INVITATION_COLUMN_MAP;
}

function listColumnMetadata() {
  return INVITATION_SECTION_LAYOUT.reduce(function (acc, section) {
    section[3].forEach(function (alias) {
      if (INVITATION_FIELD_MAP[alias]) {
        acc.push({ alias: alias, header: INVITATION_FIELD_MAP[alias].header, section: section[0] });
      }
    });
    return acc;
  }, []);
}

function getContactRoles() {
  return INVITATION_CONTACT_ROLES.map(function (role) {
    return {
      key: role.key,
      email: role.email.alias,
      phone: role.phone.alias,
    };
  });
}
function getInvitationFieldBlueprint() {
  return INVITATION_FIELD_MAP;
}
