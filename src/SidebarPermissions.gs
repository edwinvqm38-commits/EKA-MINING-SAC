const PERMISSIONS_SHEET_NAME = 'Permisos Invitaciones';
const LEGACY_PERMISSIONS_SHEET_NAME = 'Permisos Sidebar';
const LEGACY_COLUMN_PERMISSIONS_SHEET_NAME = 'Permisos Columnas';
const PERMISSIONS_BASE_HEADERS = ['Correo', 'Teléfono', 'Rol', 'Estado', 'Notas'];
const PERMISSIONS_STATUS_ENABLED = 'Habilitado';
const PERMISSIONS_STATUS_DISABLED = 'Deshabilitado';
const PERMISSIONS_STATUS_OPTIONS = [PERMISSIONS_STATUS_ENABLED, PERMISSIONS_STATUS_DISABLED];
const PERMISSION_ROLE_ADMIN = 'ADMIN';
const PERMISSION_ROLE_EDITOR = 'EDITOR';
const PERMISSION_ROLE_VIEWER = 'LECTOR';
const COLUMN_PERMISSIONS_SHEET_NAME = PERMISSIONS_SHEET_NAME;

function ensurePermissionsSheet() {
  const ss = SpreadsheetApp.getActive();
  const metadata = listColumnMetadata();
  const columnHeaders = metadata.map(function (item) {
    return item.header;
  });
  const headers = PERMISSIONS_BASE_HEADERS.concat(columnHeaders);

  const legacy = collectLegacyPermissionRows(metadata, headers.length);
  let sheet = ss.getSheetByName(PERMISSIONS_SHEET_NAME);
  let existingRows = [];

  if (sheet) {
    existingRows = extractPermissionSheetRows(sheet, metadata, headers.length);
  } else {
    sheet = ss.insertSheet(PERMISSIONS_SHEET_NAME);
  }

  const rowsToWrite = existingRows.length ? existingRows : legacy.rows;

  syncPermissionSheetLayout(sheet, headers);
  clearPermissionSheetRows(sheet, headers.length);
  applyBaseValidations(sheet, headers.length);

  if (rowsToWrite.length) {
    sheet.getRange(2, 1, rowsToWrite.length, headers.length).setValues(rowsToWrite);
  } else if (sheet.getLastRow() < 2) {
    seedDefaultAdminRow(sheet, headers.length);
  }

  ensureStatusDefaults(sheet);

  if (legacy.cleanupSheets.length) {
    legacy.cleanupSheets.forEach(function (legacySheet) {
      if (legacySheet && legacySheet.getParent()) {
        ss.deleteSheet(legacySheet);
      }
    });
  }

  return sheet;
}

function ensureColumnPermissionsSheet() {
  return ensurePermissionsSheet();
}

function collectLegacyPermissionRows(metadata, totalColumns) {
  const ss = SpreadsheetApp.getActive();
  const cleanupSheets = [];
  const rowsByEmail = {};

  const legacyUserSheet = ss.getSheetByName(LEGACY_PERMISSIONS_SHEET_NAME);
  const legacyColumnSheet = ss.getSheetByName(LEGACY_COLUMN_PERMISSIONS_SHEET_NAME);

  if (legacyUserSheet) {
    cleanupSheets.push(legacyUserSheet);
    const lastRow = legacyUserSheet.getLastRow();
    const lastColumn = legacyUserSheet.getLastColumn();
    if (lastRow > 1 && lastColumn > 0) {
      const values = legacyUserSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
      values.forEach(function (row) {
        const email = (row[0] || '').toString().trim().toLowerCase();
        if (!email) {
          return;
        }
        const role = (row[1] || PERMISSION_ROLE_VIEWER).toString().trim().toUpperCase();
        const notes = (row[2] || '').toString();
        rowsByEmail[email] = {
          email: email,
          phone: '',
          role: normalizePermissionRole(role),
          status: PERMISSIONS_STATUS_ENABLED,
          notes: notes,
          columns: {},
        };
      });
    }
  }

  const legacyMatrix = legacyColumnSheet
    ? extractLegacyColumnMatrix(legacyColumnSheet, metadata)
    : {};
  if (legacyColumnSheet) {
    cleanupSheets.push(legacyColumnSheet);
  }

  Object.keys(legacyMatrix).forEach(function (email) {
    if (!rowsByEmail[email]) {
      rowsByEmail[email] = {
        email: email,
        phone: '',
        role: PERMISSION_ROLE_VIEWER,
        status: PERMISSIONS_STATUS_ENABLED,
        notes: '',
        columns: {},
      };
    }
    rowsByEmail[email].columns = legacyMatrix[email];
  });

  const columnHeaders = metadata.map(function (item) {
    return item.header;
  });
  const aliasByHeader = {};
  metadata.forEach(function (item) {
    aliasByHeader[normalizeHeaderKey(item.header)] = item.alias;
  });
  const rows = Object.keys(rowsByEmail)
    .sort()
    .map(function (email) {
      const record = rowsByEmail[email];
      const rowValues = [
        record.email,
        record.phone || '',
        record.role || PERMISSION_ROLE_VIEWER,
        record.status || PERMISSIONS_STATUS_ENABLED,
        record.notes || '',
      ];
      columnHeaders.forEach(function (header) {
        const aliasKey = aliasByHeader[normalizeHeaderKey(header)];
        if (!aliasKey) {
          rowValues.push(true);
          return;
        }
        const matrix = record.columns || {};
        const value = Object.prototype.hasOwnProperty.call(matrix, aliasKey)
          ? matrix[aliasKey]
          : true;
        rowValues.push(value === false ? false : true);
      });
      while (rowValues.length < totalColumns) {
        rowValues.push(true);
      }
      return rowValues;
    });

  return {
    rows: rows,
    cleanupSheets: cleanupSheets,
  };
}

function extractLegacyColumnMatrix(sheet, metadata) {
  const matrix = {};
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= 1 || lastColumn <= 1) {
    return matrix;
  }
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const aliasByHeader = {};
  metadata.forEach(function (item) {
    aliasByHeader[normalizeHeaderKey(item.header)] = item.alias;
  });
  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  values.forEach(function (row) {
    const email = (row[0] || '').toString().trim().toLowerCase();
    if (!email) {
      return;
    }
    if (!matrix[email]) {
      matrix[email] = {};
    }
    for (var i = 1; i < headers.length; i++) {
      const header = headers[i];
      const alias = aliasByHeader[normalizeHeaderKey(header)];
      if (!alias) {
        continue;
      }
      const cellValue = row[i];
      const normalized = normalizeCheckboxValue(cellValue);
      if (normalized === null) {
        continue;
      }
      matrix[email][alias] = normalized;
    }
  });
  return matrix;
}

function extractPermissionSheetRows(sheet, metadata, totalColumns) {
  const rows = [];
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= 1 || lastColumn === 0) {
    return rows;
  }
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const headerIndex = {};
  headers.forEach(function (header, index) {
    const key = normalizeHeaderKey(header);
    if (key) {
      headerIndex[key] = index;
    }
  });
  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  const columnHeaders = metadata.map(function (item) {
    return item.header;
  });
  values.forEach(function (row) {
    const email = (getValueByHeader(row, headerIndex, 'correo') || '').toLowerCase();
    if (!email) {
      return;
    }
    const phone = getValueByHeader(row, headerIndex, 'teléfono');
    const role = normalizePermissionRole(
      getValueByHeader(row, headerIndex, 'rol') || PERMISSION_ROLE_VIEWER
    );
    const status = normalizePermissionStatus(
      getValueByHeader(row, headerIndex, 'estado') || PERMISSIONS_STATUS_ENABLED
    );
    const notes = getValueByHeader(row, headerIndex, 'notas');
    const rowValues = [email, phone, role, status, notes];
    columnHeaders.forEach(function (header) {
      const index = headerIndex[normalizeHeaderKey(header)];
      const cellValue = index === undefined ? true : row[index];
      const normalized = normalizeCheckboxValue(cellValue);
      rowValues.push(normalized === false ? false : true);
    });
    while (rowValues.length < totalColumns) {
      rowValues.push(true);
    }
    rows.push(rowValues);
  });
  return rows;
}

function syncPermissionSheetLayout(sheet, headers) {
  const requiredColumns = headers.length;
  const currentMaxColumns = sheet.getMaxColumns();
  if (currentMaxColumns < requiredColumns) {
    sheet.insertColumnsAfter(currentMaxColumns, requiredColumns - currentMaxColumns);
  }
  const lastColumn = sheet.getLastColumn();
  if (lastColumn > requiredColumns) {
    sheet.deleteColumns(requiredColumns + 1, lastColumn - requiredColumns);
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function clearPermissionSheetRows(sheet, totalColumns) {
  const maxRows = sheet.getMaxRows();
  if (maxRows <= 1) {
    return;
  }
  sheet.getRange(2, 1, maxRows - 1, totalColumns).clearContent();
}

function applyBaseValidations(sheet, totalColumns) {
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([PERMISSION_ROLE_ADMIN, PERMISSION_ROLE_EDITOR, PERMISSION_ROLE_VIEWER], true)
    .setAllowInvalid(false)
    .build();
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PERMISSIONS_STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 3, maxRows, 1).setDataValidation(roleRule);
  sheet.getRange(2, 4, maxRows, 1).setDataValidation(statusRule);
  applyColumnCheckboxValidation(sheet, totalColumns);
}

function applyColumnCheckboxValidation(sheet, totalColumns) {
  if (totalColumns <= PERMISSIONS_BASE_HEADERS.length) {
    return;
  }
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet
    .getRange(2, PERMISSIONS_BASE_HEADERS.length + 1, maxRows, totalColumns - PERMISSIONS_BASE_HEADERS.length)
    .setDataValidation(checkboxRule);
}

function ensureStatusDefaults(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }
  const range = sheet.getRange(2, 4, lastRow - 1, 1);
  const values = range.getValues();
  let needsUpdate = false;
  for (var i = 0; i < values.length; i++) {
    const value = (values[i][0] || '').toString().trim();
    if (!value) {
      values[i][0] = PERMISSIONS_STATUS_ENABLED;
      needsUpdate = true;
    }
  }
  if (needsUpdate) {
    range.setValues(values);
  }
}

function seedDefaultAdminRow(sheet, totalColumns) {
  const adminEmail = (getActiveUserEmail() || '').toLowerCase() || 'admin@empresa.com';
  const defaults = [
    adminEmail,
    '',
    PERMISSION_ROLE_ADMIN,
    PERMISSIONS_STATUS_ENABLED,
    'Administrador inicial. Actualiza esta tabla para delegar acceso.',
  ];
  while (defaults.length < totalColumns) {
    defaults.push(true);
  }
  sheet.getRange(2, 1, 1, defaults.length).setValues([defaults]);
}

function getActiveUserEmail() {
  const active = Session.getActiveUser();
  if (active && typeof active.getEmail === 'function') {
    const email = active.getEmail();
    if (email) {
      return email;
    }
  }
  const effective = Session.getEffectiveUser();
  if (effective && typeof effective.getEmail === 'function') {
    return effective.getEmail();
  }
  return '';
}

function getPermissionRecords() {
  const sheet = ensurePermissionsSheet();
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= 1 || lastColumn === 0) {
    return [];
  }
  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  const records = [];
  values.forEach(function (row) {
    const email = (row[0] || '').toString().trim().toLowerCase();
    if (!email) {
      return;
    }
    const phone = (row[1] || '').toString().trim();
    const role = normalizePermissionRole((row[2] || PERMISSION_ROLE_VIEWER).toString());
    const status = normalizePermissionStatus((row[3] || PERMISSIONS_STATUS_ENABLED).toString());
    const notes = (row[4] || '').toString();
    const enabled = status !== PERMISSIONS_STATUS_DISABLED;
    records.push({
      email: email,
      phone: phone,
      role: role,
      status: status,
      enabled: enabled,
      notes: notes,
      explicit: true,
    });
  });
  return records;
}

function getColumnPermissionMatrix() {
  const sheet = ensurePermissionsSheet();
  const metadata = listColumnMetadata();
  const aliasByHeader = {};
  metadata.forEach(function (item) {
    aliasByHeader[item.header] = item.alias;
  });
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const matrix = {};
  if (lastRow <= 1 || lastColumn <= PERMISSIONS_BASE_HEADERS.length) {
    return matrix;
  }
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  values.forEach(function (row) {
    const email = (row[0] || '').toString().trim().toLowerCase();
    if (!email) {
      return;
    }
    if (!matrix[email]) {
      matrix[email] = {};
    }
    for (var i = PERMISSIONS_BASE_HEADERS.length; i < headers.length; i++) {
      const header = headers[i];
      const alias = aliasByHeader[header];
      if (!alias) {
        continue;
      }
      const normalized = normalizeCheckboxValue(row[i]);
      if (normalized === null) {
        continue;
      }
      matrix[email][alias] = normalized;
    }
  });
  return matrix;
}

function getCurrentUserPermissions() {
  const email = (getActiveUserEmail() || '').toLowerCase();
  const records = getPermissionRecords();
  const admins = records.filter(function (record) {
    return record.role === PERMISSION_ROLE_ADMIN && record.enabled !== false;
  });
  let record = null;
  if (email) {
    record =
      records.find(function (candidate) {
        return candidate.email === email;
      }) || null;
  }
  if (!record && email && admins.length === 0) {
    record = {
      email: email,
      phone: '',
      role: PERMISSION_ROLE_ADMIN,
      status: PERMISSIONS_STATUS_ENABLED,
      enabled: true,
      notes: 'Administrador por defecto (no registrado en la tabla).',
      explicit: false,
    };
  }
  if (!record) {
    record = {
      email: email,
      phone: '',
      role: PERMISSION_ROLE_VIEWER,
      status: PERMISSIONS_STATUS_DISABLED,
      enabled: false,
      notes: '',
      explicit: false,
    };
  }
  const enabled = record.enabled !== false;
  const role = normalizePermissionRole(record.role || PERMISSION_ROLE_VIEWER);
  const canEdit = enabled && (role === PERMISSION_ROLE_ADMIN || role === PERMISSION_ROLE_EDITOR);
  const canSync = enabled && (role === PERMISSION_ROLE_ADMIN || role === PERMISSION_ROLE_EDITOR);
  const canDelete = enabled && role === PERMISSION_ROLE_ADMIN;
  const canManageAccess = enabled && role === PERMISSION_ROLE_ADMIN;
  return {
    email: record.email,
    phone: record.phone || '',
    role: role,
    status: record.status || (enabled ? PERMISSIONS_STATUS_ENABLED : PERMISSIONS_STATUS_DISABLED),
    enabled: enabled,
    canEdit: canEdit,
    canSync: canSync,
    canDelete: canDelete,
    canManageAccess: canManageAccess,
    hasExplicitGrant: !!record.explicit,
    notes: record.notes || '',
  };
}

function getColumnVisibilityForEmail(email) {
  const matrix = getColumnPermissionMatrix();
  const lowerEmail = (email || '').toLowerCase();
  const metadata = listColumnMetadata();
  const userConfig = matrix[lowerEmail] || {};
  const visibility = {};
  metadata.forEach(function (item) {
    if (item.alias === 'invitation_id') {
      visibility[item.alias] = true;
      return;
    }
    if (Object.prototype.hasOwnProperty.call(userConfig, item.alias)) {
      visibility[item.alias] = userConfig[item.alias] !== false;
    } else {
      visibility[item.alias] = true;
    }
  });
  return visibility;
}

function ensureColumnPermissionRow(email) {
  const normalizedEmail = (email || '').toString().trim().toLowerCase();
  if (!normalizedEmail) {
    return null;
  }
  const sheet = ensurePermissionsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const emailValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < emailValues.length; i++) {
      const candidate = (emailValues[i][0] || '').toString().trim().toLowerCase();
      if (candidate === normalizedEmail) {
        return i + 2;
      }
    }
  }
  const totalColumns = sheet.getLastColumn();
  const rowValues = [normalizedEmail, '', PERMISSION_ROLE_VIEWER, PERMISSIONS_STATUS_ENABLED, ''];
  while (rowValues.length < totalColumns) {
    rowValues.push(true);
  }
  sheet.appendRow(rowValues);
  return sheet.getLastRow();
}

function assertUserCanManageAccess() {
  const permissions = getCurrentUserPermissions();
  if (!permissions.canManageAccess) {
    throw new Error(
      'Solo un administrador habilitado puede gestionar accesos desde este panel. Revisa la hoja "' +
        PERMISSIONS_SHEET_NAME +
        '".'
    );
  }
}

function assertUserCanEditInvitations() {
  const permissions = getCurrentUserPermissions();
  if (!permissions.canEdit) {
    throw new Error(
      'No tienes permisos para editar invitaciones. Verifica tu rol en la hoja "' + PERMISSIONS_SHEET_NAME + '".'
    );
  }
}

function assertUserCanSyncInvitations() {
  const permissions = getCurrentUserPermissions();
  if (!permissions.canSync) {
    throw new Error(
      'No tienes permisos para sincronizar con Supabase. Revisa la configuración en "' + PERMISSIONS_SHEET_NAME + '".'
    );
  }
}

function assertUserCanDeleteInvitations() {
  const permissions = getCurrentUserPermissions();
  if (!permissions.canDelete) {
    throw new Error(
      'Solo un administrador habilitado puede eliminar invitaciones desde este panel. Revisa "' +
        PERMISSIONS_SHEET_NAME +
        '" para actualizar los permisos.'
    );
  }
}

function upsertPermissionUser(request) {
  assertUserCanManageAccess();
  const rawEmail = (request && request.email) || '';
  const email = rawEmail.toString().trim().toLowerCase();
  if (!email) {
    throw new Error('Proporciona un correo electrónico válido.');
  }
  const phone = (request && request.phone ? request.phone.toString() : '').trim();
  const role = normalizePermissionRole(
    ((request && request.role) || PERMISSION_ROLE_VIEWER).toString().trim().toUpperCase()
  );
  const status = normalizePermissionStatus(
    ((request && request.status) || PERMISSIONS_STATUS_ENABLED).toString()
  );
  const notes = (request && request.notes) ? request.notes.toString() : '';

  const sheet = ensurePermissionsSheet();
  const lastRow = sheet.getLastRow();
  let rowIndex = -1;
  if (lastRow > 1) {
    const emails = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < emails.length; i++) {
      const candidate = (emails[i][0] || '').toString().trim().toLowerCase();
      if (candidate === email) {
        rowIndex = i + 2;
        break;
      }
    }
  }
  if (rowIndex === -1) {
    const totalColumns = sheet.getLastColumn();
    const rowValues = [email, phone, role, status, notes];
    while (rowValues.length < totalColumns) {
      rowValues.push(true);
    }
    sheet.appendRow(rowValues);
  } else {
    const baseValues = [[email, phone, role, status, notes]];
    sheet.getRange(rowIndex, 1, 1, PERMISSIONS_BASE_HEADERS.length).setValues(baseValues);
  }
  return getPermissionRecords();
}

function saveColumnPermissionsForUser(request) {
  assertUserCanManageAccess();
  const rawEmail = (request && request.email) || '';
  const email = rawEmail.toString().trim().toLowerCase();
  if (!email) {
    throw new Error('Proporciona un correo válido para actualizar permisos.');
  }
  const columns = (request && request.columns) || {};
  const sheet = ensurePermissionsSheet();
  const metadata = listColumnMetadata();
  const aliasByHeader = {};
  metadata.forEach(function (item) {
    aliasByHeader[item.header] = item.alias;
  });
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowIndex = ensureColumnPermissionRow(email);
  const rowRange = sheet.getRange(rowIndex, 1, 1, headers.length);
  const rowValues = rowRange.getValues()[0];
  rowValues[0] = email;
  rowValues[1] = (rowValues[1] || '').toString();
  rowValues[2] = normalizePermissionRole(rowValues[2] || PERMISSION_ROLE_VIEWER);
  rowValues[3] = normalizePermissionStatus(rowValues[3] || PERMISSIONS_STATUS_ENABLED);
  for (var i = PERMISSIONS_BASE_HEADERS.length; i < headers.length; i++) {
    const header = headers[i];
    const alias = aliasByHeader[header];
    if (!alias) {
      rowValues[i] = true;
      continue;
    }
    const requested = Object.prototype.hasOwnProperty.call(columns, alias) ? columns[alias] : true;
    rowValues[i] = requested === false ? false : true;
  }
  rowRange.setValues([rowValues]);
  return {
    email: email,
    columns: getColumnVisibilityForEmail(email),
  };
}

function getPermissionManagementData() {
  assertUserCanManageAccess();
  return {
    users: getPermissionRecords(),
    columns: listColumnMetadata(),
    matrix: getColumnPermissionMatrix(),
    permissionsSheetName: PERMISSIONS_SHEET_NAME,
  };
}

function openPermissionsSheet() {
  assertUserCanManageAccess();
  const sheet = ensurePermissionsSheet();
  const ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(sheet);
  SpreadsheetApp.flush();
}

function openColumnPermissionsSheet() {
  openPermissionsSheet();
}

function normalizePermissionRole(role) {
  const normalized = (role || '').toString().trim().toUpperCase();
  if (normalized === PERMISSION_ROLE_ADMIN) {
    return PERMISSION_ROLE_ADMIN;
  }
  if (normalized === PERMISSION_ROLE_EDITOR) {
    return PERMISSION_ROLE_EDITOR;
  }
  return PERMISSION_ROLE_VIEWER;
}

function normalizePermissionStatus(status) {
  const normalized = (status || '').toString().trim().toLowerCase();
  if (normalized === PERMISSIONS_STATUS_DISABLED.toLowerCase()) {
    return PERMISSIONS_STATUS_DISABLED;
  }
  return PERMISSIONS_STATUS_ENABLED;
}

function normalizeCheckboxValue(value) {
  if (value === true || value === false) {
    return value;
  }
  if (value === null || value === '') {
    return null;
  }
  if (typeof value === 'number') {
    return value !== 0;
  }
  const text = value.toString().trim().toLowerCase();
  if (!text) {
    return null;
  }
  if (['true', 'sí', 'si', 'yes', 'y', '1', 'habilitado'].indexOf(text) !== -1) {
    return true;
  }
  if (['false', 'no', '0', 'n', 'deshabilitado'].indexOf(text) !== -1) {
    return false;
  }
  return null;
}

function normalizeHeaderKey(header) {
  if (header === null || header === undefined) {
    return '';
  }
  var key = header.toString().trim().toLowerCase();
  if (!key) {
    return '';
  }
  if (typeof key.normalize === 'function') {
    key = key.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }
  key = key.replace(/[^a-z0-9_ ]+/g, '');
  key = key.replace(/\s+/g, '_');
  return key;
}

function getValueByHeader(row, headerIndex, headerName) {
  const key = normalizeHeaderKey(headerName);
  const index = headerIndex[key];
  if (index === undefined || index >= row.length) {
    return '';
  }
  return (row[index] || '').toString().trim();
}
