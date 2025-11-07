const PERMISSIONS_SHEET_NAME = 'Permisos Sidebar';
const PERMISSIONS_HEADERS = ['Correo', 'Rol', 'Notas'];
const COLUMN_PERMISSIONS_SHEET_NAME = 'Permisos Columnas';
const COLUMN_PERMISSIONS_HEADER_EMAIL = 'Correo';
const PERMISSION_ROLE_ADMIN = 'ADMIN';
const PERMISSION_ROLE_EDITOR = 'EDITOR';
const PERMISSION_ROLE_VIEWER = 'LECTOR';

function ensurePermissionsSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(PERMISSIONS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PERMISSIONS_SHEET_NAME);
    sheet.getRange(1, 1, 1, PERMISSIONS_HEADERS.length).setValues([PERMISSIONS_HEADERS]);
    sheet.getRange(1, 1, 1, PERMISSIONS_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    const adminEmail = getActiveUserEmail();
    const initialEmail = adminEmail || 'coloca@tu.correo';
    sheet.getRange(2, 1, 1, PERMISSIONS_HEADERS.length).setValues([
      [
        initialEmail,
        PERMISSION_ROLE_ADMIN,
        'Administrador inicial. Actualiza esta lista para delegar acceso.',
      ],
    ]);
  } else {
    const headerRange = sheet.getRange(1, 1, 1, PERMISSIONS_HEADERS.length);
    const currentHeaders = headerRange.getValues()[0];
    let needsHeaderUpdate = false;
    for (let i = 0; i < PERMISSIONS_HEADERS.length; i++) {
      if ((currentHeaders[i] || '').toString().trim() !== PERMISSIONS_HEADERS[i]) {
        needsHeaderUpdate = true;
        break;
      }
    }
    if (needsHeaderUpdate) {
      headerRange.setValues([PERMISSIONS_HEADERS]);
      headerRange.setFontWeight('bold');
    }
    if (sheet.getFrozenRows() < 1) {
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function ensureColumnPermissionsSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(COLUMN_PERMISSIONS_SHEET_NAME);
  const metadata = listColumnMetadata();
  const headers = [COLUMN_PERMISSIONS_HEADER_EMAIL].concat(
    metadata.map(function (item) {
      return item.header;
    })
  );
  if (!sheet) {
    sheet = ss.insertSheet(COLUMN_PERMISSIONS_SHEET_NAME);
  }
  const requiredColumns = headers.length;
  const currentMaxColumns = sheet.getMaxColumns();
  if (currentMaxColumns < requiredColumns) {
    sheet.insertColumnsAfter(currentMaxColumns, requiredColumns - currentMaxColumns);
  }
  const lastColumn = sheet.getLastColumn();
  if (lastColumn > requiredColumns) {
    sheet.deleteColumns(requiredColumns + 1, lastColumn - requiredColumns);
  }
  sheet.getRange(1, 1, 1, requiredColumns).setValues([headers]);
  sheet.getRange(1, 1, 1, requiredColumns).setFontWeight('bold');
  sheet.setFrozenRows(1);
  applyColumnCheckboxValidation(sheet, requiredColumns);
  return sheet;
}

function applyColumnCheckboxValidation(sheet, headerLength) {
  if (headerLength <= 1) {
    return;
  }
  const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  const range = sheet.getRange(2, 2, maxRows, headerLength - 1);
  range.setDataValidation(rule);
}

function getActiveUserEmail() {
  const active = Session.getActiveUser();
  const email = active && typeof active.getEmail === 'function' ? active.getEmail() : '';
  if (email) {
    return email;
  }
  const effective = Session.getEffectiveUser();
  return effective && typeof effective.getEmail === 'function' ? effective.getEmail() : '';
}

function getPermissionRecords() {
  const sheet = ensurePermissionsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }
  const range = sheet.getRange(2, 1, lastRow - 1, PERMISSIONS_HEADERS.length);
  const values = range.getValues();
  return values
    .map(function (row) {
      const email = (row[0] || '').toString().trim().toLowerCase();
      if (!email) {
        return null;
      }
      const role = (row[1] || PERMISSION_ROLE_VIEWER).toString().trim().toUpperCase();
      const notes = (row[2] || '').toString();
      return {
        email: email,
        role: role,
        notes: notes,
        explicit: true,
      };
    })
    .filter(function (record) {
      return record !== null;
    });
}

function getColumnPermissionMatrix() {
  const sheet = ensureColumnPermissionsSheet();
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const metadata = listColumnMetadata();
  const aliasByHeader = {};
  metadata.forEach(function (item) {
    aliasByHeader[item.header] = item.alias;
  });
  const matrix = {};
  if (lastRow <= 1) {
    return matrix;
  }
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
      const alias = aliasByHeader[header];
      if (!alias) {
        continue;
      }
      const cellValue = row[i];
      if (cellValue === null || cellValue === '') {
        continue;
      }
      const normalized = typeof cellValue === 'string' ? cellValue.trim().toLowerCase() : cellValue;
      if (normalized === false || normalized === 'false' || normalized === 0) {
        matrix[email][alias] = false;
      } else if (normalized === true || normalized === 'true' || normalized === 1) {
        matrix[email][alias] = true;
      }
    }
  });
  return matrix;
}

function getCurrentUserPermissions() {
  const email = (getActiveUserEmail() || '').toLowerCase();
  const records = getPermissionRecords();
  const admins = records.filter(function (record) {
    return record.role === PERMISSION_ROLE_ADMIN;
  });
  let record = null;
  if (email) {
    record = records.find(function (candidate) {
      return candidate.email === email;
    }) || null;
  }
  if (!record && email && admins.length === 0) {
    record = {
      email: email,
      role: PERMISSION_ROLE_ADMIN,
      notes: 'Administrador por defecto (no registrado en la tabla).',
      explicit: false,
    };
  }
  if (!record) {
    record = {
      email: email,
      role: PERMISSION_ROLE_VIEWER,
      notes: '',
      explicit: false,
    };
  }
  const role = record.role || PERMISSION_ROLE_VIEWER;
  const canEdit = role === PERMISSION_ROLE_ADMIN || role === PERMISSION_ROLE_EDITOR;
  const canSync = role === PERMISSION_ROLE_ADMIN || role === PERMISSION_ROLE_EDITOR;
  const canDelete = role === PERMISSION_ROLE_ADMIN;
  const canManageAccess = role === PERMISSION_ROLE_ADMIN;
  return {
    email: email,
    role: role,
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
  const sheet = ensureColumnPermissionsSheet();
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
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowValues = [normalizedEmail];
  for (var columnIndex = 1; columnIndex < headers.length; columnIndex++) {
    rowValues.push(true);
  }
  sheet.appendRow(rowValues);
  return sheet.getLastRow();
}

function assertUserCanManageAccess() {
  const permissions = getCurrentUserPermissions();
  if (!permissions.canManageAccess) {
    throw new Error(
      'Solo un administrador puede gestionar accesos desde este panel. Revisa la hoja "' +
        PERMISSIONS_SHEET_NAME +
        '".'
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
  const role = ((request && request.role) || PERMISSION_ROLE_VIEWER).toString().trim().toUpperCase();
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
  const rowValues = [[email, role, notes]];
  if (rowIndex === -1) {
    sheet.appendRow(rowValues[0]);
  } else {
    sheet.getRange(rowIndex, 1, 1, PERMISSIONS_HEADERS.length).setValues(rowValues);
  }
  ensureColumnPermissionRow(email);
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
  const sheet = ensureColumnPermissionsSheet();
  const metadata = listColumnMetadata();
  const aliasByHeader = {};
  metadata.forEach(function (item) {
    aliasByHeader[item.header] = item.alias;
  });
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowIndex = ensureColumnPermissionRow(email);
  const rowValues = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  rowValues[0] = email;
  for (var i = 1; i < headers.length; i++) {
    const header = headers[i];
    const alias = aliasByHeader[header];
    if (!alias) {
      rowValues[i] = true;
      continue;
    }
    const requested = Object.prototype.hasOwnProperty.call(columns, alias) ? columns[alias] : true;
    rowValues[i] = requested === false ? false : true;
  }
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowValues]);
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
    columnPermissionSheetName: COLUMN_PERMISSIONS_SHEET_NAME,
  };
}

function openColumnPermissionsSheet() {
  assertUserCanManageAccess();
  const sheet = ensureColumnPermissionsSheet();
  const ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(sheet);
  SpreadsheetApp.flush();
}

function openPermissionsSheet() {
  assertUserCanManageAccess();
  const sheet = ensurePermissionsSheet();
  const ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(sheet);
  SpreadsheetApp.flush();
}

function assertUserCanEditInvitations() {
  const permissions = getCurrentUserPermissions();
  if (!permissions.canEdit) {
    throw new Error(
      'No tienes permisos para editar invitaciones. Solicita acceso en la hoja "' +
        PERMISSIONS_SHEET_NAME +
        '".'
    );
  }
}

function assertUserCanSyncInvitations() {
  const permissions = getCurrentUserPermissions();
  if (!permissions.canSync) {
    throw new Error(
      'No tienes permisos para sincronizar con Supabase. Pide acceso al administrador en "' +
        PERMISSIONS_SHEET_NAME +
        '".'
    );
  }
}

function assertUserCanDeleteInvitations() {
  const permissions = getCurrentUserPermissions();
  if (!permissions.canDelete) {
    throw new Error(
      'Solo un administrador puede eliminar registros desde el panel. Revisa la hoja "' +
        PERMISSIONS_SHEET_NAME +
        '".'
    );
  }
}
