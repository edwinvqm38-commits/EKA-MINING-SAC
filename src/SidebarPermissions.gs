const PERMISSIONS_SHEET_NAME = 'Permisos Sidebar';
const PERMISSIONS_HEADERS = ['Correo', 'Rol', 'Notas'];
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
