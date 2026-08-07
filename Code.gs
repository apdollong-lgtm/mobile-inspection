/**
 * FactoryStudio Inspect — Google Apps Script Backend
 * เก็บข้อมูลการตรวจ QC/QA ใน Google Sheets
 */

const CONFIG = {
  SPREADSHEET_NAME: 'FactoryStudio Inspect — Database',
  SHEET_INSPECTIONS: 'Inspections',
  SHEET_LOG: 'Log',
  SHEET_CONFIG: 'Config_Checklist',
  SHEET_USERS: 'Users',
  SHEET_SESSIONS: 'Sessions',
  FOLDER_NAME: 'FactoryStudio_Inspect_Images',
  SESSION_HOURS: 12,
};

const PROP_SPREADSHEET_ID = 'SPREADSHEET_ID';
const PROP_FOLDER_ID = 'FOLDER_ID';
const PROP_AUTH_SALT = 'AUTH_SALT';

const USER_HEADERS = ['id', 'username', 'password_hash', 'name', 'role', 'active', 'created_at'];
const SESSION_HEADERS = ['token', 'user_id', 'username', 'name', 'role', 'expires_at', 'created_at'];

const HEADERS = [
  'id', 'no', 'date', 'type', 'lot', 'inspector',
  'checklist_json', 'note', 'result', 'nc_open', 'created_at',
  'status', 'approver', 'approve_note', 'nc_action', 'nc_closed_by', 'files'
];

/* ============================================================
   HTTP Entry Points
   ============================================================ */

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleRequest_(e, 'GET');
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('FactoryStudio Inspect — QC/QA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  return handleRequest_(e, 'POST');
}

function handleRequest_(e, method) {
  try {
    const params = method === 'POST' ? parseBody_(e) : (e && e.parameter) || {};
    const action = params.action;

    if (!action) {
      return jsonResponse_({ ok: false, error: 'Missing action parameter' });
    }

    let result;
    switch (action) {
      case 'ping':
        result = { ok: true, message: 'FactoryStudio Inspect API', time: new Date().toISOString() };
        break;
      case 'getConfig':
        result = { ok: true, data: getConfig_() };
        break;
      case 'getRecords':
        result = { ok: true, data: getRecords_(params) };
        break;
      case 'getRecord':
        result = { ok: true, data: getRecordById_(params.id) };
        break;
      case 'saveRecord':
        result = saveRecord_(params.record || params);
        break;
      case 'approveRecord':
        result = approveRecord_(params.id, params.approver, params.note, params.status);
        break;
      case 'closeNc':
        result = closeNc_(params.id, params.action_note, params.by);
        break;
      case 'nextNo':
        result = { ok: true, no: getNextInspectionNo_() };
        break;
      case 'getDashboard':
        result = { ok: true, data: getDashboardData_() };
        break;
      case 'exportCsv':
        result = { ok: true, csv: exportCsv_(), filename: 'inspection_export_' + Date.now() + '.csv' };
        break;
      case 'login':
        result = login_(params.username, params.password);
        break;
      case 'verifySession':
        result = { ok: true, user: verifySession_(params.token) };
        break;
      case 'logout':
        result = logout_(params.token);
        break;
      case 'call':
        var fnArgs = params.args;
        if (typeof fnArgs === 'string') {
          try { fnArgs = JSON.parse(fnArgs); } catch (e) { fnArgs = []; }
        }
        result = handleApiCall_(params.fn, fnArgs || []);
        break;
      default:
        result = { ok: false, error: 'Unknown action: ' + action };
    }

    return jsonResponse_(result);
  } catch (err) {
    logError_('handleRequest', err);
    return jsonResponse_({ ok: false, error: String(err.message || err) });
  }
}

/* ============================================================
   Sheet Setup (รันครั้งแรก)
   ============================================================ */

function setupSheets() {
  const ss = ensureSpreadsheet_();
  const isNew = !ss.getSheetByName(CONFIG.SHEET_INSPECTIONS);

  // 1. Inspections Sheet
  let sheet = ss.getSheetByName(CONFIG.SHEET_INSPECTIONS);
  if (!sheet) {
    const defaultSheet = ss.getSheets()[0];
    if (defaultSheet.getName() === 'Sheet1' && ss.getSheets().length === 1) {
      defaultSheet.setName(CONFIG.SHEET_INSPECTIONS);
      sheet = defaultSheet;
    } else {
      sheet = ss.insertSheet(CONFIG.SHEET_INSPECTIONS);
    }
  }

  // Update Headers if needed
  const currentHeaders = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  if (currentHeaders.length < HEADERS.length || currentHeaders[0] !== HEADERS[0]) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#0B1F3A')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }

  // 2. Config Sheet
  let configSheet = ss.getSheetByName(CONFIG.SHEET_CONFIG);
  if (!configSheet) {
    configSheet = ss.insertSheet(CONFIG.SHEET_CONFIG);
    configSheet.getRange(1, 1, 1, 2).setValues([['Inspection Type', 'Checklist Items (Comma separated)']]);
    configSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
    configSheet.setFrozenRows(1);
    
    // Seed default data
    const defaultData = [
      ["Raw Material Inspection", "ลักษณะภายนอก, สี / กลิ่น, ความชื้น, สิ่งแปลกปลอม, เอกสาร COA, อุณหภูมิรับเข้า"],
      ["Packaging Inspection", "สภาพบรรจุภัณฑ์, การพิมพ์ฉลาก, การซีล / ปิดผนึก, น้ำหนักบรรจุ, วันหมดอายุ"],
      ["Line Inspection", "ความสะอาดไลน์, อุณหภูมิเครื่อง, ความเร็วสายพาน, การตั้งค่าเครื่อง, สิ่งปนเปื้อน"],
      ["Hygiene Inspection", "ความสะอาดพื้นที่, สุขอนามัยพนักงาน, อุปกรณ์ป้องกัน, น้ำยาทำความสะอาด, การกำจัดของเสีย"],
      ["FG Inspection", "ลักษณะสินค้า, น้ำหนัก / ขนาด, การซีล / Retort, ฉลาก / บาร์โค้ด, การบรรจุลงกล่อง"],
      ["Lab Sample Inspection", "ค่า pH, ปริมาณจุลินทรีย์, ความชื้น, สารตกค้าง, สี / เนื้อสัมผัส"],
      ["Supplier Audit", "ระบบเอกสาร, สภาพโรงงาน, การจัดเก็บ, มาตรฐาน / ใบรับรอง, การขนส่ง"]
    ];
    configSheet.getRange(2, 1, defaultData.length, 2).setValues(defaultData);
    configSheet.setColumnWidth(1, 250);
    configSheet.setColumnWidth(2, 500);
  }

  // 3. Log Sheet
  let logSheet = ss.getSheetByName(CONFIG.SHEET_LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.SHEET_LOG);
    logSheet.getRange(1, 1, 1, 3).setValues([['timestamp', 'action', 'detail']]);
    logSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#0B1F3A').setFontColor('#FFFFFF');
    logSheet.setFrozenRows(1);
  }

  // 4. Users Sheet
  setupUsersSheet_(ss);

  // 5. Sessions Sheet
  setupSessionsSheet_(ss);

  // 6. Ensure Folder
  ensureFolder_();

  writeLog_('SETUP', isNew ? 'Created & linked spreadsheet' : 'Sheets verified');
  SpreadsheetApp.flush();

  const result = {
    ok: true,
    spreadsheetId: ss.getId(),
    url: ss.getUrl(),
    message: isNew ? 'สร้าง Google Sheet และเชื่อมอัตโนมัติเรียบร้อย' : 'อัปเดต Google Sheet ที่มีอยู่แล้วเรียบร้อย',
  };
  Logger.log(result.message);
  return result;
}

/* ============================================================
   Authentication
   ============================================================ */

function setupUsersSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_USERS);
    sheet.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]);
    sheet.getRange(1, 1, 1, USER_HEADERS.length)
      .setFontWeight('bold').setBackground('#0B1F3A').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(4, 180);

    const now = new Date().toISOString();
    const defaults = [
      ['u1', 'inspector', hashPassword_('inspector123'), 'สมชาย ใจดี', 'Inspector', true, now],
      ['u2', 'manager', hashPassword_('manager123'), 'วราภรณ์ ศรีสุข', 'Manager', true, now],
    ];
    sheet.getRange(2, 1, defaults.length, USER_HEADERS.length).setValues(defaults);
    writeLog_('SETUP', 'Created default users: inspector / manager');
  }
}

function setupSessionsSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_SESSIONS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_SESSIONS);
    sheet.getRange(1, 1, 1, SESSION_HEADERS.length).setValues([SESSION_HEADERS]);
    sheet.getRange(1, 1, 1, SESSION_HEADERS.length)
      .setFontWeight('bold').setBackground('#0B1F3A').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
}

function initAuthSalt_() {
  const salt = Utilities.getUuid() + Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty(PROP_AUTH_SALT, salt);
  return salt;
}

function hashPassword_(password) {
  const props = PropertiesService.getScriptProperties();
  let salt = props.getProperty(PROP_AUTH_SALT);
  if (!salt) salt = initAuthSalt_();
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + String(password)
  );
  return Utilities.base64Encode(digest);
}

function getUsersSheet_() {
  const ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  }
  return sheet;
}

function getSessionsSheet_() {
  const ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.SHEET_SESSIONS);
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName(CONFIG.SHEET_SESSIONS);
  }
  return sheet;
}

function findUserByUsername_(username) {
  const sheet = getUsersSheet_();
  const data = sheet.getDataRange().getValues();
  const uname = String(username).trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === uname) {
      return {
        id: data[i][0],
        username: data[i][1],
        passwordHash: data[i][2],
        name: data[i][3],
        role: data[i][4],
        active: data[i][5] === true || data[i][5] === 'TRUE' || data[i][5] === 1,
      };
    }
  }
  return null;
}

function login_(username, password) {
  if (!username || !password) throw new Error('กรุณากรอก Username และ Password');

  const user = findUserByUsername_(username);
  if (!user) throw new Error('Username หรือ Password ไม่ถูกต้อง');
  if (!user.active) throw new Error('บัญชีนี้ถูกระงับการใช้งาน');

  if (hashPassword_(password) !== user.passwordHash) {
    writeLog_('LOGIN_FAIL', username);
    throw new Error('Username หรือ Password ไม่ถูกต้อง');
  }

  cleanExpiredSessions_();
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  const now = new Date();
  const expires = new Date(now.getTime() + CONFIG.SESSION_HOURS * 60 * 60 * 1000);

  getSessionsSheet_().appendRow([
    token, user.id, user.username, user.name, user.role, expires, now,
  ]);

  writeLog_('LOGIN', user.username + ' (' + user.role + ')');
  return {
    ok: true,
    token: token,
    user: { id: user.id, username: user.username, name: user.name, role: user.role },
    expiresAt: expires.toISOString(),
  };
}

function verifySession_(token) {
  if (!token) throw new Error('กรุณาเข้าสู่ระบบ');
  cleanExpiredSessions_();

  const sheet = getSessionsSheet_();
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      const expires = new Date(data[i][5]);
      if (expires <= now) {
        sheet.deleteRow(i + 1);
        throw new Error('Session หมดอายุ กรุณาเข้าสู่ระบบใหม่');
      }
      return {
        id: data[i][1],
        username: data[i][2],
        name: data[i][3],
        role: data[i][4],
        token: token,
      };
    }
  }
  throw new Error('Session ไม่ถูกต้อง กรุณาเข้าสู่ระบบใหม่');
}

function requireAuth_(token, roles) {
  const user = verifySession_(token);
  if (roles && roles.length && roles.indexOf(user.role) === -1) {
    throw new Error('ไม่มีสิทธิ์เข้าถึงฟังก์ชันนี้');
  }
  return user;
}

function logout_(token) {
  if (!token) return { ok: true };
  const sheet = getSessionsSheet_();
  const data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === token) {
      sheet.deleteRow(i + 1);
      writeLog_('LOGOUT', data[i][2]);
      break;
    }
  }
  return { ok: true };
}

function cleanExpiredSessions_() {
  try {
    const sheet = getSessionsSheet_();
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    for (var i = data.length - 1; i >= 1; i--) {
      if (new Date(data[i][5]) <= now) {
        sheet.deleteRow(i + 1);
      }
    }
  } catch (e) { /* ignore */ }
}

function changePassword_(token, oldPassword, newPassword) {
  const user = requireAuth_(token);
  if (!newPassword || String(newPassword).length < 6) {
    throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 6 ตัวอักษร');
  }

  const full = findUserByUsername_(user.username);
  if (hashPassword_(oldPassword) !== full.passwordHash) {
    throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
  }

  const sheet = getUsersSheet_();
  const data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === user.id) {
      sheet.getRange(i + 1, 3).setValue(hashPassword_(newPassword));
      writeLog_('CHANGE_PW', user.username);
      return { ok: true };
    }
  }
  throw new Error('ไม่พบผู้ใช้');
}

function getUsers_() {
  const sheet = getUsersSheet_();
  const data = sheet.getDataRange().getValues();
  const users = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    users.push({
      id: data[i][0],
      username: data[i][1],
      name: data[i][3],
      role: data[i][4],
      active: data[i][5] === true || data[i][5] === 'TRUE' || data[i][5] === 1,
      createdAt: data[i][6] instanceof Date ? data[i][6].toISOString() : String(data[i][6] || ''),
    });
  }
  return users;
}

function createUser_(token, input) {
  requireAuth_(token, ['Manager']);
  const username = String(input.username || '').trim().toLowerCase();
  const password = String(input.password || '');
  const name = String(input.name || '').trim();
  const role = input.role === 'Manager' ? 'Manager' : 'Inspector';

  if (!username || username.length < 3) throw new Error('Username ต้องมีอย่างน้อย 3 ตัวอักษร');
  if (!password || password.length < 6) throw new Error('Password ต้องมีอย่างน้อย 6 ตัวอักษร');
  if (!name) throw new Error('กรุณากรอกชื่อ');
  if (findUserByUsername_(username)) throw new Error('Username นี้มีอยู่แล้ว');

  const id = 'u' + Date.now();
  getUsersSheet_().appendRow([id, username, hashPassword_(password), name, role, true, new Date()]);
  writeLog_('CREATE_USER', username);
  return { ok: true, user: { id: id, username: username, name: name, role: role, active: true } };
}

function updateUser_(token, userId, input) {
  const actor = requireAuth_(token, ['Manager']);
  const sheet = getUsersSheet_();
  const data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== userId) continue;

    if (input.name !== undefined) sheet.getRange(i + 1, 4).setValue(String(input.name).trim());
    if (input.role !== undefined) sheet.getRange(i + 1, 5).setValue(input.role === 'Manager' ? 'Manager' : 'Inspector');
    if (input.active !== undefined) {
      if (data[i][0] === actor.id && !input.active) {
        throw new Error('ไม่สามารถระงับบัญชีของตัวเองได้');
      }
      sheet.getRange(i + 1, 6).setValue(!!input.active);
    }
    if (input.password) {
      if (String(input.password).length < 6) throw new Error('Password ต้องมีอย่างน้อย 6 ตัวอักษร');
      sheet.getRange(i + 1, 3).setValue(hashPassword_(input.password));
    }

    writeLog_('UPDATE_USER', data[i][1]);
    return { ok: true };
  }
  throw new Error('ไม่พบผู้ใช้');
}

/* ============================================================
   CRUD & Business Logic
   ============================================================ */

function getConfig_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG.SHEET_CONFIG);
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName(CONFIG.SHEET_CONFIG);
  }
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { types: [], checklists: {} };
  
  const types = [];
  const checklists = {};
  for (let i = 1; i < data.length; i++) {
    const type = data[i][0];
    const items = String(data[i][1]).split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
    if (type) {
      types.push(type);
      checklists[type] = items;
    }
  }
  return { types: types, checklists: checklists };
}

function getRecords_(params) {
  const sheet = getInspectionsSheet_();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  const headers = rows[0];
  let records = rows.slice(1)
    .filter(function (row) { return row[0]; })
    .map(function (row) { return rowToRecord_(headers, row); });

  if (params.type) records = records.filter(function (r) { return r.type === params.type; });
  if (params.result) records = records.filter(function (r) { return r.result === params.result; });
  if (params.status) records = records.filter(function (r) { return r.status === params.status; });
  if (params.search) {
    const q = String(params.search).toLowerCase();
    records = records.filter(function (r) {
      return (r.no + r.lot + r.inspector).toLowerCase().indexOf(q) !== -1;
    });
  }
  
  records.sort(function (a, b) { return new Date(b.date) - new Date(a.date); });
  return records;
}

function getRecordById_(id) {
  if (!id) return null;
  const sheet = getInspectionsSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      return rowToRecord_(headers, data[i]);
    }
  }
  return null;
}

function saveRecord_(input) {
  const record = normalizeRecord_(input.record || input);
  
  // Handle file uploads (Base64)
  const filesPayload = input.filesPayload || [];
  let fileUrls = [];
  
  if (filesPayload.length > 0) {
    const folder = ensureFolder_();
    filesPayload.forEach(function(f, index) {
      try {
        const parts = f.data.split(',');
        const base64Data = parts.length > 1 ? parts[1] : parts[0];
        const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), f.mime, record.no + '_' + index + '_' + f.name);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrls.push(file.getUrl());
      } catch (e) {
        logError_('FileUpload', e);
      }
    });
    record.files = fileUrls.join('\n');
  }

  const sheet = getInspectionsSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || HEADERS;

  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === record.id) {
      rowIndex = i + 1;
      break;
    }
  }

  const row = recordToRow_(headers, record);

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, rowIndex, row.length).setValues([row]);
    writeLog_('UPDATE', record.no);
  } else {
    sheet.appendRow(row);
    writeLog_('CREATE', record.no);
  }

  SpreadsheetApp.flush();
  return { ok: true, data: record };
}

function approveRecord_(id, approver, note, status) {
  if (!id) throw new Error('Missing id');
  const sheet = getInspectionsSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const statusCol = headers.indexOf('status') + 1;
  const approverCol = headers.indexOf('approver') + 1;
  const noteCol = headers.indexOf('approve_note') + 1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, statusCol).setValue(status);
      sheet.getRange(i + 1, approverCol).setValue(approver || '');
      sheet.getRange(i + 1, noteCol).setValue(note || '');
      writeLog_('APPROVE', id + ' -> ' + status);
      return { ok: true, id: id, status: status };
    }
  }
  throw new Error('Record not found: ' + id);
}

function closeNc_(id, action_note, by) {
  if (!id) throw new Error('Missing id');
  const sheet = getInspectionsSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const ncCol = headers.indexOf('nc_open') + 1;
  const actionCol = headers.indexOf('nc_action') + 1;
  const byCol = headers.indexOf('nc_closed_by') + 1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, ncCol).setValue(false);
      sheet.getRange(i + 1, actionCol).setValue(action_note || '');
      sheet.getRange(i + 1, byCol).setValue(by || '');
      writeLog_('CLOSE_NC', id);
      return { ok: true, id: id };
    }
  }
  throw new Error('Record not found: ' + id);
}

/* ============================================================
   Dashboard & Export
   ============================================================ */

function getDashboardData_() {
  const records = getRecords_({});

  const total = records.length;
  const pass = records.filter(function (r) { return r.result === 'PASS'; }).length;
  const fail = records.filter(function (r) { return r.result === 'FAIL'; }).length;
  const hold = records.filter(function (r) { return r.result === 'HOLD'; }).length;
  const ncOpen = records.filter(function (r) { return r.ncOpen; }).length;

  const today = new Date();
  const last14 = [];
  for (var d = 13; d >= 0; d--) {
    var day = new Date(today);
    day.setDate(today.getDate() - d);
    day.setHours(0, 0, 0, 0);

    var dayRecords = records.filter(function (r) {
      var rd = new Date(r.date);
      return rd.getFullYear() === day.getFullYear() &&
        rd.getMonth() === day.getMonth() &&
        rd.getDate() === day.getDate();
    });

    last14.push({
      date: day.toISOString(),
      label: pad_(day.getDate(), 2) + '/' + pad_(day.getMonth() + 1, 2),
      pass: dayRecords.filter(function (r) { return r.result === 'PASS'; }).length,
      hold: dayRecords.filter(function (r) { return r.result === 'HOLD'; }).length,
      fail: dayRecords.filter(function (r) { return r.result === 'FAIL'; }).length,
      total: dayRecords.length,
    });
  }

  const cfg = getConfig_();
  const byType = cfg.types.map(function (t) {
    var rs = records.filter(function (r) { return r.type === t; });
    var p = rs.filter(function (r) { return r.result === 'PASS'; }).length;
    return {
      type: t,
      count: rs.length,
      passRate: rs.length ? Math.round(p / rs.length * 100) : 0,
    };
  }).filter(function (x) { return x.count > 0; });

  return {
    kpi: { total: total, pass: pass, fail: fail, hold: hold, ncOpen: ncOpen, passRate: total ? Math.round(pass / total * 100) : 0 },
    trend14: last14,
    byType: byType,
    recent: records.slice(0, 10),
    ncList: records.filter(function (r) { return r.ncOpen; }).slice(0, 20),
  };
}

function exportCsv_() {
  const records = getRecords_({});
  const lines = [
    'Inspection No.,Date,Type,Lot/Batch,Inspector,Result,Status,Approver,Checklist,Note,NC Action'
  ];

  records.forEach(function (r) {
    var cl = r.checklist.map(function (c) { return c.name + ':' + c.result; }).join(' | ');
    var row = [
      r.no, formatDateTh_(r.date), r.type, r.lot, r.inspector, r.result, r.status, r.approver||'', cl, r.note || '', r.ncAction || ''
    ].map(csvEscape_).join(',');
    lines.push(row);
  });

  return '\uFEFF' + lines.join('\n');
}

function getNextInspectionNo_() {
  const d = new Date();
  const ymd = d.getFullYear() + pad_(d.getMonth() + 1, 2) + pad_(d.getDate(), 2);
  const prefix = 'INS-' + ymd + '-';

  const records = getRecords_({});
  var max = 0;
  records.forEach(function (r) {
    if (r.no && r.no.indexOf(prefix) === 0) {
      var num = parseInt(r.no.slice(prefix.length), 10);
      if (!isNaN(num) && num > max) max = num;
    }
  });

  return prefix + pad_(max + 1, 3);
}

function seedSampleRecords_(count) {
  count = Number(count) || 10;
  const cfg = getConfig_();
  const types = cfg.types.length ? cfg.types : ['Raw Material Inspection'];
  const inspectors = ['สมชาย ใจดี', 'วราภรณ์ ศรีสุข', 'ธนกร พงษ์ไพ', 'อรทัย มั่นคง'];
  const baseResults = ['PASS', 'PASS', 'PASS', 'HOLD', 'FAIL'];
  const created = [];
  const today = new Date();

  for (var i = 0; i < count; i++) {
    const type = types[i % types.length];
    const items = cfg.checklists[type] || ['รายการตรวจ'];
    const target = baseResults[i % baseResults.length];

    const checklist = items.map(function (name, idx) {
      var v = 'pass';
      if (target === 'FAIL' && idx === 0) v = 'fail';
      else if (target === 'HOLD' && idx === items.length - 1) v = 'hold';
      return { name: name, result: v };
    });

    const result = computeResult_(checklist);
    const day = new Date(today);
    day.setDate(today.getDate() - (i % 7));
    day.setHours(8 + (i % 9), (i * 11) % 60, 0, 0);

    const ymd = day.getFullYear() + pad_(day.getMonth() + 1, 2) + pad_(day.getDate(), 2);
    const record = {
      id: 'r' + Date.now() + i + Math.floor(Math.random() * 1000),
      no: 'INS-' + ymd + '-' + pad_(100 + i, 3),
      date: day.toISOString(),
      type: type,
      lot: 'LOT-' + ymd.slice(2) + '-' + pad_(10 + i, 3),
      inspector: inspectors[i % inspectors.length],
      checklist: checklist,
      note: i % 4 === 0 ? 'ข้อมูลตัวอย่าง · สร้างจาก Web App' : '',
      result: result,
      ncOpen: result === 'FAIL',
      status: result === 'PASS' ? 'Approved' : 'Pending',
      approver: result === 'PASS' ? 'วราภรณ์ ศรีสุข' : '',
    };

    saveRecord_({ record: record });
    created.push(record.no);
  }

  writeLog_('SEED', 'Created ' + created.length + ' sample records');
  return { ok: true, count: created.length, records: created };
}

/* ============================================================
   Record Transform
   ============================================================ */

function normalizeRecord_(input) {
  var checklist = input.checklist;
  if (typeof checklist === 'string') {
    try { checklist = JSON.parse(checklist); } catch (e) { checklist = []; }
  }

  var result = input.result || computeResult_(checklist);
  var id = input.id || ('r' + Date.now());
  var no = input.no || getNextInspectionNo_();

  return {
    id: id,
    no: no,
    date: input.date || new Date().toISOString(),
    type: input.type || '',
    lot: String(input.lot || '').trim(),
    inspector: input.inspector || 'ไม่ระบุ',
    checklist: checklist || [],
    note: input.note || '',
    result: result,
    ncOpen: input.ncOpen !== undefined ? !!input.ncOpen : result === 'FAIL',
    createdAt: input.createdAt || new Date().toISOString(),
    status: input.status || 'Pending',
    approver: input.approver || '',
    approveNote: input.approveNote || '',
    ncAction: input.ncAction || '',
    ncClosedBy: input.ncClosedBy || '',
    files: input.files || ''
  };
}

function computeResult_(checklist) {
  if (!checklist || !checklist.length) return 'PASS';
  if (checklist.some(function (c) { return c.result === 'fail'; })) return 'FAIL';
  if (checklist.some(function (c) { return c.result === 'hold'; })) return 'HOLD';
  return 'PASS';
}

function rowToRecord_(headers, row) {
  var obj = {};
  headers.forEach(function (h, i) { obj[h] = row[i]; });

  var checklist = [];
  try { checklist = JSON.parse(obj.checklist_json || '[]'); } catch (e) { checklist = []; }

  return {
    id: obj.id,
    no: obj.no,
    date: obj.date instanceof Date ? obj.date.toISOString() : String(obj.date),
    type: obj.type,
    lot: obj.lot,
    inspector: obj.inspector,
    checklist: checklist,
    note: obj.note || '',
    result: obj.result,
    ncOpen: obj.nc_open === true || obj.nc_open === 'TRUE' || obj.nc_open === 1,
    createdAt: obj.created_at instanceof Date ? obj.created_at.toISOString() : String(obj.created_at || ''),
    status: obj.status || 'Pending',
    approver: obj.approver || '',
    approveNote: obj.approve_note || '',
    ncAction: obj.nc_action || '',
    ncClosedBy: obj.nc_closed_by || '',
    files: obj.files || ''
  };
}

function recordToRow_(headers, record) {
  return headers.map(function (h) {
    switch (h) {
      case 'checklist_json': return JSON.stringify(record.checklist);
      case 'nc_open': return !!record.ncOpen;
      case 'created_at': return record.createdAt || new Date();
      case 'date': return record.date;
      case 'approve_note': return record.approveNote;
      case 'nc_action': return record.ncAction;
      case 'nc_closed_by': return record.ncClosedBy;
      default: return record[h] !== undefined ? record[h] : '';
    }
  });
}

/* ============================================================
   Helpers
   ============================================================ */

function ensureSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(PROP_SPREADSHEET_ID);
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) {
      props.deleteProperty(PROP_SPREADSHEET_ID);
      id = null;
    }
  }
  const ss = SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
  id = ss.getId();
  props.setProperty(PROP_SPREADSHEET_ID, id);
  return ss;
}

function ensureFolder_() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty(PROP_FOLDER_ID);
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch(e) {
      props.deleteProperty(PROP_FOLDER_ID);
    }
  }
  const folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(CONFIG.FOLDER_NAME);
  }
  props.setProperty(PROP_FOLDER_ID, folder.getId());
  return folder;
}

function getSpreadsheet_() { return ensureSpreadsheet_(); }
function getInspectionsSheet_() {
  const ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.SHEET_INSPECTIONS);
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName(CONFIG.SHEET_INSPECTIONS);
  }
  return sheet;
}

function parseBody_(e) {
  if (!e || !e.postData || !e.postData.contents) return {};
  try { return JSON.parse(e.postData.contents); } catch (err) { return (e && e.parameter) || {}; }
}
function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function pad_(n, len) { return String(n).padStart(len, '0'); }
function formatDateTh_(iso) {
  var d = new Date(iso);
  return pad_(d.getDate(), 2) + '/' + pad_(d.getMonth() + 1, 2) + '/' + d.getFullYear() + ' ' + pad_(d.getHours(), 2) + ':' + pad_(d.getMinutes(), 2);
}
function csvEscape_(val) { return '"' + String(val).replace(/"/g, '""') + '"'; }
function writeLog_(action, detail) {
  try {
    var ss = getSpreadsheet_();
    var log = ss.getSheetByName(CONFIG.SHEET_LOG);
    if (log) log.appendRow([new Date(), action, detail]);
  } catch (e) { }
}
function logError_(where, err) {
  Logger.log('[' + where + '] ' + (err.message || err));
  writeLog_('ERROR', where + ': ' + (err.message || err));
}

/* ============================================================
   API สำหรับ index.html (google.script.run)
   ============================================================ */

function apiLogin(username, password) { return login_(username, password); }
function apiVerifySession(token) {
  const user = verifySession_(token);
  return { ok: true, user: user };
}
function apiLogout(token) { return logout_(token); }
function apiChangePassword(token, oldPw, newPw) { return changePassword_(token, oldPw, newPw); }
function apiGetUsers(token) {
  requireAuth_(token, ['Manager']);
  return getUsers_();
}
function apiCreateUser(token, input) { return createUser_(token, input); }
function apiUpdateUser(token, userId, input) { return updateUser_(token, userId, input); }

function apiGetRecords(token) {
  requireAuth_(token);
  var records = getRecords_({});
  if (records.length === 0) {
    seedSampleRecords_(10);
    records = getRecords_({});
  }
  return records;
}
function apiSaveRecord(token, record, filesPayload) {
  const user = requireAuth_(token);
  record.inspector = record.inspector || user.name;
  return saveRecord_({ record: record, filesPayload: filesPayload });
}
function apiGetNextNo(token) {
  requireAuth_(token);
  return getNextInspectionNo_();
}
function apiGetConfig(token) {
  requireAuth_(token);
  return getConfig_();
}
function apiApproveRecord(token, id, approver, note, status) {
  const user = requireAuth_(token, ['Manager']);
  return approveRecord_(id, approver || user.name, note, status);
}
function apiCloseNC(token, id, action, by) {
  const user = requireAuth_(token, ['Manager']);
  return closeNc_(id, action, by || user.name);
}

function handleApiCall_(fn, args) {
  switch (fn) {
    case 'apiLogin': return apiLogin(args[0], args[1]);
    case 'apiVerifySession': return apiVerifySession(args[0]);
    case 'apiLogout': return apiLogout(args[0]);
    case 'apiChangePassword': return apiChangePassword(args[0], args[1], args[2]);
    case 'apiGetUsers': return apiGetUsers(args[0]);
    case 'apiCreateUser': return apiCreateUser(args[0], args[1]);
    case 'apiUpdateUser': return apiUpdateUser(args[0], args[1], args[2]);
    case 'apiGetRecords': return apiGetRecords(args[0]);
    case 'apiSaveRecord': return apiSaveRecord(args[0], args[1], args[2]);
    case 'apiGetNextNo': return apiGetNextNo(args[0]);
    case 'apiGetConfig': return apiGetConfig(args[0]);
    case 'apiApproveRecord': return apiApproveRecord(args[0], args[1], args[2], args[3], args[4]);
    case 'apiCloseNC': return apiCloseNC(args[0], args[1], args[2], args[3]);
    default: throw new Error('Unknown API function: ' + fn);
  }
}
