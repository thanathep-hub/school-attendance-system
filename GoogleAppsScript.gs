// ═════════════════════════════════════════════════════════
// Google Apps Script Backend — ระบบเช็คชื่อโรงเรียน
// Deployment URL: https://script.google.com/macros/d/{SCRIPT_ID}/usercontent/v1/execute
// ═════════════════════════════════════════════════════════

// ─── Configuration ─────────────────────────────────────
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const SHEETS = {
  USERS: 'users',
  SCHOOLS: 'schools',
  STUDENTS: 'students',
  ATTENDANCE: 'attendance',
};

// ─── Helper: Get or Create Sheet ─────────────────────
function getOrCreateSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

// ─── Helper: Sheet to Array ──────────────────────────
function sheetToArray(sheetName) {
  const sheet = getOrCreateSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return [];
  
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] || ''; });
    return obj;
  });
}

// ─── Helper: Find Row Index ──────────────────────────
function findRowIndex(sheetName, columnName, value) {
  const sheet = getOrCreateSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const colIndex = data[0].indexOf(columnName);
  if (colIndex === -1) return -1;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][colIndex] === value) return i + 1;
  }
  return -1;
}

// ─── Initialize Sheets ────────────────────────────────
function initializeSheets() {
  // Create users sheet
  const usersSheet = getOrCreateSheet(SHEETS.USERS);
  if (usersSheet.getLastRow() === 0) {
    usersSheet.appendRow([
      'id', 'username', 'password', 'name', 'role', 'email', 'phone', 'color', 'active', 'created_at'
    ]);
  }

  // Create schools sheet
  const schoolsSheet = getOrCreateSheet(SHEETS.SCHOOLS);
  if (schoolsSheet.getLastRow() === 0) {
    schoolsSheet.appendRow([
      'id', 'name', 'code', 'province', 'director', 'phone', 'email'
    ]);
  }

  // Create students sheet
  const studentsSheet = getOrCreateSheet(SHEETS.STUDENTS);
  if (studentsSheet.getLastRow() === 0) {
    studentsSheet.appendRow([
      'id', 'student_code', 'name', 'nickname', 'classroom', 'gender',
      'phone', 'parent_name', 'parent_phone', 'line_id', 'birthdate', 'active', 'created_at'
    ]);
  }

  // Create attendance sheet (Phase 2.B)
  const attendanceSheet = getOrCreateSheet(SHEETS.ATTENDANCE);
  if (attendanceSheet.getLastRow() === 0) {
    attendanceSheet.appendRow([
      'id', 'date', 'period', 'student_id', 'status', 'note', 'recorded_by', 'recorded_at'
    ]);
  }

  // Seed demo data if empty
  const users = sheetToArray(SHEETS.USERS);
  if (users.length === 0) {
    seedDemoData();
  }

  // Seed students separately so we can run on already-initialized sheets
  const students = sheetToArray(SHEETS.STUDENTS);
  if (students.length === 0) {
    seedDemoStudents();
  }
}

// ─── Seed Demo Students ───────────────────────────────
function seedDemoStudents() {
  const studentsSheet = getOrCreateSheet(SHEETS.STUDENTS);
  const now = new Date().toISOString();
  const demo = [
    // ป.3/1 (ครูพรทิพย์) — 4 คน
    ['1',  '12301', 'ด.ช.ภูมินทร์ ใจดี',        'ภูมิ',  'ป.3/1', 'M', '081-100-0001', 'นางสายฝน แก้วสุข',     '089-100-0001', '', '2017-05-15', true, now],
    ['2',  '12302', 'ด.ญ.พิมพ์ลภัส วงศ์ศรี',     'พิม',   'ป.3/1', 'F', '081-100-0002', 'นางกาญจนา วงศ์ศรี',    '089-100-0002', '', '2017-07-22', true, now],
    ['3',  '12303', 'ด.ช.ธีรภัทร แสงทอง',         'เต้',   'ป.3/1', 'M', '081-100-0003', 'นายธวัช แสงทอง',        '089-100-0003', '', '2017-03-08', true, now],
    ['4',  '12304', 'ด.ญ.ชลธิชา สมบูรณ์',         'ใบเตย', 'ป.3/1', 'F', '081-100-0004', 'นางอำพร สมบูรณ์',       '089-100-0004', '', '2017-11-30', true, now],
    // ป.3/2 (ครูสุดา) — 3 คน
    ['5',  '12305', 'ด.ช.ปรัชญา ขำดี',            'แชมป์', 'ป.3/2', 'M', '081-100-0005', 'นางวิภา ขำดี',           '089-100-0005', '', '2017-08-12', true, now],
    ['6',  '12306', 'ด.ญ.กชกร เจริญสุข',          'แก้ม',  'ป.3/2', 'F', '081-100-0006', 'นางสุนิสา เจริญสุข',     '089-100-0006', '', '2017-02-19', true, now],
    ['7',  '12307', 'ด.ช.วรพล อยู่ดี',            'พล',    'ป.3/2', 'M', '081-100-0007', 'นายสมบัติ อยู่ดี',       '089-100-0007', '', '2017-09-05', true, now],
    // ป.4/1 (ครูประวิทย์) — 3 คน
    ['8',  '12401', 'ด.ช.อัครเดช พิทักษ์',        'เดช',   'ป.4/1', 'M', '081-100-0008', 'นางพรพรรณ พิทักษ์',     '089-100-0008', '', '2016-06-25', true, now],
    ['9',  '12402', 'ด.ญ.รัชนีกร ทองคำ',          'นุ่น',  'ป.4/1', 'F', '081-100-0009', 'นางวลัยพร ทองคำ',       '089-100-0009', '', '2016-12-03', true, now],
    ['10', '12403', 'ด.ช.ณัฐภัทร สุขเกษม',        'เน็ต',  'ป.4/1', 'M', '081-100-0010', 'นายชาญณรงค์ สุขเกษม',   '089-100-0010', '', '2016-04-17', true, now],
  ];
  demo.forEach(row => studentsSheet.appendRow(row));
}

// ─── Seed Demo Data ───────────────────────────────────
function seedDemoData() {
  const usersSheet = getOrCreateSheet(SHEETS.USERS);
  const demoUsers = [
    ['1', 'somying', Utilities.base64Encode('1234'), 'นางสมหญิง ใจดี', 'admin', 'somying@school.th', '081-xxx-1000', '#B58A2B', true, new Date().toISOString()],
    ['2', 'porntip', Utilities.base64Encode('1234'), 'ครูพรทิพย์ สวัสดี', 'teacher', 'porntip@school.th', '081-xxx-1001', '#1E47A3', true, new Date().toISOString()],
    ['3', 'suda', Utilities.base64Encode('1234'), 'ครูสุดา แสงดี', 'teacher', 'suda@school.th', '081-xxx-1002', '#B91C1C', true, new Date().toISOString()],
    ['4', 'pravit', Utilities.base64Encode('1234'), 'ครูประวิทย์ มากมี', 'teacher', 'pravit@school.th', '081-xxx-1003', '#15803D', true, new Date().toISOString()],
  ];
  demoUsers.forEach(user => usersSheet.appendRow(user));

  const schoolsSheet = getOrCreateSheet(SHEETS.SCHOOLS);
  schoolsSheet.appendRow([
    '1', 'โรงเรียนบ้านหนองบัว', '1030200001', 'นครราชสีมา', 'นางสมหญิง ใจดี', '0-4422-xxx-xxx', 'nongbua.school@go.th'
  ]);
}

// ─── API: GET (health check) ───────────────────────
// เปิด URL ที่ deploy ใน browser แล้วต้องเห็น JSON {"ok":true,...}
// ถ้าเปิดแล้วเห็น error คือ deploy ยังไม่สำเร็จ
function doGet(e) {
  try {
    initializeSheets();
    return formatResponse(true, 'API is running', {
      version: '2.0.0',
      spreadsheetId: SPREADSHEET_ID ? 'set' : 'NOT_SET',
      time: new Date().toISOString(),
    });
  } catch (error) {
    return formatResponse(false, 'Error: ' + error.message);
  }
}

// ─── API: POST (main entry point) ──────────────────
function doPost(e) {
  try {
    initializeSheets();

    if (!e || !e.postData || !e.postData.contents) {
      return formatResponse(false, 'Missing request body');
    }

    const request = JSON.parse(e.postData.contents);
    const { action } = request;

    if (action === 'login') {
      return handleLogin(request.username, request.password, request.role);
    } else if (action === 'getSchool') {
      return handleGetSchool();
    } else if (action === 'verifyToken') {
      return handleVerifyToken(request.token);
    } else if (action === 'getUsers') {
      return handleGetUsers(request.token);
    } else if (action === 'createUser') {
      return handleCreateUser(request.token, request.data);
    } else if (action === 'updateUser') {
      return handleUpdateUser(request.token, request.userId, request.data);
    } else if (action === 'deleteUser') {
      return handleDeleteUser(request.token, request.userId);
    } else if (action === 'getStudents') {
      return handleGetStudents(request.token, request.classroom);
    } else if (action === 'createStudent') {
      return handleCreateStudent(request.token, request.data);
    } else if (action === 'updateStudent') {
      return handleUpdateStudent(request.token, request.studentId, request.data);
    } else if (action === 'deleteStudent') {
      return handleDeleteStudent(request.token, request.studentId, request.hardDelete);
    } else if (action === 'getAttendance') {
      return handleGetAttendance(request.token, request.date, request.classroom, request.period);
    } else if (action === 'markAttendance') {
      return handleMarkAttendance(request.token, request.data);
    } else if (action === 'bulkMarkAttendance') {
      return handleBulkMarkAttendance(request.token, request.data);
    } else {
      return formatResponse(false, 'Unknown action: ' + action);
    }
  } catch (error) {
    return formatResponse(false, 'Error: ' + error.message);
  }
}

// ─── Helper: เช็คว่า field active เป็น "เปิดใช้งาน" หรือไม่ ─
// Google Sheets อาจเก็บเป็น boolean true, string "TRUE", string "true" ก็ได้
function isActive(v) {
  if (v === true) return true;
  if (typeof v === 'string') {
    const s = v.toString().trim().toUpperCase();
    return s === 'TRUE' || s === '1' || s === 'YES';
  }
  return false;
}

// ─── Handler: Login ───────────────────────────────────
function handleLogin(username, password, role) {
  if (!username || !password) {
    return formatResponse(false, 'Missing username or password');
  }

  const users = sheetToArray(SHEETS.USERS);
  const hashedInput = Utilities.base64Encode(password);

  // เช็คทุกเงื่อนไขแบบ explicit (อย่ารวมกับ || เพราะ precedence ผิดได้ง่าย)
  const user = users.find(function(u) {
    const matchUser = (u.username === username);
    const matchPass = (u.password === hashedInput);
    const matchRole = role ? (u.role === role) : true;
    const isOn = isActive(u.active);
    return matchUser && matchPass && matchRole && isOn;
  });

  if (!user) {
    return formatResponse(false, 'Invalid credentials');
  }

  // Generate simple token (ใส่ timestamp ไว้ตรวจวันหมดอายุได้)
  const token = Utilities.base64Encode(user.id + ':' + user.username + ':' + new Date().getTime());

  return formatResponse(true, 'Login successful', {
    token: token,
    user: {
      id: user.id,
      username: user.username,
      name: user.name,
      role: user.role,
      email: user.email,
    }
  });
}

// ─── Handler: Verify Token ─────────────────────────
// token หมดอายุใน 7 วัน
function handleVerifyToken(token) {
  if (!token) return formatResponse(false, 'No token');
  try {
    const decoded = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    const parts = decoded.split(':');
    if (parts.length !== 3) return formatResponse(false, 'Invalid token');
    const id = parts[0], username = parts[1], ts = parseInt(parts[2], 10);
    const sevenDays = 7 * 24 * 60 * 60 * 1000;
    if (isNaN(ts) || (new Date().getTime() - ts) > sevenDays) {
      return formatResponse(false, 'Token expired');
    }
    const users = sheetToArray(SHEETS.USERS);
    const user = users.find(function(u) { return String(u.id) === String(id) && u.username === username && isActive(u.active); });
    if (!user) return formatResponse(false, 'User not found');
    return formatResponse(true, 'Token valid', {
      user: { id: user.id, username: user.username, name: user.name, role: user.role, email: user.email }
    });
  } catch (err) {
    return formatResponse(false, 'Invalid token: ' + err.message);
  }
}

// ─── Handler: Get School Info ──────────────────────
function handleGetSchool() {
  const schools = sheetToArray(SHEETS.SCHOOLS);
  return formatResponse(true, 'School data', schools[0] || {});
}

// ─── Handler: Get All Users ───────────────────────
function handleGetUsers(token) {
  // Verify token is admin (simplified)
  if (!token) return formatResponse(false, 'Unauthorized');

  const users = sheetToArray(SHEETS.USERS);
  const safeUsers = users.map(u => ({
    id: u.id,
    username: u.username,
    name: u.name,
    role: u.role,
    email: u.email,
    phone: u.phone,
    active: u.active,
    created_at: u.created_at,
  }));

  return formatResponse(true, 'Users list', safeUsers);
}

// ─── Handler: Create User ──────────────────────────
function handleCreateUser(token, data) {
  if (!token) return formatResponse(false, 'Unauthorized');

  const { username, password, name, role, email, phone } = data;
  if (!username || !password || !name || !role) {
    return formatResponse(false, 'Missing required fields');
  }

  const usersSheet = getOrCreateSheet(SHEETS.USERS);
  const users = sheetToArray(SHEETS.USERS);
  const newId = Math.max(...users.map(u => parseInt(u.id) || 0)) + 1;

  usersSheet.appendRow([
    newId,
    username,
    Utilities.base64Encode(password),
    name,
    role,
    email || '',
    phone || '',
    '#1E47A3',
    true,
    new Date().toISOString()
  ]);

  return formatResponse(true, 'User created', { id: newId });
}

// ─── Handler: Update User ──────────────────────────
function handleUpdateUser(token, userId, data) {
  if (!token) return formatResponse(false, 'Unauthorized');

  const { name, email, phone, active } = data;
  const rowIndex = findRowIndex(SHEETS.USERS, 'id', userId);
  
  if (rowIndex === -1) return formatResponse(false, 'User not found');

  const usersSheet = getOrCreateSheet(SHEETS.USERS);
  const row = usersSheet.getRange(rowIndex, 1, 1, 10).getValues()[0];
  
  if (name) row[3] = name;
  if (email !== undefined) row[5] = email;
  if (phone !== undefined) row[6] = phone;
  if (active !== undefined) row[7] = active;

  usersSheet.getRange(rowIndex, 1, 1, 10).setValues([row]);
  return formatResponse(true, 'User updated');
}

// ─── Handler: Delete User ──────────────────────────
function handleDeleteUser(token, userId) {
  if (!token) return formatResponse(false, 'Unauthorized');

  const rowIndex = findRowIndex(SHEETS.USERS, 'id', userId);
  if (rowIndex === -1) return formatResponse(false, 'User not found');

  const usersSheet = getOrCreateSheet(SHEETS.USERS);
  usersSheet.deleteRow(rowIndex);

  return formatResponse(true, 'User deleted');
}

// ═══════════════════════════════════════════════════════
// Students CRUD Handlers (Phase 2.A)
// ═══════════════════════════════════════════════════════

// ─── Handler: Get Students ─────────────────────────
// Returns all active students; optionally filter by classroom (e.g. "ป.3/1")
function handleGetStudents(token, classroom) {
  if (!token) return formatResponse(false, 'Unauthorized');
  const all = sheetToArray(SHEETS.STUDENTS);
  let filtered = all.filter(function(s) { return isActive(s.active); });
  if (classroom) filtered = filtered.filter(function(s) { return s.classroom === classroom; });
  return formatResponse(true, 'Students list', filtered);
}

// ─── Handler: Create Student ───────────────────────
function handleCreateStudent(token, data) {
  if (!token) return formatResponse(false, 'Unauthorized');
  if (!data) return formatResponse(false, 'Missing data');
  const { student_code, name, nickname, classroom, gender, phone, parent_name, parent_phone, line_id, birthdate } = data;
  if (!name || !classroom) return formatResponse(false, 'Missing required fields: name, classroom');

  const studentsSheet = getOrCreateSheet(SHEETS.STUDENTS);
  const students = sheetToArray(SHEETS.STUDENTS);
  const newId = Math.max(0, ...students.map(function(s) { return parseInt(s.id) || 0; })) + 1;

  studentsSheet.appendRow([
    newId,
    student_code || '',
    name,
    nickname || '',
    classroom,
    gender || '',
    phone || '',
    parent_name || '',
    parent_phone || '',
    line_id || '',
    birthdate || '',
    true,
    new Date().toISOString()
  ]);
  return formatResponse(true, 'Student created', { id: newId });
}

// ─── Handler: Update Student ───────────────────────
// Column order: 0=id, 1=student_code, 2=name, 3=nickname, 4=classroom, 5=gender,
//               6=phone, 7=parent_name, 8=parent_phone, 9=line_id, 10=birthdate, 11=active, 12=created_at
function handleUpdateStudent(token, studentId, data) {
  if (!token) return formatResponse(false, 'Unauthorized');
  if (!studentId) return formatResponse(false, 'Missing studentId');
  if (!data) return formatResponse(false, 'Missing data');

  const rowIndex = findRowIndex(SHEETS.STUDENTS, 'id', studentId);
  if (rowIndex === -1) return formatResponse(false, 'Student not found');

  const studentsSheet = getOrCreateSheet(SHEETS.STUDENTS);
  const row = studentsSheet.getRange(rowIndex, 1, 1, 13).getValues()[0];

  if (data.student_code !== undefined) row[1] = data.student_code;
  if (data.name !== undefined)         row[2] = data.name;
  if (data.nickname !== undefined)     row[3] = data.nickname;
  if (data.classroom !== undefined)    row[4] = data.classroom;
  if (data.gender !== undefined)       row[5] = data.gender;
  if (data.phone !== undefined)        row[6] = data.phone;
  if (data.parent_name !== undefined)  row[7] = data.parent_name;
  if (data.parent_phone !== undefined) row[8] = data.parent_phone;
  if (data.line_id !== undefined)      row[9] = data.line_id;
  if (data.birthdate !== undefined)    row[10] = data.birthdate;
  if (data.active !== undefined)       row[11] = data.active;

  studentsSheet.getRange(rowIndex, 1, 1, 13).setValues([row]);
  return formatResponse(true, 'Student updated');
}

// ─── Handler: Delete Student (soft delete) ─────────
// Sets active=false so data is preserved (attendance history needs reference)
// To hard-delete: set request.hardDelete = true
function handleDeleteStudent(token, studentId, hardDelete) {
  if (!token) return formatResponse(false, 'Unauthorized');
  if (!studentId) return formatResponse(false, 'Missing studentId');

  const rowIndex = findRowIndex(SHEETS.STUDENTS, 'id', studentId);
  if (rowIndex === -1) return formatResponse(false, 'Student not found');

  const studentsSheet = getOrCreateSheet(SHEETS.STUDENTS);
  if (hardDelete === true) {
    studentsSheet.deleteRow(rowIndex);
    return formatResponse(true, 'Student deleted');
  }
  // Soft delete: set active (column 12, 1-indexed) to false
  studentsSheet.getRange(rowIndex, 12).setValue(false);
  return formatResponse(true, 'Student deactivated');
}

// ═══════════════════════════════════════════════════════
// Attendance Handlers (Phase 2.B)
// ═══════════════════════════════════════════════════════

// ─── Helper: Extract user_id from token (no expiry check) ────
function tokenToUserId(token) {
  if (!token) return null;
  try {
    const decoded = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    const parts = decoded.split(':');
    if (parts.length !== 3) return null;
    return parts[0];
  } catch (e) { return null; }
}

// ─── Helper: Find attendance row by composite key (date, period, student_id) ──
// Columns: 0=id, 1=date, 2=period, 3=student_id, 4=status, 5=note, 6=recorded_by, 7=recorded_at
function findAttendanceRowIndex(date, period, studentId) {
  const sheet = getOrCreateSheet(SHEETS.ATTENDANCE);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return -1;
  const p = String(period || '');
  for (let i = 1; i < data.length; i++) {
    // Date in sheet may be Date object or string; normalize to YYYY-MM-DD prefix match
    const rowDate = data[i][1] instanceof Date
      ? Utilities.formatDate(data[i][1], 'GMT', 'yyyy-MM-dd')
      : String(data[i][1]).substring(0, 10);
    if (rowDate === date && String(data[i][2]) === p && String(data[i][3]) === String(studentId)) {
      return i + 1;
    }
  }
  return -1;
}

// ─── Handler: Get Attendance ─────────────────────────
// Returns attendance records for a given date. Optional filter by classroom + period.
// Each record is enriched with student info (name, nickname, classroom) for UI convenience.
function handleGetAttendance(token, date, classroom, period) {
  if (!token) return formatResponse(false, 'Unauthorized');
  if (!date) return formatResponse(false, 'Missing date (YYYY-MM-DD)');

  const all = sheetToArray(SHEETS.ATTENDANCE);
  const p = (period !== undefined && period !== null) ? String(period) : null;

  // Filter by date (string startsWith allows date or full ISO)
  let recs = all.filter(function(a) {
    const aDate = a.date instanceof Date
      ? Utilities.formatDate(a.date, 'GMT', 'yyyy-MM-dd')
      : String(a.date).substring(0, 10);
    if (aDate !== date) return false;
    if (p !== null && String(a.period) !== p) return false;
    return true;
  });

  // Build student lookup
  const students = sheetToArray(SHEETS.STUDENTS);
  const studentMap = {};
  students.forEach(function(s) { studentMap[String(s.id)] = s; });

  // Optional classroom filter
  if (classroom) {
    recs = recs.filter(function(a) {
      const s = studentMap[String(a.student_id)];
      return s && s.classroom === classroom;
    });
  }

  // Enrich
  const enriched = recs.map(function(a) {
    const s = studentMap[String(a.student_id)] || {};
    return {
      id: a.id,
      date: (a.date instanceof Date ? Utilities.formatDate(a.date, 'GMT', 'yyyy-MM-dd') : String(a.date).substring(0,10)),
      period: a.period || '',
      student_id: a.student_id,
      student_name: s.name || '',
      student_nickname: s.nickname || '',
      classroom: s.classroom || '',
      status: a.status,
      note: a.note || '',
      recorded_by: a.recorded_by,
      recorded_at: a.recorded_at,
    };
  });

  return formatResponse(true, 'Attendance list', enriched);
}

// ─── Handler: Mark Attendance (single, upsert) ────────
// data: { student_id, date, status, note?, period? }
// status must be one of: present, absent, leave, late
function handleMarkAttendance(token, data) {
  if (!token) return formatResponse(false, 'Unauthorized');
  if (!data) return formatResponse(false, 'Missing data');
  const { student_id, date, status, note, period } = data;
  if (!student_id || !date || !status) return formatResponse(false, 'Missing required: student_id, date, status');

  const validStatuses = ['present', 'absent', 'leave', 'late'];
  if (validStatuses.indexOf(status) === -1) return formatResponse(false, 'Invalid status (must be present/absent/leave/late)');

  const sheet = getOrCreateSheet(SHEETS.ATTENDANCE);
  const userId = tokenToUserId(token) || '';
  const now = new Date().toISOString();
  const p = period || '';
  const rowIndex = findAttendanceRowIndex(date, p, student_id);

  if (rowIndex !== -1) {
    const row = sheet.getRange(rowIndex, 1, 1, 8).getValues()[0];
    row[4] = status;
    if (note !== undefined) row[5] = note;
    row[6] = userId;
    row[7] = now;
    sheet.getRange(rowIndex, 1, 1, 8).setValues([row]);
    return formatResponse(true, 'Attendance updated', { id: row[0], action: 'updated' });
  }

  const all = sheetToArray(SHEETS.ATTENDANCE);
  const newId = Math.max(0, ...all.map(function(a) { return parseInt(a.id) || 0; })) + 1;
  sheet.appendRow([newId, date, p, student_id, status, note || '', userId, now]);
  return formatResponse(true, 'Attendance created', { id: newId, action: 'created' });
}

// ─── Handler: Bulk Mark Attendance ──────────────────
// data: { date, period?, records: [{student_id, status, note?}] }
// Upserts each row efficiently. Returns counts.
function handleBulkMarkAttendance(token, data) {
  if (!token) return formatResponse(false, 'Unauthorized');
  if (!data) return formatResponse(false, 'Missing data');
  const { date, period, records } = data;
  if (!date || !Array.isArray(records)) return formatResponse(false, 'Missing required: date, records[]');

  const validStatuses = ['present', 'absent', 'leave', 'late'];
  const sheet = getOrCreateSheet(SHEETS.ATTENDANCE);
  const userId = tokenToUserId(token) || '';
  const now = new Date().toISOString();
  const p = period || '';

  // Build index map for fast upsert
  const sheetData = sheet.getDataRange().getValues();
  const indexMap = {};
  for (let i = 1; i < sheetData.length; i++) {
    const rowDate = sheetData[i][1] instanceof Date
      ? Utilities.formatDate(sheetData[i][1], 'GMT', 'yyyy-MM-dd')
      : String(sheetData[i][1]).substring(0, 10);
    const key = rowDate + '|' + String(sheetData[i][2]) + '|' + String(sheetData[i][3]);
    indexMap[key] = i + 1;
  }

  let nextId = 1;
  for (let i = 1; i < sheetData.length; i++) {
    const id = parseInt(sheetData[i][0]) || 0;
    if (id >= nextId) nextId = id + 1;
  }

  let created = 0, updated = 0, skipped = 0;
  records.forEach(function(r) {
    if (!r.student_id || !r.status || validStatuses.indexOf(r.status) === -1) { skipped++; return; }
    const key = date + '|' + p + '|' + String(r.student_id);
    const rowIndex = indexMap[key];
    if (rowIndex) {
      const row = sheet.getRange(rowIndex, 1, 1, 8).getValues()[0];
      row[4] = r.status;
      if (r.note !== undefined) row[5] = r.note;
      row[6] = userId;
      row[7] = now;
      sheet.getRange(rowIndex, 1, 1, 8).setValues([row]);
      updated++;
    } else {
      sheet.appendRow([nextId, date, p, r.student_id, r.status, r.note || '', userId, now]);
      indexMap[key] = sheet.getLastRow();
      nextId++;
      created++;
    }
  });

  return formatResponse(true, 'Bulk attendance saved', { created: created, updated: updated, skipped: skipped, total: records.length });
}

// ─── Helper: Format Response ───────────────────────
function formatResponse(success, message, data = null) {
  return ContentService
    .createTextOutput(JSON.stringify({
      success,
      message,
      data: data || {}
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Deployed Endpoint ────────────────────────────
// POST request to: https://script.google.com/macros/d/{SCRIPT_ID}/usercontent/v1/execute
// Body: { "action": "login", "username": "somying", "password": "1234", "role": "admin" }
