// ============================================================
//  HRMS PORTAL — Code.gs  (Server-Side Entry Point)
//  Google Apps Script + Google Sheets Backend
// ============================================================

// ── Sheet names ──────────────────────────────────────────────
const SHEET = {
  EMPLOYEES : 'Employees',
  ATTENDANCE: 'Attendance',
  LEAVES    : 'Leaves',
  SESSIONS  : 'Sessions',
  SETTINGS  : 'Settings'
};

// ── Web-App Entry Points ──────────────────────────────────────

function doGet(e) {

  e = e || {};

  const page = e.parameter ? e.parameter.page : '';
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('HRMS Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Spreadsheet helpers ───────────────────────────────────────
function getSpreadsheet() {
  // Reads ID from Script Properties; falls back to creating a new sheet
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  if (!id) {
    const ss = SpreadsheetApp.create('HRMS Database');
    id = ss.getId();
    props.setProperty('SHEET_ID', id);
    initializeSheets(ss);
  }
  return SpreadsheetApp.openById(id);
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    seedHeaders(sh, name);
  }
  return sh;
}

function seedHeaders(sh, name) {
  const headers = {
    Employees : ['ID','Name','Email','Password','Department','Position','Phone','JoinDate','Status','Role','Avatar'],
    Attendance: ['ID','EmployeeID','Date','ClockIn','ClockOut','Hours','Status'],
    Leaves    : ['ID','EmployeeID','EmployeeName','Type','StartDate','EndDate','Days','Reason','Status','AppliedOn','ReviewedBy','ReviewedOn'],
    Sessions  : ['Token','EmployeeID','Role','CreatedAt','ExpiresAt'],
    Settings  : ['Key','Value']
  };
  if (headers[name]) sh.appendRow(headers[name]);
}

function initializeSheets(ss) {
  Object.values(SHEET).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    seedHeaders(sh, name);
  });
  // Seed default admin
  const empSh = ss.getSheetByName(SHEET.EMPLOYEES);
  empSh.appendRow([
    'EMP001','Admin User','admin@hrms.com',hashPassword('admin123'),
    'Management','HR Administrator','+1-000-000-0000',
    new Date().toISOString().split('T')[0],'Active','admin',''
  ]);
  // Seed demo employee
  empSh.appendRow([
    'EMP002','Jane Smith','jane@hrms.com',hashPassword('employee123'),
    'Engineering','Software Engineer','+1-555-123-4567',
    new Date().toISOString().split('T')[0],'Active','employee',''
  ]);
  // Settings defaults
  const setsSh = ss.getSheetByName(SHEET.SETTINGS);
  setsSh.appendRow(['company_name','HRMS Corp']);
  setsSh.appendRow(['work_hours','8']);
  setsSh.appendRow(['leave_quota','20']);
}

// ── Auth ──────────────────────────────────────────────────────
function hashPassword(pw) {
  return Utilities.base64Encode(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, pw
  ));
}

function login(email, password) {
  try {
    const sh    = getSheet(SHEET.EMPLOYEES);
    const data  = sh.getDataRange().getValues();
    const hashed = hashPassword(password);
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[2].toString().toLowerCase() === email.toLowerCase() && row[3] === hashed) {
        if (row[8] !== 'Active') return { success: false, message: 'Account is inactive.' };
        const token = createSession(row[0], row[9]);
        return {
          success: true,
          token,
          role  : row[9],
          name  : row[1],
          id    : row[0],
          dept  : row[4],
          pos   : row[5]
        };
      }
    }
    return { success: false, message: 'Invalid email or password.' };
  } catch(e) { return { success: false, message: e.message }; }
}

function createSession(empId, role) {
  const token   = Utilities.getUuid();
  const now     = new Date();
  const expires = new Date(now.getTime() + 8 * 3600 * 1000); // 8h
  getSheet(SHEET.SESSIONS).appendRow([token, empId, role, now.toISOString(), expires.toISOString()]);
  // Purge old sessions
  purgeExpiredSessions();
  return token;
}

function validateSession(token) {
  if (!token) return null;
  const sh   = getSheet(SHEET.SESSIONS);
  const data = sh.getDataRange().getValues();
  const now  = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token && new Date(data[i][4]) > now) {
      return { empId: data[i][1], role: data[i][2] };
    }
  }
  return null;
}

function logout(token) {
  const sh   = getSheet(SHEET.SESSIONS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) { sh.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

function purgeExpiredSessions() {
  const sh   = getSheet(SHEET.SESSIONS);
  const data = sh.getDataRange().getValues();
  const now  = new Date();
  for (let i = data.length - 1; i >= 1; i--) {
    if (new Date(data[i][4]) < now) sh.deleteRow(i + 1);
  }
}

// ── Dashboard ─────────────────────────────────────────────────
function getDashboardData(token) {
  const sess = validateSession(token);
  if (!sess) return { error: 'Unauthorized' };
  const today = new Date().toISOString().split('T')[0];

  const empSh  = getSheet(SHEET.EMPLOYEES);
  const attSh  = getSheet(SHEET.ATTENDANCE);
  const lvSh   = getSheet(SHEET.LEAVES);

  const empData = empSh.getDataRange().getValues();
  const attData = attSh.getDataRange().getValues();
  const lvData  = lvSh.getDataRange().getValues();

  const totalEmployees  = empData.slice(1).filter(r => r[8] === 'Active').length;
  const presentToday    = attData.slice(1).filter(r => r[2] === today).length;
  const pendingLeaves   = lvData.slice(1).filter(r => r[8] === 'Pending').length;
  const onLeaveToday    = lvData.slice(1).filter(r =>
    r[8] === 'Approved' && r[4] <= today && r[5] >= today
  ).length;

  // Recent attendance (last 10)
  const recentAtt = attData.slice(1).reverse().slice(0, 10).map(r => ({
    empId: r[1], date: r[2], clockIn: r[3], clockOut: r[4], hours: r[5], status: r[6]
  }));

  // Dept breakdown
  const depts = {};
  empData.slice(1).filter(r => r[8] === 'Active').forEach(r => {
    depts[r[4]] = (depts[r[4]] || 0) + 1;
  });

  // My attendance for employee role
  let myAtt = null;
  if (sess.role === 'employee') {
    const todayRec = attData.slice(1).find(r => r[1] === sess.empId && r[2] === today);
    myAtt = todayRec ? { clockIn: todayRec[3], clockOut: todayRec[4], hours: todayRec[5] } : null;
  }

  return {
    totalEmployees, presentToday, pendingLeaves, onLeaveToday,
    recentAtt, depts, myAtt, role: sess.role, empId: sess.empId
  };
}

// ── Attendance ────────────────────────────────────────────────
function clockIn(token) {
  const sess = validateSession(token);
  if (!sess) return { error: 'Unauthorized' };
  const today = new Date().toISOString().split('T')[0];
  const now   = new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
  const sh    = getSheet(SHEET.ATTENDANCE);
  const data  = sh.getDataRange().getValues();
  const existing = data.slice(1).find(r => r[1] === sess.empId && r[2] === today);
  if (existing) return { success: false, message: 'Already clocked in today.' };
  const id = 'ATT' + Date.now();
  sh.appendRow([id, sess.empId, today, now, '', '', 'Present']);
  return { success: true, time: now };
}

function clockOut(token) {
  const sess = validateSession(token);
  if (!sess) return { error: 'Unauthorized' };
  const today = new Date().toISOString().split('T')[0];
  const now   = new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
  const sh    = getSheet(SHEET.ATTENDANCE);
  const data  = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === sess.empId && data[i][2] === today) {
      if (data[i][4]) return { success: false, message: 'Already clocked out.' };
      const hrs = calcHours(data[i][3], now);
      sh.getRange(i + 1, 5).setValue(now);
      sh.getRange(i + 1, 6).setValue(hrs);
      return { success: true, time: now, hours: hrs };
    }
  }
  return { success: false, message: 'No clock-in record found for today.' };
}

function calcHours(inTime, outTime) {
  const [ih, im] = inTime.split(':').map(Number);
  const [oh, om] = outTime.split(':').map(Number);
  const diff = (oh * 60 + om) - (ih * 60 + im);
  return Math.max(0, (diff / 60).toFixed(2));
}

function getAttendanceHistory(token, empId) {
  const sess = validateSession(token);
  if (!sess) return { error: 'Unauthorized' };
  const targetId = (sess.role === 'admin' && empId) ? empId : sess.empId;
  const sh   = getSheet(SHEET.ATTENDANCE);
  const data = sh.getDataRange().getValues();
  const records = data.slice(1)
    .filter(r => r[1] === targetId)
    .reverse()
    .slice(0, 30)
    .map(r => ({ id: r[0], empId: r[1], date: r[2], clockIn: r[3], clockOut: r[4], hours: r[5], status: r[6] }));
  return { success: true, records };
}

function getAllAttendance(token, dateFilter) {
  const sess = validateSession(token);
  if (!sess || sess.role !== 'admin') return { error: 'Unauthorized' };
  const sh   = getSheet(SHEET.ATTENDANCE);
  const data = sh.getDataRange().getValues();
  let records = data.slice(1).map(r => ({
    id: r[0], empId: r[1], date: r[2], clockIn: r[3], clockOut: r[4], hours: r[5], status: r[6]
  }));
  if (dateFilter) records = records.filter(r => r.date === dateFilter);
  return { success: true, records: records.reverse().slice(0, 100) };
}

// ── Leaves ────────────────────────────────────────────────────
function applyLeave(token, data) {
  const sess = validateSession(token);
  if (!sess) return { error: 'Unauthorized' };
  const empSh = getSheet(SHEET.EMPLOYEES);
  const empData = empSh.getDataRange().getValues();
  const emp = empData.slice(1).find(r => r[0] === sess.empId);
  const empName = emp ? emp[1] : sess.empId;
  const days = calcDays(data.startDate, data.endDate);
  const id = 'LV' + Date.now();
  getSheet(SHEET.LEAVES).appendRow([
    id, sess.empId, empName, data.type, data.startDate, data.endDate,
    days, data.reason, 'Pending', new Date().toISOString().split('T')[0], '', ''
  ]);
  return { success: true, message: 'Leave request submitted.' };
}

function calcDays(start, end) {
  const s = new Date(start), e = new Date(end);
  return Math.max(1, Math.ceil((e - s) / 86400000) + 1);
}

function getLeaves(token, empId) {
  const sess = validateSession(token);
  if (!sess) return { error: 'Unauthorized' };
  const sh   = getSheet(SHEET.LEAVES);
  const data = sh.getDataRange().getValues();
  let records = data.slice(1).map(r => ({
    id: r[0], empId: r[1], empName: r[2], type: r[3],
    startDate: r[4], endDate: r[5], days: r[6], reason: r[7],
    status: r[8], appliedOn: r[9], reviewedBy: r[10], reviewedOn: r[11]
  }));
  if (sess.role === 'employee') records = records.filter(r => r.empId === sess.empId);
  else if (empId) records = records.filter(r => r.empId === empId);
  return { success: true, records: records.reverse().slice(0, 50) };
}

function reviewLeave(token, leaveId, action, comment) {
  const sess = validateSession(token);
  if (!sess || sess.role !== 'admin') return { error: 'Unauthorized' };
  const sh   = getSheet(SHEET.LEAVES);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === leaveId) {
      sh.getRange(i + 1, 9).setValue(action === 'approve' ? 'Approved' : 'Rejected');
      sh.getRange(i + 1, 11).setValue(sess.empId);
      sh.getRange(i + 1, 12).setValue(new Date().toISOString().split('T')[0]);
      return { success: true };
    }
  }
  return { success: false, message: 'Leave not found.' };
}

// ── Employees ─────────────────────────────────────────────────
function getEmployees(token) {

  const ss = SpreadsheetApp.openById("1EIj7TUpVUzh8GnI242_a2tGvNH0e52Z_cg9UugVUqls");

  const sheet = ss.getSheetByName("Employees");

  if (!sheet) {
    return {
      success: false,
      message: "Employees sheet not found"
    };
  }

  const data = sheet.getDataRange().getValues();

  let employees = [];

  for (let i = 1; i < data.length; i++) {

    employees.push({

      id: data[i][0] || "",
      name: data[i][1] || "",
      email: data[i][2] || "",
      password: data[i][3] || "",
      dept: data[i][4] || "",
      position: data[i][5] || "",
      phone: data[i][6] || "",
      joinDate: data[i][7] || "",
      role: data[i][8] || "",
      status: data[i][9] || ""

    });
  }

  return {
    success: true,
    employees: employees
  };
}

function addEmployee(token, emp) {
  const sess = validateSession(token);
  if (!sess || sess.role !== 'admin') return { error: 'Unauthorized' };
  const sh   = getSheet(SHEET.EMPLOYEES);
  const id   = 'EMP' + String(Date.now()).slice(-6);
  sh.appendRow([
    id, emp.name, emp.email, hashPassword(emp.password || 'Welcome@123'),
    emp.dept, emp.position, emp.phone || '', emp.joinDate || new Date().toISOString().split('T')[0],
    'Active', emp.role || 'employee', ''
  ]);
  return { success: true, id, message: 'Employee added successfully.' };
}

function updateEmployee(token, emp) {
  const sess = validateSession(token);
  if (!sess || sess.role !== 'admin') return { error: 'Unauthorized' };
  const sh   = getSheet(SHEET.EMPLOYEES);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === emp.id) {
      sh.getRange(i + 1, 2).setValue(emp.name      || data[i][1]);
      sh.getRange(i + 1, 3).setValue(emp.email     || data[i][2]);
      sh.getRange(i + 1, 5).setValue(emp.dept      || data[i][4]);
      sh.getRange(i + 1, 6).setValue(emp.position  || data[i][5]);
      sh.getRange(i + 1, 7).setValue(emp.phone     || data[i][6]);
      sh.getRange(i + 1, 9).setValue(emp.status    || data[i][8]);
      sh.getRange(i + 1, 10).setValue(emp.role     || data[i][9]);
      if (emp.password) sh.getRange(i + 1, 4).setValue(hashPassword(emp.password));
      return { success: true, message: 'Employee updated.' };
    }
  }
  return { success: false, message: 'Employee not found.' };
}

function deleteEmployee(token, empId) {
  const sess = validateSession(token);
  if (!sess || sess.role !== 'admin') return { error: 'Unauthorized' };
  const sh   = getSheet(SHEET.EMPLOYEES);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === empId) { sh.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false, message: 'Employee not found.' };
}

function getMyProfile(token) {
  const sess = validateSession(token);
  if (!sess) return { error: 'Unauthorized' };
  const sh   = getSheet(SHEET.EMPLOYEES);
  const data = sh.getDataRange().getValues();
  const row  = data.slice(1).find(r => r[0] === sess.empId);
  if (!row) return { error: 'Not found' };
  return {
    success: true,
    profile: { id: row[0], name: row[1], email: row[2], dept: row[4], position: row[5], phone: row[6], joinDate: row[7], status: row[8], role: row[9] }
  };
}

function getSettings(token) {
  const sess = validateSession(token);
  if (!sess || sess.role !== 'admin') return { error: 'Unauthorized' };
  const sh   = getSheet(SHEET.SETTINGS);
  const data = sh.getDataRange().getValues();
  const obj  = {};
  data.slice(1).forEach(r => { obj[r[0]] = r[1]; });
  return { success: true, settings: obj };
}
function resetJanePassword() {
  const sh = getSheet(SHEET.EMPLOYEES);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === 'jane@hrms.com') {
      sh.getRange(i + 1, 4).setValue(hashPassword('employee123'));
    }
    if (data[i][2] === 'admin@hrms.com') {
      sh.getRange(i + 1, 4).setValue(hashPassword('admin123'));
    }
  }
}
function forceInit() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('SHEET_ID'); // Clear old broken ID
  const ss = SpreadsheetApp.create('HRMS Database');
  props.setProperty('SHEET_ID', ss.getId());
  initializeSheets(ss);
  Logger.log('New Sheet ID: ' + ss.getId());
}
function generateEmployeeId() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Employees");
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return "EMP001";

  const lastId = sheet.getRange(lastRow, 1).getValue();
  const num = parseInt(lastId.replace("EMP", "")) + 1;

  return "EMP" + String(num).padStart(3, "0");
}
function addEmployee(token, emp) {

  const sheet = SpreadsheetApp.getActive().getSheetByName("Employees");

  const empId = generateEmployeeId();

  sheet.appendRow([
    empId,
    emp.name,
    emp.email,
    emp.password,
    emp.dept,
    emp.position,
    emp.phone,
    emp.joinDate,
    emp.role,
    "Active"
  ]);

  return {
    success: true,
    message: "Employee added successfully.",
    empId: empId,
    email: emp.email,
    password: emp.password
  };
}
function login(email, password) {

  const sheet = SpreadsheetApp.getActive().getSheetByName("Employees");

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    if (
      row[2] == email &&
      row[3] == password &&
      row[9] == "Active"
    ) {

      return {
        success: true,
        token: Utilities.getUuid(),
        id: row[0],
        name: row[1],
        dept: row[4],
        pos: row[5],
        role: row[8]
      };
    }
  }

  return {
    success: false,
    message: "Invalid email or password"
  };
}
function getDashboardData(token) {

  const ss = SpreadsheetApp.getActive();

  const empSheet = ss.getSheetByName("Employees");
  const attSheet = ss.getSheetByName("Attendance");
  const leaveSheet = ss.getSheetByName("Leaves");

  const empData = empSheet.getDataRange().getValues();
  const attData = attSheet.getDataRange().getValues();
  const leaveData = leaveSheet.getDataRange().getValues();

  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  let totalEmployees = empData.length - 1;
  let presentToday = 0;
  let pendingLeaves = 0;
  let onLeaveToday = 0;

  let recentAtt = [];
  let depts = {};

  // Departments + Employees
  for (let i = 1; i < empData.length; i++) {

    const dept = empData[i][4];

    if (!depts[dept]) {
      depts[dept] = 0;
    }

    depts[dept]++;
  }

}
function checkSpreadsheet() {

  const ss = SpreadsheetApp.getActive();

  Logger.log("Spreadsheet Name: " + ss.getName());

  const sheets = ss.getSheets();

  for (let i = 0; i < sheets.length; i++) {
    Logger.log(sheets[i].getName());
  }
}
