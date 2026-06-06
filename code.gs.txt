// ================================================================
// ATTENDANCE MANAGEMENT SYSTEM — Code.gs
// EMPLOYEES      → Employee_ID | Password | Role | Name
// ADMIN         → Email id | Password | Role | Name
// MAIN           → Employee_ID | Name | Date | Clock In | Clock out | Status
// LEAVE_REQUESTS → Employee_ID | Date | Reason | Status
// HOLIDAY        → Date | Holiday Name
// SALARY         → Empolyee_ID | Salary | Month
// ================================================================

var SPREADSHEET_ID = '1_5lew4gSU1tk8YaPH6qlBpouNoFa8RVo308SMz-ETb8';

var SH = {
  EMP   : 'EMPLOYEES',
  ADMIN : 'ADMIN',
  MAIN  : 'MAIN',
  LEAVE : 'LEAVE_REQUESTS',
  HOL   : 'HOLIDAY',
  SAL   : 'SALARY'
};

// ── WEB APP ENTRY ──────────────────────────────────────────────
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Attendance Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── HELPERS ────────────────────────────────────────────────────
function getSheet(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found. Check tab name is exactly: ' + name);
  return sh;
}

function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      obj[String(h).trim()] = row[i] !== undefined ? row[i] : '';
    });
    return obj;
  });
}

// ── KEY FIX: robust date formatter that handles both Date objects and strings ──
function formatDateValue(val) {
  if (!val || val === '') return '';
  // If it's already a string like "2025-04-09", return as-is
  if (typeof val === 'string') {
    // Handle "yyyy-MM-dd" or ISO string
    return val.substring(0, 10);
  }
  // If it's a Date object (Google Sheets stores dates as Date objects)
  if (val instanceof Date) {
    var tz = Session.getScriptTimeZone();
    return Utilities.formatDate(val, tz, 'yyyy-MM-dd');
  }
  return String(val).substring(0, 10);
}

function todayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function nowTimeStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');
}

// ── LOGIN ──────────────────────────────────────────────────────
// Admin tab  → checks ADMIN sheet (Email id | Password | Role | Name)
// Employee tab → checks EMPLOYEES sheet (Employee_ID | Password | Role | Name)
function login(identifier, password, mode) {
  try {
    if (mode === 'admin') {
      // ── ADMIN LOGIN via ADMIN sheet ──
      var adminRows = sheetToObjects(getSheet(SH.ADMIN));
      var admin = adminRows.find(function(r) {
        return String(r['Email id']).trim().toLowerCase() === String(identifier).trim().toLowerCase()
            && String(r['Password']).trim() === String(password).trim();
      });
      if (!admin) return { success: false, message: 'Invalid email or password.' };
      return {
        success: true,
        user: {
          id      : String(admin['Email id']).trim(),
          name    : String(admin['Name']).trim(),
          role    : 'admin',
          password: String(admin['Password']).trim()
        }
      };
    } else {
      // ── EMPLOYEE LOGIN via EMPLOYEES sheet ──
      var empRows = sheetToObjects(getSheet(SH.EMP));
      var emp = empRows.find(function(r) {
        return String(r['Employee_ID']).trim() === String(identifier).trim()
            && String(r['Password']).trim()    === String(password).trim();
      });
      if (!emp) return { success: false, message: 'Invalid Employee ID or Password.' };
      return {
        success: true,
        user: {
          id      : String(emp['Employee_ID']).trim(),
          name    : String(emp['Name']).trim(),
          role    : String(emp['Role']).toLowerCase().trim(),
          password: String(emp['Password']).trim()
        }
      };
    }
  } catch(e) { return { success: false, message: e.message }; }
}

// ── CHANGE PASSWORD ────────────────────────────────────────────
// Works for both admins (ADMIN sheet, matched by Email id) and employees (EMPLOYEES sheet)
function changePassword(identifier, currentPass, newPass, role) {
  try {
    if (role === 'admin') {
      var sh   = getSheet(SH.ADMIN);
      var data = sh.getDataRange().getValues();
      var hdrs = data[0];
      var idCol   = hdrs.findIndex(function(h){ return String(h).trim() === 'Email id'; });
      var passCol = hdrs.findIndex(function(h){ return String(h).trim() === 'Password'; });
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idCol]).trim().toLowerCase() === String(identifier).trim().toLowerCase()) {
          if (String(data[i][passCol]).trim() !== String(currentPass).trim())
            return { success: false, message: 'Current password is incorrect.' };
          sh.getRange(i + 1, passCol + 1).setValue(newPass);
          return { success: true };
        }
      }
      return { success: false, message: 'Admin account not found.' };
    } else {
      var sh   = getSheet(SH.EMP);
      var data = sh.getDataRange().getValues();
      var hdrs = data[0];
      var idCol   = hdrs.indexOf('Employee_ID');
      var passCol = hdrs.indexOf('Password');
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idCol]).trim() === String(identifier).trim()) {
          if (String(data[i][passCol]).trim() !== String(currentPass).trim())
            return { success: false, message: 'Current password is incorrect.' };
          sh.getRange(i + 1, passCol + 1).setValue(newPass);
          return { success: true };
        }
      }
      return { success: false, message: 'Employee not found.' };
    }
  } catch(e) { return { success: false, message: e.message }; }
}

// ── EMPLOYEES ──────────────────────────────────────────────────
function getEmployees() {
  try {
    return { success: true, data: sheetToObjects(getSheet(SH.EMP)) };
  } catch(e) { return { success: false, message: e.message }; }
}

function addEmployee(emp) {
  try {
    var sh = getSheet(SH.EMP);
    var rows = sheetToObjects(sh);
    if (rows.find(function(r){ return String(r['Employee_ID']).trim() === String(emp.id).trim(); }))
      return { success: false, message: 'Employee ID already exists.' };
    sh.appendRow([emp.id, emp.password, emp.role, emp.name]);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateEmployee(emp) {
  try {
    var sh = getSheet(SH.EMP);
    var data = sh.getDataRange().getValues();
    var hdrs = data[0];
    var idCol   = hdrs.indexOf('Employee_ID');
    var passCol = hdrs.indexOf('Password');
    var roleCol = hdrs.indexOf('Role');
    var nameCol = hdrs.indexOf('Name');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]).trim() === String(emp.id).trim()) {
        if (emp.password !== undefined && emp.password !== '') sh.getRange(i+1, passCol+1).setValue(emp.password);
        if (emp.role     !== undefined && emp.role     !== '') sh.getRange(i+1, roleCol+1).setValue(emp.role);
        if (emp.name     !== undefined && emp.name     !== '') sh.getRange(i+1, nameCol+1).setValue(emp.name);
        return { success: true };
      }
    }
    return { success: false, message: 'Employee not found.' };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteEmployee(empId) {
  try {
    var sh = getSheet(SH.EMP);
    var data = sh.getDataRange().getValues();
    var idCol = data[0].indexOf('Employee_ID');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]).trim() === String(empId).trim()) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'Employee not found.' };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── CLOCK IN ───────────────────────────────────────────────────
// Appends: Employee_ID | Name | Date | Clock In | Clock out | Status
function clockIn(empId, empName, location) {
  try {
    var sh = getSheet(SH.MAIN);
    var td = todayStr();
    var data = sh.getDataRange().getValues();
    var hdrs = data[0];
    var idCol = hdrs.indexOf('Employee_ID');
    var dtCol = hdrs.indexOf('Date');

    // Check if already clocked in today using robust date comparison
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]).trim() === String(empId).trim()) {
        var rowDate = formatDateValue(data[i][dtCol]);
        if (rowDate === td) {
          return { success: false, message: 'Already clocked in today.' };
        }
      }
    }

    var t = nowTimeStr();
    sh.appendRow([empId, empName, td, t, '', 'Present']);
    return { success: true, time: t };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── CLOCK OUT ──────────────────────────────────────────────────
// Updates "Clock out" column for today's MAIN row
function clockOut(empId) {
  try {
    var sh = getSheet(SH.MAIN);
    var td = todayStr();
    var data = sh.getDataRange().getValues();
    var hdrs = data[0];
    var idCol  = hdrs.indexOf('Employee_ID');
    var dtCol  = hdrs.indexOf('Date');
    // FIX: case-insensitive match so "Clock out" and "Clock Out" both work
    var outCol = hdrs.findIndex(function(h){ return String(h).trim().toLowerCase() === 'clock out'; });

    // FIX: guard against missing column header
    if (outCol < 0) return { success: false, message: 'Column "Clock out" not found in MAIN sheet. Check header spelling.' };

    // Walk rows bottom-up so we get the latest clock-in for today
    for (var i = data.length - 1; i >= 1; i--) {
      var rowDate = formatDateValue(data[i][dtCol]);
      if (String(data[i][idCol]).trim() === String(empId).trim() && rowDate === td) {
        var existing = data[i][outCol];
        if (existing && String(existing).trim() !== '') {
          return { success: false, message: 'Already clocked out today.' };
        }
        var t = nowTimeStr();
        sh.getRange(i + 1, outCol + 1).setValue(t);
        SpreadsheetApp.flush(); // Force immediate write
        return { success: true, time: t };
      }
    }
    return { success: false, message: 'No clock-in record found for today. Please clock in first.' };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── NORMALIZE TIME VALUE ───────────────────────────────────────
function formatTimeValue(val) {
  if (!val || val === '') return '';
  if (val instanceof Date) {
    var tz = Session.getScriptTimeZone();
    return Utilities.formatDate(val, tz, 'HH:mm');
  }
  var s = String(val).trim();
  if (s === '' || s === '-') return s;
  if (/^\d{1,2}:\d{2}/.test(s)) return s.substring(0, 5);
  return s;
}

// ── GET ATTENDANCE ─────────────────────────────────────────────
function getAttendance(filters) {
  try {
    var sh   = getSheet(SH.MAIN);
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    var hdrs = data[0];
    var dtCol  = hdrs.findIndex(function(h){ return String(h).trim() === 'Date'; });
    var ciCol  = hdrs.findIndex(function(h){ return String(h).trim() === 'Clock In'; });
    var coCol  = hdrs.findIndex(function(h){ return String(h).trim().toLowerCase() === 'clock out'; });
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var obj = {};
      hdrs.forEach(function(h, c) {
        obj[String(h).trim()] = (data[i][c] !== undefined && data[i][c] !== null) ? String(data[i][c]) : '';
      });
      if (dtCol >= 0) obj['Date'] = formatDateValue(data[i][dtCol]);
      if (ciCol >= 0) obj['Clock In'] = formatTimeValue(data[i][ciCol]);
      if (coCol >= 0) {
        var t = formatTimeValue(data[i][coCol]);
        obj['Clock out'] = t;
        obj['Clock Out'] = t;
      }
      result.push(obj);
    }
    if (filters) {
      if (filters.empId)
        result = result.filter(function(r){ return String(r['Employee_ID']).trim() === String(filters.empId).trim(); });
      if (filters.month)
        result = result.filter(function(r){ return String(r['Date']).substring(0, 7) === filters.month; });
      if (filters.date)
        result = result.filter(function(r){ return String(r['Date']).substring(0, 10) === filters.date; });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── GET CLOCK STATUS ───────────────────────────────────────────
function getClockStatus(empId) {
  try {
    var td   = todayStr();
    var sh   = getSheet(SH.MAIN);
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: true, clockedIn: false, clockedOut: false, clockInTime: '', clockOutTime: '' };
    var hdrs  = data[0];
    var idCol = hdrs.findIndex(function(h){ return String(h).trim() === 'Employee_ID'; });
    var dtCol = hdrs.findIndex(function(h){ return String(h).trim() === 'Date'; });
    var ciCol = hdrs.findIndex(function(h){ return String(h).trim() === 'Clock In'; });
    var coCol = hdrs.findIndex(function(h){ return String(h).trim().toLowerCase() === 'clock out'; });
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idCol]).trim() !== String(empId).trim()) continue;
      if (formatDateValue(data[i][dtCol]) !== td) continue;
      var ci  = ciCol >= 0 ? formatTimeValue(data[i][ciCol]) : '';
      var co  = coCol >= 0 ? formatTimeValue(data[i][coCol]) : '';
      return { success: true, clockedIn: ci !== '' && ci !== '-', clockedOut: co !== '' && co !== '-', clockInTime: ci, clockOutTime: co };
    }
    return { success: true, clockedIn: false, clockedOut: false, clockInTime: '', clockOutTime: '' };
  } catch(e) { return { success: false, message: e.message }; }
}


// ── LEAVE REQUESTS ─────────────────────────────────────────────
function getLeaves(empId) {
  try {
    var sh   = getSheet(SH.LEAVE);
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    var hdrs = data[0];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var obj = { _rowIndex: i };
      hdrs.forEach(function(h, c){ obj[String(h).trim()] = data[i][c]; });
      obj['Date'] = formatDateValue(data[i][hdrs.indexOf('Date')]);
      result.push(obj);
    }
    if (empId)
      result = result.filter(function(r){ return String(r['Employee_ID']).trim() === String(empId).trim(); });
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function getAllLeaves() {
  try {
    var sh   = getSheet(SH.LEAVE);
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    var hdrs = data[0];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var obj = { _rowIndex: i };
      hdrs.forEach(function(h, c){ obj[String(h).trim()] = data[i][c]; });
      obj['Date'] = formatDateValue(data[i][hdrs.indexOf('Date')]);
      result.push(obj);
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function submitLeave(empId, date, reason) {
  try {
    getSheet(SH.LEAVE).appendRow([empId, date, reason, 'Pending']);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateLeaveStatus(rowIndex, empId, empName, date, newStatus) {
  try {
    var sh        = getSheet(SH.LEAVE);
    var data      = sh.getDataRange().getValues();
    var hdrs      = data[0];
    var statusCol = hdrs.indexOf('Status');
    sh.getRange(rowIndex + 1, statusCol + 1).setValue(newStatus);

    if (newStatus === 'Approved') {
      var mainSh   = getSheet(SH.MAIN);
      var mainData = mainSh.getDataRange().getValues();
      var mHdrs    = mainData[0];
      var mIdCol   = mHdrs.indexOf('Employee_ID');
      var mDtCol   = mHdrs.indexOf('Date');
      var mStCol   = mHdrs.indexOf('Status');
      var found    = false;
      var dateStr  = String(date).substring(0, 10);
      for (var i = 1; i < mainData.length; i++) {
        var rowDate = formatDateValue(mainData[i][mDtCol]);
        if (String(mainData[i][mIdCol]).trim() === String(empId).trim() && rowDate === dateStr) {
          mainSh.getRange(i + 1, mStCol + 1).setValue('Leave');
          found = true; break;
        }
      }
      if (!found) {
        mainSh.appendRow([empId, empName || '', dateStr, '-', '-', 'Leave']);
      }
      SpreadsheetApp.flush();
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── HOLIDAYS ───────────────────────────────────────────────────
function getHolidays() {
  try {
    var sh   = getSheet(SH.HOL);
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    var hdrs = data[0];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var obj = { _rowIndex: i };
      hdrs.forEach(function(h, c){ obj[String(h).trim()] = data[i][c]; });
      obj['Date'] = formatDateValue(data[i][hdrs.indexOf('Date')]);
      result.push(obj);
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addHoliday(date, name) {
  try {
    getSheet(SH.HOL).appendRow([date, name]);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteHolidayRow(rowIndex) {
  try {
    getSheet(SH.HOL).deleteRow(rowIndex + 1);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── SALARY ─────────────────────────────────────────────────────
function getSalaries(empId) {
  try {
    var sh   = getSheet(SH.SAL);
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    var hdrs = data[0];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var obj = { _rowIndex: i };
      hdrs.forEach(function(h, c){ obj[String(h).trim()] = data[i][c]; });
      result.push(obj);
    }
    if (empId)
      result = result.filter(function(r){ return String(r['Empolyee_ID']).trim() === String(empId).trim(); });
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addSalary(empId, salary, month) {
  try {
    getSheet(SH.SAL).appendRow([empId, salary, month]);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteSalaryRow(rowIndex) {
  try {
    getSheet(SH.SAL).deleteRow(rowIndex + 1);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── DASHBOARD SUMMARY ──────────────────────────────────────────
function getDashboardSummary() {
  try {
    var td   = todayStr();
    var emps = sheetToObjects(getSheet(SH.EMP)).filter(function(u){
      return String(u['Role']).toLowerCase().trim() !== 'admin';
    });

    // Get attendance with normalised dates
    var attRes = getAttendance(null);
    var att    = attRes.success ? attRes.data : [];

    var leavSh   = getSheet(SH.LEAVE);
    var leavData = leavSh.getDataRange().getValues();
    var leavHdrs = leavData[0];
    var allLeaves = [];
    for (var i = 1; i < leavData.length; i++) {
      var obj = { _rowIndex: i };
      leavHdrs.forEach(function(h, c){ obj[String(h).trim()] = leavData[i][c]; });
      obj['Date'] = formatDateValue(leavData[i][leavHdrs.indexOf('Date')]);
      allLeaves.push(obj);
    }

    var todayAtt   = att.filter(function(a){ return String(a['Date']).substring(0,10) === td; });
    var presentTdy = todayAtt.filter(function(a){ return String(a['Status']).toLowerCase() === 'present'; }).length;
    var leaveTdy   = todayAtt.filter(function(a){ return String(a['Status']).toLowerCase() === 'leave'; }).length;
    var pending    = allLeaves.filter(function(l){ return String(l['Status']).toLowerCase() === 'pending'; });

    return {
      success        : true,
      totalEmployees : emps.length,
      presentToday   : presentTdy,
      leaveToday     : leaveTdy,
      absentToday    : Math.max(0, emps.length - presentTdy - leaveTdy),
      pendingCount   : pending.length,
      todayAttendance: todayAtt,
      pendingLeaves  : pending,
      employees      : emps
    };
  } catch(e) { return { success: false, message: e.message }; }
}
